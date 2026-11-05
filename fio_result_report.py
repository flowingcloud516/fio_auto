#!/usr/bin/env python3
"""Build an Excel performance report from fio results, optionally checking design coverage."""
from __future__ import annotations

import argparse
import csv
import datetime as dt
from html import escape
import json
from pathlib import Path
import re
import sys
import zipfile

from fio_excel_runner import CASE_COLUMNS, ConfigError, load_cases

VERSION = "1.4.0"
FINAL_EXECUTED = {"passed", "failed", "timeout", "interrupted", "error"}
BAD = {"failed", "timeout", "interrupted", "error", "preflight_failed", "skipped", "corrupt"}
STATUS_CN = {
    "passed": "已通过", "failed": "执行失败", "timeout": "超时", "interrupted": "中断",
    "error": "运行错误", "preflight_failed": "预检失败", "skipped": "已跳过",
    "planned": "仅配置验证", "pending": "等待执行", "running": "执行未完成",
    "never_run": "从未执行", "config_changed": "配置已变更", "disabled": "已禁用",
    "corrupt": "结果损坏", "collected": "已收集（执行状态未知）",
}


def case_signature(case):
    jobs = [{k: job.get(k) for k in ("job_index", "rw", "bs_bytes", "iodepth", "numjobs", "rwmixread",
                                             "numa_mem_policy")}
            for job in case.get("job_list", [])]
    return {
        "device": case.get("device", case.get("requested_device", "")), "cpu_start": case.get("cpu_start"),
        "cpu_end": case.get("cpu_end"), "time_based": case.get("time_based", True),
        "runtime_s": case.get("runtime_s"), "timeout_s": case.get("timeout_s"), "jobs": jobs,
    }


def read_json(path):
    try:
        return json.loads(path.read_text(encoding="utf-8")), ""
    except Exception as exc:
        return None, f"{path}: {exc}"


def percentile(stats, target):
    for key in ("clat_ns", "clat_us", "clat_ms", "clat"):
        section = stats.get(key)
        if not isinstance(section, dict):
            continue
        unit = {"clat_ns": 0.001, "clat_us": 1.0, "clat_ms": 1000.0, "clat": 1.0}[key]
        values = section.get("percentile") or {}
        if not values:
            return None
        wanted = min(values, key=lambda x: abs(float(x) - target))
        return float(values[wanted]) * unit
    return None


def latency_mean_us(stats):
    for key, factor in (("clat_ns", .001), ("clat_us", 1), ("clat_ms", 1000), ("clat", 1)):
        if isinstance(stats.get(key), dict) and stats[key].get("mean") is not None:
            return float(stats[key]["mean"]) * factor
    return None


def metrics_from_json(path):
    data, error = read_json(path)
    if data is None:
        return {}, error
    jobs = data.get("jobs")
    if not isinstance(jobs, list) or not jobs:
        return {}, f"{path}: missing jobs array"
    result = {"fio_version": data.get("fio version") or data.get("fio_version")}
    for direction in ("read", "write"):
        entries = [j.get(direction, {}) for j in jobs if isinstance(j, dict)]
        result[f"{direction}_iops"] = sum(float(x.get("iops", 0) or 0) for x in entries)
        result[f"{direction}_bw_mib_s"] = sum(float(x.get("bw_bytes", 0) or 0) for x in entries) / 1048576
        if not result[f"{direction}_bw_mib_s"]:
            result[f"{direction}_bw_mib_s"] = sum(float(x.get("bw", 0) or 0) for x in entries) / 1024
        result[f"{direction}_runtime_s"] = max([float(x.get("runtime", 0) or 0) for x in entries] or [0]) / 1000
        ios = [float(x.get("total_ios", 0) or 0) for x in entries]
        means = [latency_mean_us(x) for x in entries]
        valid = [(m, n) for m, n in zip(means, ios) if m is not None and n > 0]
        result[f"{direction}_clat_avg_us"] = (sum(m*n for m, n in valid) / sum(n for _, n in valid)) if valid else None
        for label, pct in (("p95", 95), ("p99", 99), ("p999", 99.9), ("p9999", 99.99)):
            values = [percentile(x, pct) for x in entries]
            result[f"{direction}_clat_{label}_us"] = max((x for x in values if x is not None), default=None)
    result["cpu_user_pct"] = sum(float(j.get("usr_cpu", 0) or 0) for j in jobs) / len(jobs)
    result["cpu_sys_pct"] = sum(float(j.get("sys_cpu", 0) or 0) for j in jobs) / len(jobs)
    return result, ""


def scaled_number(text):
    match = re.fullmatch(r"\s*([0-9.]+)\s*([kKmMgGtT]?)\s*", text)
    if not match:
        return None
    return float(match[1]) * (1000 ** (" kMGT".lower().index(match[2].lower()) if match[2] else 0))


def bandwidth_mib(value, unit):
    scale = {"B/s": 1/1048576, "KiB/s": 1/1024, "MiB/s": 1, "GiB/s": 1024,
             "kB/s": 1000/1048576, "MB/s": 1000000/1048576, "GB/s": 1000000000/1048576}
    return float(value) * scale.get(unit, 0)


def metrics_from_text(path):
    """Best-effort parser for historical fio normal output; unknown fields stay blank."""
    text = path.read_text(encoding="utf-8", errors="replace")
    result = {}
    for direction in ("read", "write"):
        starts = list(re.finditer(rf"(?mi)^\s*{direction}:\s+IOPS=([^,]+),\s+BW=([0-9.]+)([KMG]?i?B/s)", text))
        if not starts:
            continue
        start = starts[-1].start()
        other = re.search(r"(?mi)^\s*(?:read|write|trim):\s+IOPS=", text[starts[-1].end():])
        end = starts[-1].end() + other.start() if other else len(text)
        block = text[start:end]
        result[f"{direction}_iops"] = scaled_number(starts[-1].group(1))
        result[f"{direction}_bw_mib_s"] = bandwidth_mib(starts[-1].group(2), starts[-1].group(3))
        runtime = re.search(r"\(([0-9.]+)[KMG]?i?B/([0-9]+)msec\)", block)
        result[f"{direction}_runtime_s"] = float(runtime.group(2))/1000 if runtime else None
        latency = re.search(r"clat \((nsec|usec|msec)\):.*?avg=\s*([0-9.]+)", block)
        factor = {"nsec": .001, "usec": 1, "msec": 1000}
        result[f"{direction}_clat_avg_us"] = float(latency.group(2))*factor[latency.group(1)] if latency else None
        pblock = re.search(r"clat percentiles \((nsec|usec|msec)\):(.*?)(?:\n\s*lat \(|\n\s*bw \(|\n\s*cpu\s*:|\Z)", block, re.S)
        percentiles = {}
        if pblock:
            percentiles = {float(p): float(v)*factor[pblock.group(1)]
                           for p, v in re.findall(r"([0-9.]+)th=\[\s*([0-9]+)\]", pblock.group(2))}
        for label, target in (("p95", 95), ("p99", 99), ("p999", 99.9), ("p9999", 99.99)):
            key = min(percentiles, key=lambda p: abs(p-target)) if percentiles else None
            result[f"{direction}_clat_{label}_us"] = percentiles.get(key) if key is not None else None
    cpu = re.search(r"(?mi)^\s*cpu\s*:.*?usr=\s*([0-9.]+)%.*?sys=\s*([0-9.]+)%", text)
    if cpu:
        result["cpu_user_pct"], result["cpu_sys_pct"] = float(cpu.group(1)), float(cpu.group(2))
    if not any(k.endswith("_iops") for k in result):
        return {}, f"{path}: 无法识别 fio normal 性能输出"
    return result, ""


CPU_AVERAGE_FIELDS = {
    "busy_pct": "cpu_avg_busy_pct",
    "user_pct": "cpu_avg_user_pct",
    "system_pct": "cpu_avg_system_pct",
    "iowait_pct": "cpu_avg_iowait_pct",
    "idle_pct": "cpu_avg_idle_pct",
    "core_equiv": "cpu_avg_core_equiv",
}


def metrics_from_cpu_csv(path):
    """Return interval-weighted all-core averages from cpu_sample.py output."""
    totals = {output: 0.0 for output in CPU_AVERAGE_FIELDS.values()}
    weights = {output: 0.0 for output in CPU_AVERAGE_FIELDS.values()}
    sample_count = 0
    sampled_cpus = None
    try:
        with path.open(newline="", encoding="utf-8-sig") as source:
            reader = csv.DictReader(source)
            if not reader.fieldnames or "cpu" not in reader.fieldnames or "interval_s" not in reader.fieldnames:
                return {}, f"{path}: CPU CSV 缺少 cpu/interval_s 列"
            for row_number, row in enumerate(reader, 2):
                if row.get("cpu") != "selected_sum":
                    continue
                try:
                    interval = float(row.get("interval_s") or 0)
                except ValueError:
                    return {}, f"{path}:{row_number}: interval_s 不是有效数字"
                if interval <= 0:
                    continue
                used = False
                for source_key, output_key in CPU_AVERAGE_FIELDS.items():
                    raw = row.get(source_key)
                    if raw in (None, ""):
                        continue
                    try:
                        value = float(raw)
                    except ValueError:
                        return {}, f"{path}:{row_number}: {source_key} 不是有效数字"
                    totals[output_key] += value * interval
                    weights[output_key] += interval
                    used = True
                if not used:
                    continue
                sample_count += 1
                try:
                    cpus = int(row.get("selected_cpus") or 0)
                except ValueError:
                    return {}, f"{path}:{row_number}: selected_cpus 不是有效整数"
                if cpus > 0:
                    sampled_cpus = cpus if sampled_cpus is None else max(sampled_cpus, cpus)
    except Exception as exc:
        return {}, f"{path}: {exc}"
    if not sample_count:
        return {}, f"{path}: 未找到可汇总的 selected_sum 采样"
    result = {key: round(totals[key] / weights[key], 3) if weights[key] else None
              for key in totals}
    result.update(cpu_sample_count=sample_count, cpu_sampled_cpus=sampled_cpus)
    return result, ""


def cpu_metrics_for_directory(directory, configured_name=None):
    """Load a configured/adjacent CPU log; old archives without one remain valid."""
    if configured_name:
        path = Path(configured_name)
        if not path.is_absolute():
            path = directory / path
        if not path.exists():
            return {}, f"CPU 采样日志不存在: {path}", str(path)
    else:
        path = directory / "cpu.csv"
        if not path.exists():
            return {}, "", ""
    metrics, error = metrics_from_cpu_csv(path)
    return metrics, error, str(path)


def scan_archives(root):
    attempts = []
    if not root.exists():
        return attempts
    seen_result_dirs = set()
    for path in sorted(root.rglob("run_info.json")):
        info, error = read_json(path)
        batch = path.parent.parent
        if info is None:
            attempts.append({"case_id": path.parent.name, "status": "corrupt", "reason": error,
                             "batch": batch.name, "case_dir": str(path.parent)})
            continue
        row = dict(info)
        row.update(batch=batch.name, case_dir=str(path.parent), source=str(path))
        row["reason"] = info.get("error") or info.get("skip_reason") or ""
        row["sort_time"] = info.get("finished_at") or info.get("started_at") or info.get("created_at") or batch.name
        metrics, metric_error = ({}, "")
        if (path.parent / "result.json").exists():
            result_path = path.parent / "result.json"
            metrics, metric_error = metrics_from_json(result_path)
        elif (path.parent / "result.txt").exists() and info.get("status") in FINAL_EXECUTED:
            result_path = path.parent / "result.txt"
            metrics, metric_error = metrics_from_text(result_path)
        elif info.get("status") == "passed":
            result_path = None
            metric_error = "结果目录缺少 result.json/result.txt，无法提取性能指标"
        else:
            result_path = None
        row["result_path"] = str(result_path) if result_path else ""
        if result_path:
            seen_result_dirs.add(path.parent.resolve())
        cpu_metrics, cpu_error, cpu_log_path = cpu_metrics_for_directory(path.parent, info.get("cpu_log"))
        row.update(metrics)
        row.update(cpu_metrics)
        row["cpu_log_path"] = cpu_log_path
        row["metric_error"] = "; ".join(note for note in (metric_error, cpu_error) if note)
        attempts.append(row)
    if root.is_file():
        raw_results = [root] if root.suffix.lower() in {".txt", ".json"} else []
    else:
        raw_results = sorted(list(root.rglob("result.json")) + list(root.rglob("result.txt")))
    # Prefer result.json when both formats exist in one directory.
    chosen = {}
    for path in raw_results:
        chosen.setdefault(path.parent.resolve(), path)
        if path.name == "result.json":
            chosen[path.parent.resolve()] = path
    for directory, path in chosen.items():
        if directory in seen_result_dirs:
            continue
        metrics, metric_error = (metrics_from_json(path) if path.suffix.lower() == ".json"
                                 else metrics_from_text(path))
        case_id = re.sub(r"^\d+_", "", path.parent.name) if root.is_dir() else path.stem
        modified = dt.datetime.fromtimestamp(path.stat().st_mtime, dt.timezone.utc).isoformat(timespec="seconds")
        row = dict(case_id=case_id or path.stem, status="collected", reason="未提供 run_info.json，执行状态未知",
                   batch=path.parent.parent.name if root.is_dir() else "standalone",
                   case_dir=str(path.parent), source=str(path), result_path=str(path),
                   created_at=modified, finished_at=modified, sort_time=modified,
                   metric_error=metric_error)
        cpu_metrics, cpu_error, cpu_log_path = cpu_metrics_for_directory(path.parent)
        row.update(metrics)
        row.update(cpu_metrics)
        row["cpu_log_path"] = cpu_log_path
        row["metric_error"] = "; ".join(note for note in (metric_error, cpu_error) if note)
        attempts.append(row)
    return attempts


def build_rows(cases, attempts):
    by_case = {}
    for attempt in attempts:
        by_case.setdefault(attempt.get("case_id"), []).append(attempt)
    overview, details, exceptions, job_rows = [], [], [], []
    design = {c["case_id"]: c for c in cases}
    for attempt in attempts:
        current = design.get(attempt.get("case_id"))
        match = bool(current) and case_signature(current) == case_signature(attempt)
        status = attempt.get("status", "corrupt")
        details.append(detail_row(attempt, match))
    for case in cases:
        history = sorted(by_case.get(case["case_id"], []), key=lambda x: x.get("sort_time", ""))
        matching = [x for x in history if case_signature(x) == case_signature(case)]
        if not case["enabled"]:
            status, latest, reason = "disabled", None, "Excel enabled=0"
        elif matching:
            latest = matching[-1]
            status, reason = latest.get("status", "corrupt"), latest.get("reason", "")
        elif history:
            latest, status, reason = history[-1], "config_changed", "存在同 case_id 历史结果，但当前参数已变化"
        else:
            latest, status, reason = None, "never_run", "未找到该 case_id 的归档"
        executed = [x for x in matching if x.get("status") in FINAL_EXECUTED]
        passed = [x for x in matching if x.get("status") == "passed"]
        row = overview_row(case, latest, status, reason, len(executed), len(passed), len(history))
        overview.append(row)
        for job in case["job_list"]:
            job_rows.append([case["case_id"], job["job_index"], job["rw"], job["bs_bytes"],
                             job["iodepth"], job["numjobs"], job["rwmixread"],
                             job.get("numa_mem_policy") or "", case["cpu_start"], case["cpu_end"],
                             STATUS_CN.get(status, status)])
        if status not in {"passed", "disabled"} or (latest and latest.get("metric_error")):
            exceptions.append([case["case_id"], STATUS_CN.get(status, status), reason,
                               latest.get("metric_error", "") if latest else "",
                               latest.get("batch", "") if latest else ""])
    return overview, details, exceptions, job_rows


def build_result_only_rows(attempts):
    details = [detail_row(attempt, None) for attempt in attempts]
    exceptions = []
    for attempt in attempts:
        status = attempt.get("status", "corrupt")
        if status in BAD or attempt.get("metric_error"):
            exceptions.append([attempt.get("case_id", ""), STATUS_CN.get(status, status),
                               attempt.get("reason", ""), attempt.get("metric_error", ""),
                               attempt.get("batch", "")])
    return details, exceptions


METRIC_KEYS = [
    "read_iops", "write_iops", "read_bw_mib_s", "write_bw_mib_s",
    "read_clat_avg_us", "read_clat_p95_us", "read_clat_p99_us", "read_clat_p999_us", "read_clat_p9999_us",
    "write_clat_avg_us", "write_clat_p95_us", "write_clat_p99_us", "write_clat_p999_us", "write_clat_p9999_us",
    "read_runtime_s", "write_runtime_s", "cpu_user_pct", "cpu_sys_pct",
]

CPU_METRIC_KEYS = [
    "cpu_sample_count", "cpu_sampled_cpus", "cpu_avg_busy_pct", "cpu_avg_user_pct",
    "cpu_avg_system_pct", "cpu_avg_iowait_pct", "cpu_avg_idle_pct", "cpu_avg_core_equiv",
]


def overview_row(case, latest, status, reason, run_count, pass_count, history_count):
    source = case.get("source_values", {})
    def compact(key, formatter=lambda value: value):
        values = {job.get(key) for job in case["job_list"]}
        if len(values) == 1:
            value = next(iter(values))
            return "" if value is None else formatter(value)
        return "多值（见Job配置）"

    def readable_bs(value):
        for suffix, unit in (("G", 1024**3), ("M", 1024**2), ("K", 1024)):
            if value % unit == 0:
                return f"{value // unit}{suffix}"
        return value

    config = [case["case_id"], 1 if case["enabled"] else 0, case.get("requested_device", ""),
              compact("rw"), compact("bs_bytes", readable_bs), compact("iodepth"), case["jobs"],
              compact("numjobs"), case["cpu_start"], case["cpu_end"], case.get("runtime_s"),
              compact("rwmixread"), source.get("notes", ""),
              1 if case.get("time_based", True) else 0, case.get("timeout_s"),
              compact("numa_mem_policy")]
    values = config + [STATUS_CN.get(status, status), reason, run_count, pass_count, history_count,
                       latest.get("batch", "") if latest else "",
                       latest.get("finished_at", "") if latest else ""]
    values += [latest.get(k) if latest else None for k in CPU_METRIC_KEYS]
    values += [latest.get(k) if latest else None for k in METRIC_KEYS]
    return values


def detail_row(a, match):
    values = [a.get("case_id"), a.get("batch"), STATUS_CN.get(a.get("status"), a.get("status")),
              "不适用" if match is None else ("是" if match else "否"),
              a.get("created_at", ""), a.get("started_at", ""),
              a.get("finished_at", ""), a.get("exit_code"), a.get("reason", ""),
              a.get("metric_error", ""), a.get("case_dir", ""), a.get("result_path", ""),
              a.get("cpu_log_path", "")]
    values += [a.get(k) for k in CPU_METRIC_KEYS]
    values += [a.get(k) for k in METRIC_KEYS]
    return values


OVERVIEW_HEADERS = list(CASE_COLUMNS) + ["当前状态", "状态说明", "执行次数", "成功次数", "同ID历史数",
                                               "最新批次", "完成时间"] + CPU_METRIC_KEYS + METRIC_KEYS
DETAIL_HEADERS = ["用例ID", "批次", "状态", "匹配当前配置", "创建时间", "开始时间", "完成时间", "退出码", "原因", "指标说明", "用例目录", "原始结果日志", "CPU采样日志"] + CPU_METRIC_KEYS + METRIC_KEYS
EXCEPTION_HEADERS = ["用例ID", "状态", "原因", "指标说明", "批次"]
JOB_HEADERS = ["case_id", "job_index", "rw", "bs_bytes", "iodepth", "numjobs", "rwmixread",
               "numa_mem_policy", "cpu_start", "cpu_end", "当前状态"]


def col_name(number):
    text = ""
    while number:
        number, rest = divmod(number - 1, 26)
        text = chr(65 + rest) + text
    return text


def cell_xml(ref, value, style=0):
    if value is None or value == "":
        return f'<c r="{ref}" s="{style}"/>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'


def sheet_xml(title, headers, rows):
    width = []
    for i, header in enumerate(headers):
        longest = max([len(str(header))] + [len(str(r[i])) for r in rows if i < len(r) and r[i] is not None])
        width.append(min(max(longest + 2, 10), 42))
    row_xml = []
    all_rows = [headers] + rows
    for rindex, row in enumerate(all_rows, 1):
        cells = "".join(cell_xml(f"{col_name(cindex)}{rindex}", value, 1 if rindex == 1 else 0)
                        for cindex, value in enumerate(row, 1))
        row_xml.append(f'<row r="{rindex}">{cells}</row>')
    cols = "".join(f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>' for i, w in enumerate(width, 1))
    end = f"{col_name(len(headers))}{len(all_rows)}"
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
            f'<cols>{cols}</cols><sheetData>{"".join(row_xml)}</sheetData>'
            f'<autoFilter ref="A1:{end}"/></worksheet>')


def write_xlsx(path, sheets):
    path.parent.mkdir(parents=True, exist_ok=True)
    content_types = ['<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                     '<Default Extension="xml" ContentType="application/xml"/>',
                     '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
                     '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    content_types += [f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets)+1)]
    workbook_sheets = "".join(f'<sheet name="{escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i, (name, _, _) in enumerate(sheets, 1))
    rels = "".join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets)+1))
    rels += f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    styles = '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'+"".join(content_types)+"</Types>")
        z.writestr("_rels/.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+workbook_sheets+'</sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+rels+'</Relationships>')
        z.writestr("xl/styles.xml", styles)
        for i, (name, headers, rows) in enumerate(sheets, 1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", sheet_xml(name, headers, rows))


def arguments(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, help="optional current design workbook for coverage checks")
    parser.add_argument("--results", required=True, type=Path,
                        help="fio_results archive root, a directory of raw results, or one result.txt/result.json")
    parser.add_argument("--output", type=Path, default=Path("统计结果.xlsx"))
    return parser.parse_args(argv)


def main(argv=None):
    if sys.version_info < (3, 11):
        raise ConfigError("Python 3.11 or newer is required")
    args = arguments(argv)
    attempts = scan_archives(args.results)
    if args.excel:
        cases = load_cases(args.excel)
        overview, details, exceptions, job_rows = build_rows(cases, attempts)
        summary = [["生成时间", dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")],
                   ["模式", "用例覆盖对照"], ["设计用例数", len(cases)],
                   ["启用用例数", sum(c["enabled"] for c in cases)],
                   ["扫描执行记录数", len(attempts)], ["异常/待执行用例数", len(exceptions)],
                   ["设计文件", str(args.excel.resolve())], ["结果目录", str(args.results.resolve())]]
        sheets = [("统计概览", ["项目", "值"], summary),
                  ("用例覆盖", OVERVIEW_HEADERS, overview), ("Job配置", JOB_HEADERS, job_rows),
                  ("运行明细", DETAIL_HEADERS, details), ("异常清单", EXCEPTION_HEADERS, exceptions)]
    else:
        details, exceptions = build_result_only_rows(attempts)
        summary = [["生成时间", dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")],
                   ["模式", "仅结果收集（未提供用例Excel）"], ["扫描结果数", len(attempts)],
                   ["无法解析/异常数", len(exceptions)], ["结果来源", str(args.results.resolve())]]
        sheets = [("统计概览", ["项目", "值"], summary),
                  ("运行明细", DETAIL_HEADERS, details), ("异常清单", EXCEPTION_HEADERS, exceptions)]
    write_xlsx(args.output, sheets)
    print(f"Report written: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
