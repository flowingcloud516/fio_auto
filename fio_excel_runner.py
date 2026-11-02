#!/usr/bin/env python3
"""Excel-driven Linux fio runner. Python 3.11+, standard library only."""
from __future__ import annotations

import argparse
import array
from contextlib import ExitStack
import csv
import datetime as dt
import json
import logging
import math
import os
from pathlib import Path
import posixpath
import re
import shutil
import shlex
import signal
import stat
import subprocess
import sys
import time
import uuid
import xml.etree.ElementTree as ET
import zipfile

VERSION = "1.8.0"
DISK_RANGE_PERCENT = 84
CASE_BASE_COLUMNS = ("case_id", "enabled", "device", "rw", "bs", "iodepth", "jobs", "numjobs",
                     "cpu_start", "cpu_end", "runtime_s", "rwmixread", "notes")
CASE_OPTIONAL_COLUMNS = ("time_based", "timeout_s", "numa_mem_policy")
CASE_COLUMNS = CASE_BASE_COLUMNS + CASE_OPTIONAL_COLUMNS
JOB_BASE_COLUMNS = ("case_id", "job_index", "rw", "bs", "iodepth", "numjobs", "rwmixread", "notes")
JOB_OPTIONAL_COLUMNS = ("numa_mem_policy",)
JOB_COLUMNS = JOB_BASE_COLUMNS + JOB_OPTIONAL_COLUMNS
RW = {"read", "write", "randread", "randwrite", "rw", "randrw"}
FIXED = ("direct=1", "ioengine=libaio", "group_reporting=1", "norandommap=1",
         "randrepeat=0", "ramp_time=0")


class ConfigError(ValueError):
    pass


def fail(message):
    raise ConfigError(message)


def integer(value, label, minimum=1, default=None, maximum=None):
    if value in (None, ""):
        if default is not None:
            return default
        fail(f"{label}: required")
    if isinstance(value, bool) or not re.fullmatch(r"[+-]?\d+(?:\.0+)?", str(value).strip()):
        fail(f"{label}: expected an integer, got {value!r}")
    number = int(str(value).strip().split(".")[0])
    if number < minimum or (maximum is not None and number > maximum):
        fail(f"{label}: out of range ({minimum} .. {maximum or 'unlimited'})")
    return number


def block_size(value, label):
    match = re.fullmatch(r"([1-9]\d*)\s*([kKmMgG]?)(?:[iI]?[bB])?", str(value).strip())
    if not match:
        fail(f"{label}: use a single block size, e.g. 4k, 128k, 1m or 4096")
    return int(match[1]) * 1024 ** (" kmg".index(match[2].lower()) if match[2] else 0)


def numa_mem_policy(value, label):
    """Validate fio's numa_mem_policy syntax and return a normalized value."""
    if value in (None, ""):
        return None
    text = re.sub(r"\s+", "", str(value).lower())
    mode, separator, nodes = text.partition(":")
    if mode not in {"default", "prefer", "bind", "interleave", "local"}:
        fail(f"{label}: mode must be default, prefer, bind, interleave or local")
    if mode in {"default", "local"}:
        if separator:
            fail(f"{label}: {mode} does not accept a node list")
        return mode
    if not separator or not nodes:
        fail(f"{label}: {mode} requires a node list, e.g. {mode}:0")
    if nodes != "all" and not re.fullmatch(r"\d+(?:-\d+)?(?:,\d+(?:-\d+)?)*", nodes):
        fail(f"{label}: invalid NUMA node list {nodes!r}")
    if mode == "prefer" and (nodes == "all" or "," in nodes or "-" in nodes):
        fail(f"{label}: prefer accepts exactly one NUMA node")
    if nodes != "all":
        for part in nodes.split(","):
            ends = [int(item) for item in part.split("-")]
            if len(ends) == 2 and ends[0] > ends[1]:
                fail(f"{label}: invalid descending NUMA node range {part}")
    return f"{mode}:{nodes}"


def read_xlsx(path):
    """Read ordinary .xlsx literal cells, including Excel shared and inline strings.

    Deliberately reject formulas in input sheets: no dependency on cached values
    or an Excel calculation engine. Styles, tables and blank styled rows are OK.
    """
    if path.suffix.lower() != ".xlsx":
        fail("Input must be an unencrypted .xlsx file (not .xls/.xlsm/CSV)")
    try:
        with zipfile.ZipFile(path) as archive:
            if sum(x.file_size for x in archive.infolist()) > 64 * 1024 * 1024:
                fail("Workbook exceeds the 64 MiB uncompressed input limit")
            workbook = ET.fromstring(archive.read("xl/workbook.xml"))
            rels = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
            targets = {r.attrib["Id"]: r.attrib["Target"] for r in rels
                       if r.attrib.get("TargetMode") != "External"}
            strings = []
            if "xl/sharedStrings.xml" in archive.namelist():
                shared = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                strings = ["".join(t.text or "" for t in s.findall(".//{*}t")) for s in shared]
            output = {}
            for sheet in workbook.findall(".//{*}sheet"):
                name = sheet.attrib["name"]
                if name not in ("Cases", "JobOverrides"):
                    continue
                relation = next(v for k, v in sheet.attrib.items() if k.endswith("}id"))
                target = targets[relation]
                member = (target.lstrip("/") if target.startswith("/")
                          else posixpath.normpath(posixpath.join("xl", target)))
                root = ET.fromstring(archive.read(member))
                cells = {}
                for cell in root.findall(".//{*}sheetData/{*}row/{*}c"):
                    coordinate = cell.attrib.get("r", "")
                    if cell.find("{*}f") is not None:
                        fail(f"{name}!{coordinate}: formulas are not supported; paste values")
                    typ = cell.attrib.get("t", "n")
                    raw = cell.findtext("{*}v", default="")
                    if typ == "s":
                        value = strings[int(raw)]
                    elif typ == "inlineStr":
                        value = "".join(t.text or "" for t in cell.findall("{*}is/.//{*}t"))
                    elif typ == "b":
                        value = "1" if raw == "1" else "0"
                    elif typ in ("e", "d"):
                        fail(f"{name}!{coordinate}: error/date cells are not valid inputs")
                    else:
                        value = raw
                    if value != "":
                        cells[coordinate] = value.strip() if isinstance(value, str) else value
                output[name] = cells
            return output
    except (zipfile.BadZipFile, KeyError, ET.ParseError, IndexError, StopIteration) as exc:
        fail(f"Invalid or unsupported .xlsx: {exc}")


def sheet_records(cells, columns, name, optional=()):
    rows = {}
    for coordinate, value in cells.items():
        match = re.fullmatch(r"([A-Z]+)([1-9]\d*)", coordinate)
        if not match:
            fail(f"{name}: invalid cell reference {coordinate!r}")
        col = 0
        for char in match[1]:
            col = col * 26 + ord(char) - 64
        rows.setdefault(int(match[2]), {})[col] = value
    headers = rows.pop(1, {})
    values = list(headers.values())
    if len(values) != len(set(values)):
        fail(f"{name}: row 1 contains duplicate headers")
    missing = set(columns) - set(values)
    unknown = set(values) - set(columns) - set(optional)
    if missing or unknown:
        fail(f"{name}: invalid headers; missing={sorted(missing)}, unknown={sorted(unknown)}")
    output_columns = tuple(columns) + tuple(optional)
    for row_number, data in sorted(rows.items()):
        if any(col not in headers for col in data):
            fail(f"{name} row {row_number}: data outside the declared columns")
        record = {header: "" for header in output_columns}
        record.update({header: data.get(col, "") for col, header in headers.items()})
        yield row_number, record


def effective_job(data, label):
    rw = str(data["rw"]).lower()
    if rw not in RW:
        fail(f"{label}.rw: expected one of {sorted(RW)}")
    mix = data.get("rwmixread", "")
    if rw not in {"rw", "randrw"} and mix not in (None, ""):
        fail(f"{label}.rwmixread: only valid for rw/randrw")
    return dict(rw=rw, bs_bytes=block_size(data["bs"], f"{label}.bs"),
                iodepth=integer(data["iodepth"], f"{label}.iodepth"),
                numjobs=integer(data["numjobs"], f"{label}.numjobs", maximum=65536),
                rwmixread=integer(mix, f"{label}.rwmixread", minimum=0, maximum=100, default=50)
                if rw in {"rw", "randrw"} else None,
                numa_mem_policy=numa_mem_policy(data.get("numa_mem_policy", ""),
                                                 f"{label}.numa_mem_policy"))


def load_cases(path):
    sheets = read_xlsx(path)
    if "Cases" not in sheets:
        fail("Missing Cases worksheet")
    cases = []
    ids = set()
    for row, values in sheet_records(sheets["Cases"], CASE_BASE_COLUMNS, "Cases", CASE_OPTIONAL_COLUMNS):
        label = f"Cases row {row}"
        case_id = values["case_id"]
        if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", case_id):
            fail(f"{label}.case_id: use 1-64 ASCII letters, digits, '_' or '-'")
        if case_id in ids:
            fail(f"{label}: duplicate case_id {case_id}")
        ids.add(case_id)
        enabled = integer(values["enabled"], f"{label}.enabled", minimum=0, maximum=1, default=1)
        count = integer(values["jobs"], f"{label}.jobs", maximum=65536)
        start = integer(values["cpu_start"], f"{label}.cpu_start", minimum=0, maximum=65535)
        end = integer(values["cpu_end"], f"{label}.cpu_end", minimum=start, maximum=65535)
        effective_job(values, label)  # Validate defaults, even for disabled examples.
        time_based = bool(integer(values["time_based"], f"{label}.time_based", minimum=0, maximum=1, default=1))
        if time_based:
            runtime_s = integer(values["runtime_s"], f"{label}.runtime_s", default=60)
            timeout_s = (integer(values["timeout_s"], f"{label}.timeout_s")
                         if values["timeout_s"] not in (None, "") else None)
        else:
            if values["runtime_s"] not in (None, ""):
                fail(f"{label}.runtime_s: must be blank when time_based=0")
            runtime_s = None
            timeout_s = integer(values["timeout_s"], f"{label}.timeout_s")
        cases.append(dict(case_id=case_id, enabled=bool(enabled), jobs=count,
                          requested_device=values["device"],
                          cpu_start=start, cpu_end=end,
                          time_based=time_based, runtime_s=runtime_s, timeout_s=timeout_s,
                          source_values=dict(values), defaults=values, overrides={}))
    by_id = {c["case_id"]: c for c in cases}
    if "JobOverrides" in sheets:
        for row, values in sheet_records(sheets["JobOverrides"], JOB_BASE_COLUMNS, "JobOverrides",
                                         JOB_OPTIONAL_COLUMNS):
            label = f"JobOverrides row {row}"
            case = by_id.get(values["case_id"])
            if case is None:
                fail(f"{label}: unknown case_id {values['case_id']!r}")
            index = integer(values["job_index"], f"{label}.job_index", maximum=case["jobs"])
            if index in case["overrides"]:
                fail(f"{label}: duplicate override for job {index}")
            case["overrides"][index] = {k: v for k, v in values.items()
                                        if k in ("rw", "bs", "iodepth", "numjobs", "rwmixread",
                                                 "numa_mem_policy") and v != ""}
    for case in cases:
        job_list = []
        for index in range(1, case["jobs"] + 1):
            merged = dict(case["defaults"])
            merged.update(case["overrides"].get(index, {}))
            # A non-mixed override must not inherit a mixed default's ratio.
            if merged["rw"].lower() not in {"rw", "randrw"}:
                merged["rwmixread"] = case["overrides"].get(index, {}).get("rwmixread", "")
            job_list.append(dict(job_index=index, **effective_job(merged, f"{case['case_id']}.job_{index:03}")))
        case["job_list"] = job_list
        case["total_workers"] = sum(j["numjobs"] for j in job_list)
        if case["total_workers"] > 65536:
            fail(f"{case['case_id']}: more than 65536 total workers")
        del case["defaults"], case["overrides"]
    if not cases:
        fail("Cases worksheet has no data")
    return cases


def allocate_job_cpus(start, end, counts):
    """One contiguous interval of exactly numjobs CPUs per independent job.

    With spare CPUs, spread the intervals using gaps. With oversubscription,
    reuse the least-loaded contiguous window; never wrap one interval.
    """
    available = end - start + 1
    if not counts or any(n < 1 or n > available for n in counts):
        fail(f"Each job's numjobs must be within 1..{available} for CPU range {start}-{end}")
    result = []
    total = sum(counts)
    if total <= available:
        spare = available - total
        used = 0
        for index, count in enumerate(counts):
            gap = index * spare // (len(counts)-1) if len(counts) > 1 else 0
            first = start + used + gap
            result.append(list(range(first, first+count)))
            used += count
    else:
        loads = [0] * available
        for count in counts:
            current = best_load = sum(loads[:count])
            best_start = 0
            for first in range(1, available-count+1):
                current += loads[first+count-1] - loads[first-1]
                if current < best_load:
                    best_load, best_start = current, first
            result.append(list(range(start+best_start, start+best_start+count)))
            for cpu in range(best_start, best_start+count):
                loads[cpu] += 1
    return result


def validate_device_name(name):
    if not re.fullmatch(r"/dev/[A-Za-z0-9_./-]+", name) or ".." in Path(name).parts:
        fail("device must be a /dev/... path without spaces or fio special characters")


def resolve_case_devices(cases, override=None):
    """CLI overrides every selected row; otherwise use that row's Excel device."""
    for case in cases:
        name = override if override is not None else case["requested_device"]
        if not name:
            fail(f"{case['case_id']}: missing device; fill Cases.device or pass --device")
        validate_device_name(name)
        case["effective_device"] = name
        case["device_source"] = "command_line" if override is not None else "excel"


def make_plan(case, device, capacity, sector):
    alignment = math.lcm(sector, *(j["bs_bytes"] for j in case["job_list"]))
    for job in case["job_list"]:
        if job["bs_bytes"] % sector:
            fail(f"{case['case_id']}: block size must be a multiple of logical sector {sector}")
    limit = capacity * DISK_RANGE_PERCENT // 100
    size = limit // alignment * alignment
    job_count = len(case["job_list"])
    units = size // alignment
    if units < job_count:
        fail(f"{case['case_id']}: 84% region is too small to give every independent job an aligned region")
    # For multiple jobs, emit percentages only if no rounding at all is needed.
    # Percent offset rounds UP in fio; our general byte boundaries round DOWN.
    exact_percent = (DISK_RANGE_PERCENT % job_count == 0 and
                     capacity * DISK_RANGE_PERCENT % (100 * job_count * alignment) == 0)
    job_cpus = allocate_job_cpus(case["cpu_start"], case["cpu_end"],
                                [j["numjobs"] for j in case["job_list"]])
    cpus = sorted({cpu for group in job_cpus for cpu in group})
    lines = ["; Generated by fio_excel_runner. Independent jobs use disjoint regions.",
             f"; Total disk range: 0%-{DISK_RANGE_PERCENT}%; capacity={capacity} bytes.",
             "; Clones created by numjobs share their parent job's region.",
             "[global]", f"filename={device}", *FIXED, "allow_file_create=0", "exitall_on_error=1"]
    if case["time_based"]:
        lines += ["time_based=1", f"runtime={case['runtime_s']}"]
    lines.append("")
    mapping = []
    regions = []
    cpu_assignments = []
    worker = 0
    for index, job in enumerate(case["job_list"]):
        name = f"job_{job['job_index']:03}"
        assigned_cpus = job_cpus[index]
        cpu_arg = (str(assigned_cpus[0]) if len(assigned_cpus) == 1
                   else f"{assigned_cpus[0]}-{assigned_cpus[-1]}")
        cpu_assignments.append(dict(job=name, numjobs=job["numjobs"],
                                    cpu_start=assigned_cpus[0], cpu_end=assigned_cpus[-1],
                                    cpus_allowed=cpu_arg))
        start = (units * index // job_count) * alignment
        end = (units * (index + 1) // job_count) * alignment
        length = end - start
        nominal_start = f"{DISK_RANGE_PERCENT * index / job_count:.6f}".rstrip("0").rstrip(".")
        nominal_end = f"{DISK_RANGE_PERCENT * (index + 1) / job_count:.6f}".rstrip("0").rstrip(".")
        offset_arg = f"{index * (DISK_RANGE_PERCENT // job_count)}%" if exact_percent and index else str(start)
        size_arg = f"{DISK_RANGE_PERCENT // job_count}%" if exact_percent or job_count == 1 else str(length)
        regions.append(dict(job=name, offset_bytes=start, size_bytes=length,
                            end_exclusive_bytes=end, nominal_range=f"[{nominal_start}%, {nominal_end}%)",
                            fio_offset=offset_arg, fio_size=size_arg))
        lines += [f"[{name}]",
                  f"; Nominal disk range: [{nominal_start}%, {nominal_end}%). End is exclusive.",
                  f"; Aligned byte range: [{start}, {end}); size={length} bytes.",
                  f"offset={offset_arg}", f"size={size_arg}",
                  f"rw={job['rw']}", f"bs={job['bs_bytes']}",
                  f"iodepth={job['iodepth']}", f"numjobs={job['numjobs']}",
                  f"cpus_allowed={cpu_arg}", "cpus_allowed_policy=split"]
        if job["rwmixread"] is not None:
            lines.append(f"rwmixread={job['rwmixread']}")
        if job.get("numa_mem_policy") is not None:
            lines.append(f"numa_mem_policy={job['numa_mem_policy']}")
        lines.append("")
        for clone in range(job["numjobs"]):
            # fio uses (global thread_number - 1) % cpus_in_mask for split.
            mapping.append(dict(job=name, clone_index=clone, global_worker_index=worker,
                                expected_cpu=assigned_cpus[worker % len(assigned_cpus)]))
            worker += 1
    return dict(**case, device=device, capacity_bytes=capacity, logical_sector_bytes=sector,
                limit_84_percent_bytes=limit, region_end_exclusive=size,
                range_policy="independent_jobs_disjoint_clones_shared", job_regions=regions,
                cpu_assignments=cpu_assignments,
                region_alignment_bytes=alignment, selected_cpus=cpus, workers=mapping,
                writes=any(j["rw"] not in {"read", "randread"} for j in case["job_list"]),
                fio_text="\n".join(lines))


def probe_device(name):
    import fcntl
    resolved = Path(name).resolve(strict=True)
    validate_device_name(str(resolved))
    info = resolved.stat()
    if not stat.S_ISBLK(info.st_mode):
        fail(f"{name}: must be a Linux block device")
    with open(resolved, "rb", buffering=0) as handle:
        capacity = array.array("Q", [0])
        sector = array.array("I", [0])
        fcntl.ioctl(handle.fileno(), 0x80081272, capacity, True)  # BLKGETSIZE64
        fcntl.ioctl(handle.fileno(), 0x1268, sector, True)  # BLKSSZGET
    return dict(path=str(resolved), major=os.major(info.st_rdev), minor=os.minor(info.st_rdev),
                capacity_bytes=capacity[0], logical_sector_bytes=sector[0])


def check_device_idle(device):
    """Reject mounted, swap-backed, held or host-managed zoned targets."""
    node = Path(f"/sys/dev/block/{device['major']}:{device['minor']}").resolve(strict=True)
    nodes = [node] + [p for p in node.iterdir() if (p / "partition").is_file()]
    ids = {(p / "dev").read_text().strip() for p in nodes}
    for p in nodes:
        if list((p / "holders").iterdir()):
            fail(f"{p.name}: in use by a device-mapper/RAID holder")
    for line in Path("/proc/self/mountinfo").read_text().splitlines():
        if line.split()[2] in ids:
            fail("Target or one of its partitions is mounted; refusing to run")
    for line in Path("/proc/swaps").read_text().splitlines()[1:]:
        path = line.split()[0]
        try:
            info = os.stat(path)
            number = info.st_rdev if stat.S_ISBLK(info.st_mode) else info.st_dev
            if f"{os.major(number)}:{os.minor(number)}" in ids:
                fail("Target is used by active swap")
        except FileNotFoundError:
            fail(f"Cannot verify active swap path {path}")
    queue = node / "queue"
    if not queue.exists():
        queue = node.parent / "queue"
    if (queue / "zoned").exists() and (queue / "zoned").read_text().strip() == "host-managed":
        fail("Host-managed zoned devices need a separate zone-aware test configuration")


def check_cpus(cases):
    allowed = os.sched_getaffinity(0)
    online = set()
    for part in Path("/sys/devices/system/cpu/online").read_text().strip().split(","):
        ends = [int(x) for x in part.split("-")]
        online.update(range(ends[0], ends[-1] + 1))
    for case in cases:
        requested = set(range(case["cpu_start"], case["cpu_end"] + 1))
        missing = requested - (online & allowed)
        if missing:
            fail(f"{case['case_id']}: requested CPU(s) offline or outside current affinity/cpuset: {sorted(missing)}")


def expand_id_list(text, label):
    values = set()
    for part in text.strip().split(","):
        if not part:
            fail(f"{label}: empty list element")
        ends = [int(item) for item in part.split("-")]
        if len(ends) > 2 or (len(ends) == 2 and ends[0] > ends[1]):
            fail(f"{label}: invalid range {part!r}")
        values.update(range(ends[0], ends[-1] + 1))
    return values


def check_numa_nodes(cases, online_path=Path("/sys/devices/system/node/online")):
    policies = [(case["case_id"], job["job_index"], job.get("numa_mem_policy"))
                for case in cases for job in case["job_list"] if job.get("numa_mem_policy")]
    if not policies:
        return
    if not online_path.is_file():
        fail("NUMA memory policy requested, but /sys/devices/system/node/online is unavailable")
    online = expand_id_list(online_path.read_text().strip(), "online NUMA nodes")
    for case_id, job_index, policy in policies:
        mode, separator, node_list = policy.partition(":")
        if not separator or node_list == "all":
            continue
        requested = expand_id_list(node_list, f"{case_id}.job_{job_index:03}.numa_mem_policy")
        missing = requested - online
        if missing:
            fail(f"{case_id}.job_{job_index:03}: NUMA node(s) not online: {sorted(missing)}")


def fio_supports_numa(fio):
    try:
        result = subprocess.run([fio, "--cmdhelp=numa_mem_policy"], capture_output=True,
                                text=True, timeout=10, check=False)
    except (OSError, subprocess.TimeoutExpired) as exc:
        fail(f"Cannot query fio NUMA support: {exc}")
    output = (result.stdout or "") + (result.stderr or "")
    return result.returncode == 0 and "numa_mem_policy" in output


def check_fio_numa_support(fio, cases):
    if (any(job.get("numa_mem_policy") for case in cases for job in case["job_list"])
            and not fio_supports_numa(fio)):
        fail("This fio build does not support numa_mem_policy. The syntax bind:<node> is valid, "
             "but fio must be built with libnuma development headers (libnuma-dev or "
             "numactl-devel). Verify with: fio --cmdhelp=numa_mem_policy")


def numactl_policy_args(policy):
    """Translate one normalized fio NUMA policy to numactl command arguments."""
    if policy in (None, "default"):
        return []
    mode, _, nodes = policy.partition(":")
    if mode == "bind":
        return [f"--membind={nodes}"]
    if mode == "prefer":
        return [f"--preferred={nodes}"]
    if mode == "interleave":
        return [f"--interleave={nodes}"]
    if mode == "local":
        return ["--localalloc"]
    fail(f"Unsupported NUMA memory policy: {policy}")


def configure_numa_backend(plans, fio, requested="auto", numactl="numactl"):
    """Choose native fio or an external numactl wrapper for every case."""
    has_policy = any(job.get("numa_mem_policy") for plan in plans for job in plan["job_list"])
    native = fio_supports_numa(fio) if has_policy and requested in {"auto", "fio"} else False
    if requested == "fio" and has_policy and not native:
        check_fio_numa_support(fio, plans)
    use_numactl = has_policy and (requested == "numactl" or (requested == "auto" and not native))
    executable = shutil.which(numactl) if use_numactl else None
    if use_numactl and not executable:
        fail(f"numactl executable not found: {numactl!r}")
    for plan in plans:
        policies = [job.get("numa_mem_policy") for job in plan["job_list"]]
        if not any(policies):
            plan.update(numa_backend="none", numa_command_prefix=[])
            continue
        if not use_numactl:
            plan.update(numa_backend="fio", numa_command_prefix=[])
            continue
        effective = [None if policy == "default" else policy for policy in policies]
        if len(set(effective)) != 1:
            fail(f"{plan['case_id']}: numactl applies one memory policy to the whole fio process; "
                 "all jobs in this case must use the same numa_mem_policy")
        args = numactl_policy_args(effective[0])
        plan["fio_text"] = re.sub(r"(?m)^numa_mem_policy=.*\n?", "", plan["fio_text"])
        plan.update(numa_backend="numactl", numa_command_prefix=[executable, *args] if args else [])


def stamp():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds")


def write_json(path, value):
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    temporary.replace(path)


class RunLog:
    """Timestamped console/file log plus a persistent current-progress snapshot."""
    def __init__(self, root):
        self.root = root
        self.state = dict(status="starting", stage="STARTUP", started_at=stamp(),
                          case_index=None, case_count=0, case_id=None, cases=[])
        self.logger = logging.getLogger(f"fio_excel_runner.{uuid.uuid4().hex}")
        self.logger.setLevel(logging.INFO)
        self.logger.propagate = False
        formatter = logging.Formatter("%(asctime)sZ %(levelname)s %(message)s", "%Y-%m-%dT%H:%M:%S")
        formatter.converter = time.gmtime
        for handler in (logging.FileHandler(root / "runner.log", encoding="utf-8"),
                        logging.StreamHandler(sys.stdout)):
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
        self.step("STARTUP", f"runner={VERSION}; python={sys.version.split()[0]}; archive={root}")

    def note(self, message, level=logging.INFO, exc_info=False):
        index = self.state["case_index"]
        case = (f"[case {index}/{self.state['case_count']} id={self.state['case_id']}]"
                if index is not None else "[suite]")
        self.logger.log(level, f"[{self.state['stage']}]{case} {message}", exc_info=exc_info)

    def step(self, stage, message, index=None, case_id=None):
        self.state.update(stage=stage, case_index=index, case_id=case_id, updated_at=stamp())
        if index is not None and index <= len(self.state["cases"]):
            if stage in ("CASE_RUN", "CASE_DONE"):
                self.state["cases"][index-1]["status"] = "running" if stage == "CASE_RUN" else "passed"
        write_json(self.root / "run_status.json", self.state)
        self.note(message)

    def select(self, cases):
        self.state["case_count"] = len(cases)
        self.state["cases"] = [dict(index=i, case_id=c["case_id"], status="pending")
                               for i, c in enumerate(cases, 1)]
        self.step("SELECT", f"Selected {len(cases)} cases in worksheet order")

    def finish(self, status, exit_code, error=None):
        """Keep the failed stage/case and record remaining cases as skipped."""
        counts = {}
        for case in self.state["cases"]:
            path = self.root / f"{case['index']:03}_{case['case_id']}" / "run_info.json"
            if path.exists():
                info = json.loads(path.read_text(encoding="utf-8"))
                case["status"] = info["status"]
            elif status in ("failed", "interrupted"):
                case["status"] = ("interrupted" if status == "interrupted" else "error") if case["index"] == self.state["case_index"] else "skipped"
            counts[case["status"]] = counts.get(case["status"], 0) + 1
        self.state.update(status=status, exit_code=exit_code, finished_at=stamp(), counts=counts)
        if error is not None:
            self.state["error"] = error
        write_json(self.root / "run_status.json", self.state)
        self.note(f"Suite {status}; exit_code={exit_code}; cases={counts}",
                  logging.ERROR if status == "failed" else logging.INFO)

    def close(self):
        for handler in self.logger.handlers[:]:
            handler.close()
            self.logger.removeHandler(handler)


def record_current_failure(root, progress, status, error):
    index, case_id = progress.state["case_index"], progress.state["case_id"]
    if index is None:
        return
    path = root / f"{index:03}_{case_id}" / "run_info.json"
    if path.exists():
        info = json.loads(path.read_text(encoding="utf-8"))
        if info["status"] in ("pending", "planned", "running"):
            info.update(status=status, error=error, failure_stage=progress.state["stage"], finished_at=stamp())
            write_json(path, info)


def skip_pending(root, reason, progress):
    for path in sorted(root.glob("*/run_info.json")):
        try:
            info = json.loads(path.read_text(encoding="utf-8"))
            if info["status"] in ("pending", "planned"):
                info.update(status="skipped", skip_reason=reason, finished_at=stamp())
                write_json(path, info)
                progress.note(f"Skipped {info.get('case_id', path.parent.name)}: {reason}", logging.WARNING)
        except (OSError, ValueError, KeyError):
            progress.note(f"Could not update skipped status: {path}", logging.ERROR, exc_info=True)


def stop_process(process, progress=None):
    """Allow fio to flush its normal report before forced cleanup."""
    def send(sig):
        if progress:
            progress.note(f"Stopping fio process group pid={process.pid}, signal={sig}", logging.WARNING)
        try:
            if os.name == "posix":
                os.killpg(process.pid, sig)
            else:
                process.terminate()
        except ProcessLookupError:
            pass
    send(signal.SIGINT if os.name == "posix" else signal.SIGTERM)
    try:
        process.wait(timeout=10)
    except subprocess.TimeoutExpired:
        send(signal.SIGTERM)
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            if os.name == "posix":
                send(signal.SIGKILL)
            else:
                process.kill()
            try:
                process.wait(timeout=5)
            except subprocess.TimeoutExpired:
                fail("fio did not exit after SIGKILL (possibly blocked kernel I/O); no further cases will run")


def execute(command, cwd, output, timeout, progress=None):
    if progress:
        progress.note(f"Command: {shlex.join(command)}; cwd={cwd}; output={output}; timeout={timeout}s")
    with output.open("wb") as log:
        process = subprocess.Popen(command, cwd=cwd, stdout=log, stderr=subprocess.STDOUT,
                                   start_new_session=(os.name == "posix"))
        started = time.monotonic()
        if progress:
            progress.note(f"Process started: pid={process.pid}")
        try:
            while True:
                remaining = timeout - (time.monotonic() - started)
                try:
                    code = process.wait(timeout=max(0, min(10, remaining)))
                    if progress:
                        progress.note(f"Process exited: pid={process.pid}; exit_code={code}; elapsed={time.monotonic()-started:.1f}s")
                    return code, False
                except subprocess.TimeoutExpired:
                    if time.monotonic() - started >= timeout:
                        if progress:
                            progress.note(f"Watchdog timeout after {timeout}s", logging.ERROR)
                        stop_process(process, progress)
                        return process.returncode, True
                    if progress:
                        progress.note(f"Process running: pid={process.pid}; elapsed={time.monotonic()-started:.1f}s")
        except BaseException:
            stop_process(process, progress)
            raise


def extract_fio_json(combined_path, json_path):
    """Extract fio's JSON object from a combined normal,json output stream."""
    text = combined_path.read_text(encoding="utf-8", errors="replace")
    decoder = json.JSONDecoder()
    for match in re.finditer(r"\{", text):
        try:
            value, _ = decoder.raw_decode(text[match.start():])
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict) and isinstance(value.get("jobs"), list):
            write_json(json_path, value)
            return value
    fail(f"fio completed but no JSON result was found in {combined_path}")


def watchdog_seconds(plan, timeout_grace):
    if plan.get("timeout_s") is not None:
        return plan["timeout_s"]
    if not plan.get("time_based", True) or plan.get("runtime_s") is None:
        fail(f"{plan.get('case_id', 'case')}: timeout_s is required when time_based=0")
    return plan["runtime_s"] + timeout_grace


def run_suite(plans, root, fio, timeout_grace, devices, progress):
    records = []
    for index, plan in enumerate(plans, 1):
        folder = root / f"{index:03}_{plan['case_id']}"
        records.append((index, plan, folder, json.loads((folder / "run_info.json").read_text(encoding="utf-8"))))
    if not progress.state["cases"]:
        progress.select(plans)

    def save_error(folder, info, status, error, code=None):
        info.update(status=status, error=error, exit_code=code,
                    failure_stage=progress.state["stage"], finished_at=stamp())
        write_json(folder / "run_info.json", info)
        progress.note(error, logging.ERROR)
        skip_pending(root, f"Suite stopped after {info.get('case_id', folder.name)}: {error}", progress)

    # Parse every case before starting any disk workload.
    for index, plan, folder, info in records:
        try:
            progress.step("PREFLIGHT", "Checking fio job-file syntax", index, plan["case_id"])
            code, timed_out = execute([fio, "--parse-only", "jobs.fio"], folder,
                                      folder / "parse_check.txt", 60, progress)
            if code or timed_out:
                error = f"fio cannot parse {plan['case_id']}; exit_code={code}; timeout={timed_out}; see {folder / 'parse_check.txt'}"
                save_error(folder, info, "preflight_failed", error, code)
                fail(error)
            info["preflight_passed"] = True
            write_json(folder / "run_info.json", info)
            progress.note("fio job-file syntax check passed")
        except KeyboardInterrupt:
            save_error(folder, info, "interrupted", "Interrupted during fio preflight")
            raise
        except Exception as exc:
            if info.get("status") != "preflight_failed":
                save_error(folder, info, "preflight_failed", str(exc))
            raise
    for index, plan, folder, info in records:
        try:
            progress.step("CASE_RECHECK", f"Rechecking device={plan['device']} and CPUs", index, plan["case_id"])
            device = devices[plan["device"]]
            if probe_device(device["path"]) != device:
                fail("Device identity, capacity or sector size changed since preflight")
            check_device_idle(device)
            check_cpus([plan])
            check_numa_nodes([plan])
            command = [*plan.get("numa_command_prefix", []), fio,
                       "--output-format=normal,json", "--eta=never"]
            if not plan["writes"]:
                command.append("--readonly")
            command.append("jobs.fio")
            watchdog = watchdog_seconds(plan, timeout_grace)
            mode = f"time_based runtime={plan['runtime_s']}s" if plan["time_based"] else "full_range"
            progress.step("CASE_RUN", f"Starting case; device={plan['device']}; workers={plan['total_workers']}; mode={mode}; numa_backend={plan.get('numa_backend', 'none')}; watchdog={watchdog}s", index, plan["case_id"])
            info.update(status="running", started_at=stamp(), command=command,
                        case_index=index, case_count=len(plans))
            write_json(folder / "run_info.json", info)
            code, timed_out = execute(command, folder, folder / "result.txt", watchdog, progress)
            if code or timed_out:
                error = f"fio workload failed; exit_code={code}; timeout={timed_out}; see {folder / 'result.txt'}"
                save_error(folder, info, "timeout" if timed_out else "failed", error, code)
                progress.state["error"] = error
                return 1
            extract_fio_json(folder / "result.txt", folder / "result.json")
            info.update(status="passed", exit_code=code, finished_at=stamp())
            write_json(folder / "run_info.json", info)
            progress.step("CASE_DONE", f"Case passed; result={folder / 'result.txt'}", index, plan["case_id"])
        except KeyboardInterrupt:
            save_error(folder, info, "interrupted", "Interrupted; no further cases will start")
            raise
        except Exception as exc:
            save_error(folder, info, "error", str(exc))
            raise
    return 0


def arguments(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, required=True)
    parser.add_argument("--device", help="optional override for ALL selected Excel device values, e.g. /dev/nvme1n1")
    parser.add_argument("--output", type=Path, default=Path("fio_results"))
    parser.add_argument("--fio", default="fio", help="fio executable or absolute path")
    parser.add_argument("--numactl", default="numactl", help="numactl executable or absolute path")
    parser.add_argument("--numa-backend", choices=("auto", "fio", "numactl"), default="auto",
                        help="NUMA policy backend (default: auto; fall back to numactl)")
    parser.add_argument("--dry-run", action="store_true", help="generate and validate input only; no fio workload")
    parser.add_argument("--device-size-bytes", type=int, help="offline --dry-run only; assume this device capacity")
    parser.add_argument("--logical-sector-bytes", type=int, default=512, help="offline --dry-run sector size")
    parser.add_argument("--allow-write", action="store_true", help="authorize destructive writes for selected write/mixed cases")
    parser.add_argument("--case", action="append", help="select case_id; repeat to select several enabled cases")
    parser.add_argument("--timeout-grace", type=int, default=60, help="watchdog seconds in addition to runtime (default 60)")
    return parser.parse_args(argv)


def prepare_and_run(args, root, progress):
    progress.step("CONFIG_READ", f"Reading Excel: {args.excel.resolve()}")
    if args.timeout_grace < 1:
        fail("--timeout-grace must be positive")
    shutil.copyfile(args.excel, root / "input.xlsx")
    cases = load_cases(args.excel)
    progress.note(f"Loaded {len(cases)} worksheet cases")
    active = [case for case in cases if case["enabled"]]
    if args.case:
        unknown = set(args.case) - {c["case_id"] for c in active}
        if unknown:
            fail(f"Unknown or disabled selected cases: {sorted(unknown)}")
        active = [c for c in active if c["case_id"] in args.case]
    if not active:
        fail("No enabled cases selected")
    progress.select(active)
    resolve_case_devices(active, args.device)
    offline = args.device_size_bytes is not None
    if offline:
        if not args.dry_run:
            fail("--device-size-bytes is allowed only with --dry-run")
        if args.device_size_bytes <= 0 or args.logical_sector_bytes not in (512, 1024, 2048, 4096):
            fail("Invalid offline capacity or logical sector size")
    else:
        if not sys.platform.startswith("linux"):
            fail("Execution/device detection requires Linux; use offline --dry-run on other systems")
        progress.step("CPU_CHECK", "Checking requested CPU ranges against online CPUs and process affinity")
        check_cpus(active)
        progress.step("NUMA_CHECK", "Checking requested NUMA memory nodes")
        check_numa_nodes(active)
    devices = {}
    plans = []
    for index, case in enumerate(active, 1):
        name = case["effective_device"]
        progress.step("DEVICE_CHECK", f"Device={name}; source={case['device_source']}; offline={offline}", index, case["case_id"])
        if offline:
            device = dict(path=name, capacity_bytes=args.device_size_bytes,
                          logical_sector_bytes=args.logical_sector_bytes, offline=True)
        else:
            device = probe_device(name)
            check_device_idle(device)
        devices[device["path"]] = device
        progress.note(f"Device resolved: {device['path']}; capacity={device['capacity_bytes']} bytes; sector={device['logical_sector_bytes']} bytes")
        progress.step("PLAN_GENERATE", "Calculating job regions and CPU ranges", index, case["case_id"])
        plans.append(make_plan(case, device["path"], device["capacity_bytes"], device["logical_sector_bytes"]))
        progress.note(f"Plan ready: jobs={case['jobs']}; workers={case['total_workers']}; region_end={plans[-1]['region_end_exclusive']}")
    progress.step("WRITE_GUARD", "Checking write authorization for all selected cases")
    if not args.dry_run and any(p["writes"] for p in plans) and not args.allow_write:
        fail("Write/mixed cases selected: data in the first 84% may be destroyed; specify --allow-write")
    fio = shutil.which(args.fio)
    version = "not checked (dry-run)"
    progress.step("FIO_CHECK", "Checking fio executable/version" if not args.dry_run else "Dry-run: fio executable check is skipped")
    if not args.dry_run:
        if not fio:
            fail("fio executable not found; install the Flexible I/O Tester")
        version = subprocess.check_output([fio, "--version"], text=True, timeout=10).strip()
        if not re.match(r"^fio-\d", version):
            fail(f"Unexpected fio executable/version: {version!r}")
        progress.note(f"fio executable={fio}; version={version}")
        configure_numa_backend(plans, fio, args.numa_backend, args.numactl)
        for plan in plans:
            progress.note(f"{plan['case_id']}: NUMA backend={plan['numa_backend']}; "
                          f"prefix={shlex.join(plan['numa_command_prefix']) or '(none)'}")
    else:
        for plan in plans:
            plan.update(numa_backend="unresolved_dry_run", numa_command_prefix=[])
    write_json(root / "environment.json", dict(runner_version=VERSION, python=sys.version,
               platform=sys.platform, fio_version=version, devices=list(devices.values()), created_at=stamp(),
               dry_run=args.dry_run, offline=offline, numa_backend_requested=args.numa_backend))
    for index, plan in enumerate(plans, 1):
        progress.step("ARCHIVE", "Saving fio configuration, CPU mapping and run metadata", index, plan["case_id"])
        folder = root / f"{index:03}_{plan['case_id']}"
        folder.mkdir()
        (folder / "jobs.fio").write_text(plan["fio_text"], encoding="utf-8", newline="\n")
        info = {k: v for k, v in plan.items() if k not in ("fio_text", "workers", "source_values")}
        info.update(status="planned" if args.dry_run else "pending", created_at=stamp(),
                    case_index=index, case_count=len(plans))
        write_json(folder / "run_info.json", info)
        with (folder / "cpu_mapping.csv").open("w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=("job", "clone_index", "global_worker_index", "expected_cpu"))
            writer.writeheader()
            writer.writerows(plan["workers"])
    if args.dry_run:
        progress.step("PREVIEW_DONE", f"Generated {len(plans)} cases. No disk I/O workload was started.")
        if offline:
            progress.note("Offline preview: actual device, CPUs and fio capabilities have NOT been checked.")
        return 0
    import fcntl
    progress.step("DEVICE_LOCK", "Acquiring locks for all selected devices")
    with ExitStack() as stack:
        # Lock each distinct device once, including aliases of the same device.
        identities = sorted({(d['major'], d['minor']) for d in devices.values()})
        for major, minor in identities:
            lock_path = Path(f"/run/lock/fio-excel-{major}-{minor}.lock")
            lock = stack.enter_context(lock_path.open("a"))
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
            except BlockingIOError:
                fail(f"Another fio_excel_runner instance is testing device {major}:{minor}")
            progress.note(f"Device lock acquired: {major}:{minor}")
        return run_suite(plans, root, fio, args.timeout_grace, devices, progress)


def main(argv=None):
    if sys.version_info < (3, 11):
        fail("Python 3.11 or newer is required")
    args = arguments(argv)
    # Initialize the log before opening Excel or probing the target devices.
    args.output.mkdir(parents=True, exist_ok=True)
    root = args.output.resolve() / (dt.datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:8])
    root.mkdir()
    progress = RunLog(root)
    progress.state["status"] = "running"
    try:
        code = prepare_and_run(args, root, progress)
        progress.finish("preview_complete" if args.dry_run else ("passed" if code == 0 else "failed"), code)
        return code
    except KeyboardInterrupt:
        progress.note("Interrupted; aborting all remaining tasks", logging.WARNING)
        record_current_failure(root, progress, "interrupted", "User interrupted execution")
        skip_pending(root, "Suite interrupted", progress)
        progress.finish("interrupted", 130, "User interrupted execution")
        raise
    except Exception as exc:
        progress.note(f"Aborting all tasks: {exc}", logging.ERROR, exc_info=True)
        record_current_failure(root, progress, "error", str(exc))
        skip_pending(root, f"Suite aborted: {exc}", progress)
        progress.finish("failed", 2, str(exc))
        raise
    finally:
        progress.close()


if __name__ == "__main__":
    def interrupted(signum, frame):
        raise KeyboardInterrupt
    signal.signal(signal.SIGTERM, interrupted)
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("Interrupted. Existing output files were retained.", file=sys.stderr)
        sys.exit(130)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(2)
