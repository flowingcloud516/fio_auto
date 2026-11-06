import json
import csv
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import fio_excel_runner as runner
import fio_result_report as report

ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "示例.xlsx"


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def make_attempt(self, case, status="passed", changed=False):
        folder = self.root / "fio_results" / "20260923_120000_abcd" / f"001_{case['case_id']}"
        folder.mkdir(parents=True)
        info = dict(case)
        info.update(status=status, created_at="2026-09-23T04:00:00+00:00",
                    finished_at="2026-09-23T04:01:00+00:00")
        if changed:
            info["runtime_s"] += 1
        runner.write_json(folder / "run_info.json", info)
        return folder

    def write_cpu_csv(self, folder, rows):
        columns = ["timestamp_utc", "elapsed_s", "interval_s", "cpu", "selected_cpus",
                   "valid_cpus", "status", "busy_pct", "user_pct", "nice_pct",
                   "system_pct", "idle_pct", "iowait_pct", "irq_pct", "softirq_pct",
                   "steal_pct", "busy_cpu_s", "core_equiv"]
        with (folder / "cpu.csv").open("w", newline="", encoding="utf-8") as target:
            writer = csv.DictWriter(target, fieldnames=columns)
            writer.writeheader()
            writer.writerows(rows)

    def write_jobs_fio(self, folder, text):
        (folder / "jobs.fio").write_text(text, encoding="utf-8")

    def test_json_metrics_and_report(self):
        case = runner.load_cases(SAMPLE)[0]
        folder = self.make_attempt(case)
        fio = {"fio version": "fio-3.36", "jobs": [{
            "read": {"iops": 1000, "bw_bytes": 10485760, "runtime": 60000,
                     "total_ios": 60000, "clat_ns": {"mean": 12000, "percentile": {"95.000000": 20000, "99.000000": 30000, "99.900000": 40000, "99.990000": 50000}}},
            "write": {"iops": 0, "bw_bytes": 0, "runtime": 0, "total_ios": 0},
            "usr_cpu": 2.5, "sys_cpu": 4.5}]}
        runner.write_json(folder / "result.json", fio)
        info = json.loads((folder / "run_info.json").read_text(encoding="utf-8"))
        info["cpu_log"] = "cpu.csv"
        runner.write_json(folder / "run_info.json", info)
        self.write_jobs_fio(folder, "[global]\nfilename=/dev/test\n[job_001]\ncpus_allowed=2\n")
        self.write_cpu_csv(folder, [
            {"timestamp_utc": "2026-09-23T04:00:02+00:00", "interval_s": 2,
             "cpu": "cpu0", "status": "ok", "busy_pct": 100},
            {"timestamp_utc": "2026-09-23T04:00:02+00:00", "interval_s": 2,
             "cpu": "cpu2", "selected_cpus": 8, "valid_cpus": 1,
             "status": "ok", "busy_pct": 50, "user_pct": 20, "system_pct": 25,
             "idle_pct": 45, "iowait_pct": 5, "core_equiv": .5},
            {"timestamp_utc": "2026-09-23T04:00:06+00:00", "interval_s": 4,
             "cpu": "cpu2", "selected_cpus": 8, "valid_cpus": 1,
             "status": "ok", "busy_pct": 25, "user_pct": 10, "system_pct": 10,
             "idle_pct": 70, "iowait_pct": 5, "core_equiv": .25},
        ])
        output = self.root / "report.xlsx"
        self.assertEqual(report.main(["--excel", str(SAMPLE), "--results", str(self.root / "fio_results"),
                                      "--output", str(output)]), 0)
        with zipfile.ZipFile(output) as archive:
            self.assertIn("xl/worksheets/sheet5.xml", archive.namelist())
            ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            content = archive.read("xl/worksheets/sheet2.xml").decode()
            self.assertIn("已通过", content)
            self.assertIn(">1000.0<", content)
            self.assertIn(">33.333<", content)
            self.assertIn("jobs.fio", content)

    def test_cpu_metrics_use_only_bound_cpus_and_weight_intervals(self):
        folder = self.root / "cpu"
        folder.mkdir()
        self.write_cpu_csv(folder, [
            {"timestamp_utc": "t1", "interval_s": 2, "cpu": "cpu0", "busy_pct": 100},
            {"timestamp_utc": "t1", "interval_s": 2, "cpu": "cpu2",
             "busy_pct": 50, "user_pct": 20, "system_pct": 25, "iowait_pct": 5,
             "idle_pct": 45, "core_equiv": .5},
            {"timestamp_utc": "t1", "interval_s": 2, "cpu": "cpu3",
             "busy_pct": 30, "user_pct": 10, "system_pct": 15, "iowait_pct": 5,
             "idle_pct": 65, "core_equiv": .3},
            {"timestamp_utc": "t2", "interval_s": 4, "cpu": "cpu2",
             "busy_pct": 25, "user_pct": 10, "system_pct": 10, "iowait_pct": 5,
             "idle_pct": 70, "core_equiv": .25},
            {"timestamp_utc": "t2", "interval_s": 4, "cpu": "cpu3",
             "busy_pct": 15, "user_pct": 5, "system_pct": 5, "iowait_pct": 5,
             "idle_pct": 80, "core_equiv": .15},
            {"timestamp_utc": "t2", "interval_s": 4, "cpu": "selected_sum",
             "busy_pct": 90, "core_equiv": 7.2},
        ])
        metrics, note = report.metrics_from_cpu_csv(folder / "cpu.csv", {2, 3})
        self.assertEqual(note, "")
        self.assertEqual(metrics["cpu_sample_count"], 2)
        self.assertEqual(metrics["cpu_sampled_cpus"], 2)
        self.assertAlmostEqual(metrics["cpu_avg_busy_pct"], 26.667, places=3)
        self.assertAlmostEqual(metrics["cpu_avg_user_pct"], 10, places=3)
        self.assertAlmostEqual(metrics["cpu_avg_core_equiv"], .533, places=3)

    def test_jobs_fio_binding_parser_and_run_info_fallback(self):
        cpus, note = report.cpus_from_jobs_fio(ROOT / "examples" / "example_jobs.fio")
        self.assertEqual(note, "")
        self.assertEqual(report.compact_cpu_list(cpus), "2-3,6-7,11-12,16-17")

        folder = self.root / "fallback"
        folder.mkdir()
        cpus, source, note = report.bound_cpus_for_directory(folder, {
            "cpu_assignments": [{"cpus_allowed": "4-5"}, {"cpus_allowed": "9"}]
        })
        self.assertEqual(cpus, {4, 5, 9})
        self.assertEqual(source, "run_info.cpu_assignments")
        self.assertEqual(note, "")

    def test_missing_configured_cpu_log_is_reported(self):
        case = runner.load_cases(SAMPLE)[0]
        folder = self.make_attempt(case)
        info = json.loads((folder / "run_info.json").read_text(encoding="utf-8"))
        info["cpu_log"] = "cpu.csv"
        runner.write_json(folder / "run_info.json", info)
        attempts = report.scan_archives(self.root / "fio_results")
        self.assertIn("CPU 采样日志不存在", attempts[0]["metric_error"])

    def test_cpu_log_without_binding_is_not_reported_as_system_average(self):
        folder = self.root / "unbound"
        folder.mkdir()
        self.write_cpu_csv(folder, [
            {"timestamp_utc": "t1", "interval_s": 3, "cpu": "cpu0",
             "busy_pct": 90, "user_pct": 10, "system_pct": 70,
             "iowait_pct": 2, "idle_pct": 8, "core_equiv": .9},
        ])
        metrics, note, path = report.cpu_metrics_for_directory(folder)
        self.assertNotIn("cpu_avg_busy_pct", metrics)
        self.assertIn("无法从 jobs.fio 或 run_info.json 确定绑定 CPU", note)
        self.assertEqual(path, str(folder / "cpu.csv"))

    def test_changed_config_is_not_current_execution(self):
        case = runner.load_cases(SAMPLE)[0]
        self.make_attempt(case, changed=True)
        overview, _, exceptions, jobs = report.build_rows(runner.load_cases(SAMPLE), report.scan_archives(self.root / "fio_results"))
        status_col = report.OVERVIEW_HEADERS.index("当前状态")
        run_count_col = report.OVERVIEW_HEADERS.index("执行次数")
        self.assertEqual(overview[0][status_col], "配置已变更")
        self.assertEqual(overview[0][run_count_col], 0)
        self.assertEqual(exceptions[0][0], case["case_id"])
        self.assertGreater(len(jobs), 0)

    def test_overview_uses_individual_cases_columns(self):
        cases = runner.load_cases(SAMPLE)
        overview, _, _, jobs = report.build_rows(cases, [])
        self.assertEqual(report.OVERVIEW_HEADERS[:len(runner.CASE_COLUMNS)], list(runner.CASE_COLUMNS))
        row = dict(zip(report.OVERVIEW_HEADERS, overview[0]))
        self.assertEqual(row["rw"], "read")
        self.assertEqual(row["bs"], "1M")
        self.assertEqual(row["iodepth"], 32)
        self.assertEqual(row["jobs"], 1)
        self.assertEqual(jobs[0][report.JOB_HEADERS.index("rw")], "read")

        repeated = dict(zip(report.OVERVIEW_HEADERS, overview[1]))
        self.assertEqual(repeated["rw"], "randread")
        self.assertEqual(repeated["bs"], "4K")
        self.assertEqual(repeated["iodepth"], 32)
        self.assertEqual(repeated["numjobs"], 2)
        self.assertEqual(repeated["numa_mem_policy"], "")

        custom_index = next(i for i, case in enumerate(cases) if case["case_id"] == "custom_jobs")
        custom = dict(zip(report.OVERVIEW_HEADERS, overview[custom_index]))
        self.assertEqual(custom["rw"], "多值（见Job配置）")
        self.assertEqual(custom["bs"], "多值（见Job配置）")

    def test_combined_output_json_extraction(self):
        source = self.root / "result.txt"
        source.write_text("normal output\n{\"fio version\":\"fio-3.36\",\"jobs\":[]}\n", encoding="utf-8")
        destination = self.root / "result.json"
        value = runner.extract_fio_json(source, destination)
        self.assertEqual(value["fio version"], "fio-3.36")
        self.assertTrue(destination.exists())

    def test_historical_plain_text_metrics(self):
        source = self.root / "result.txt"
        source.write_text("""read: IOPS=12.5k, BW=48.8MiB/s (2928MiB/60001msec)\n"
                          "  clat (usec): min=1, max=99, avg=12.50, stdev=1.0\n"
                          "  clat percentiles (usec):\n"
                          "   | 95.00th=[   20], 99.00th=[   30], 99.90th=[   40], 99.99th=[   50]\n"
                          "  cpu          : usr=2.50%, sys=4.50%, ctx=1\n""", encoding="utf-8")
        metrics, note = report.metrics_from_text(source)
        self.assertEqual(metrics["read_iops"], 12500)
        self.assertAlmostEqual(metrics["read_bw_mib_s"], 48.8)
        self.assertEqual(metrics["read_clat_p9999_us"], 50)
        self.assertEqual(note, "")

    def test_report_without_excel_from_one_raw_result(self):
        source = self.root / "my_raw_result.txt"
        source.write_text("read: IOPS=1000, BW=10MiB/s (600MiB/60000msec)\n", encoding="utf-8")
        output = self.root / "raw_report.xlsx"
        self.assertEqual(report.main(["--results", str(source), "--output", str(output)]), 0)
        attempts = report.scan_archives(source)
        self.assertEqual(len(attempts), 1)
        self.assertEqual(attempts[0]["status"], "collected")
        self.assertEqual(attempts[0]["read_iops"], 1000)
        self.assertEqual(attempts[0]["result_path"], str(source))
        with zipfile.ZipFile(output) as archive:
            self.assertIn("xl/worksheets/sheet3.xml", archive.namelist())
            self.assertNotIn("xl/worksheets/sheet4.xml", archive.namelist())
            detail = archive.read("xl/worksheets/sheet2.xml").decode()
            self.assertIn("已收集（执行状态未知）", detail)
            self.assertIn("my_raw_result.txt", detail)

    def test_report_without_excel_scans_raw_result_directories(self):
        first = self.root / "raw" / "001_case_a"
        second = self.root / "raw" / "002_case_b"
        first.mkdir(parents=True)
        second.mkdir(parents=True)
        (first / "result.txt").write_text("read: IOPS=10, BW=1MiB/s (60MiB/60000msec)\n")
        self.write_jobs_fio(first, "[global]\nfilename=/dev/test\n[job_001]\ncpus_allowed=4-5\n")
        self.write_cpu_csv(first, [
            {"timestamp_utc": "t1", "interval_s": 3, "cpu": "cpu4",
             "busy_pct": 10, "user_pct": 2, "system_pct": 7, "iowait_pct": 1,
             "idle_pct": 89, "core_equiv": .1},
            {"timestamp_utc": "t1", "interval_s": 3, "cpu": "cpu5",
             "busy_pct": 15, "user_pct": 3, "system_pct": 9, "iowait_pct": 1,
             "idle_pct": 84, "core_equiv": .15},
        ])
        runner.write_json(second / "result.json", {"fio version":"fio-3.34", "jobs":[{
            "read":{"iops":20, "bw_bytes":2097152, "runtime":60000, "total_ios":1200},
            "write":{"iops":0, "bw_bytes":0, "runtime":0, "total_ios":0}}]})
        attempts = report.scan_archives(self.root / "raw")
        self.assertEqual({a["case_id"] for a in attempts}, {"case_a", "case_b"})
        self.assertEqual(len(attempts), 2)
        first_attempt = next(a for a in attempts if a["case_id"] == "case_a")
        self.assertEqual(first_attempt["cpu_avg_busy_pct"], 12.5)
        self.assertEqual(first_attempt["cpu_sampled_cpus"], 2)
        self.assertEqual(first_attempt["cpu_bound_cpus"], "4-5")


if __name__ == "__main__":
    unittest.main()
