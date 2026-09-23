import json
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

    def test_json_metrics_and_report(self):
        case = runner.load_cases(SAMPLE)[0]
        folder = self.make_attempt(case)
        fio = {"fio version": "fio-3.36", "jobs": [{
            "read": {"iops": 1000, "bw_bytes": 10485760, "runtime": 60000,
                     "total_ios": 60000, "clat_ns": {"mean": 12000, "percentile": {"95.000000": 20000, "99.000000": 30000, "99.900000": 40000, "99.990000": 50000}}},
            "write": {"iops": 0, "bw_bytes": 0, "runtime": 0, "total_ios": 0},
            "usr_cpu": 2.5, "sys_cpu": 4.5}]}
        runner.write_json(folder / "result.json", fio)
        output = self.root / "report.xlsx"
        self.assertEqual(report.main(["--excel", str(SAMPLE), "--results", str(self.root / "fio_results"),
                                      "--output", str(output)]), 0)
        with zipfile.ZipFile(output) as archive:
            self.assertIn("xl/worksheets/sheet4.xml", archive.namelist())
            ET.fromstring(archive.read("xl/worksheets/sheet2.xml"))
            content = archive.read("xl/worksheets/sheet2.xml").decode()
            self.assertIn("已通过", content)
            self.assertIn(">1000.0<", content)

    def test_changed_config_is_not_current_execution(self):
        case = runner.load_cases(SAMPLE)[0]
        self.make_attempt(case, changed=True)
        overview, _, exceptions = report.build_rows(runner.load_cases(SAMPLE), report.scan_archives(self.root / "fio_results"))
        self.assertEqual(overview[0][2], "配置已变更")
        self.assertEqual(overview[0][4], 0)
        self.assertEqual(exceptions[0][0], case["case_id"])

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


if __name__ == "__main__":
    unittest.main()
