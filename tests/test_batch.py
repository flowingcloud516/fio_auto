from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

import runner_batch as batch


ROOT = Path(__file__).resolve().parents[1]


class BatchRunnerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.base = self.root / "nightly"
        self.excel = Path(f"{self.base}.xlsx")
        self.excel.write_bytes(b"test placeholder")

    def tearDown(self):
        self.temp.cleanup()

    @patch("runner_batch.subprocess.run")
    def test_derives_all_five_paths_and_forwards_runner_options(self, run):
        run.side_effect = [subprocess.CompletedProcess([], 0), subprocess.CompletedProcess([], 0)]
        code = batch.main([str(self.base), "--dry-run", "--device", "/dev/test"])
        self.assertEqual(code, 0)
        runner_command = run.call_args_list[0].args[0]
        self.assertEqual(runner_command, [
            sys.executable, str(ROOT / "fio_excel_runner.py"),
            "--excel", str(self.excel.resolve()),
            "--output", str(self.base.resolve()),
            "--dry-run", "--device", "/dev/test",
        ])
        report_command = run.call_args_list[1].args[0]
        self.assertEqual(report_command, [
            sys.executable, str(ROOT / "fio_result_report.py"),
            "--excel", str(self.excel.resolve()),
            "--results", str(self.base.resolve()),
            "--output", str(Path(f"{self.base}_result.xlsx").resolve()),
        ])

    @patch("runner_batch.subprocess.run")
    def test_runner_failure_still_builds_report_and_is_returned(self, run):
        run.side_effect = [subprocess.CompletedProcess([], 7), subprocess.CompletedProcess([], 0)]
        self.assertEqual(batch.main([str(self.base)]), 7)
        self.assertEqual(run.call_count, 2)

    @patch("runner_batch.subprocess.run")
    def test_missing_excel_stops_before_children(self, run):
        self.excel.unlink()
        self.assertEqual(batch.main([str(self.base)]), 2)
        run.assert_not_called()

    def test_accepts_xlsx_suffix_and_rejects_path_overrides(self):
        self.assertEqual(batch.suite_paths(str(self.excel))[0], self.excel.resolve())
        with self.assertRaises(SystemExit):
            batch.main([str(self.base), "--output", "other"])
        with self.assertRaises(SystemExit):
            batch.main([str(self.base), "--excel=other.xlsx"])


if __name__ == "__main__":
    unittest.main()
