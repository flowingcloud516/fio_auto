from pathlib import Path
import tempfile
import unittest

import fio_case_generator as generator
import fio_excel_runner as runner


class GeneratorTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def base(self):
        return ["--rw", "read,write,randread", "--bs", "8K,128K,1M",
                "--iodepth", "1,2,4,8", "--device", "/dev/nvme1n1",
                "--cpu-start", "2", "--cpu-end", "17"]

    def test_requested_cartesian_product(self):
        output = self.root / "cases.xlsx"
        self.assertEqual(generator.main(self.base() + ["--output", str(output)]), 0)
        cases = runner.load_cases(output)
        self.assertEqual(len(cases), 36)
        self.assertEqual({j["rw"] for c in cases for j in c["job_list"]}, {"read", "write", "randread"})
        self.assertEqual({j["bs_bytes"] for c in cases for j in c["job_list"]}, {8192, 131072, 1048576})
        self.assertEqual({j["iodepth"] for c in cases for j in c["job_list"]}, {1, 2, 4, 8})

    def test_optional_dimensions_and_mixed_ratio(self):
        output = self.root / "cases.xlsx"
        argv = ["--rw", "randrw", "--bs", "4K", "--iodepth", "1",
                "--jobs", "1,2", "--numjobs", "1,4", "--runtime", "30,60",
                "--rwmixread", "70", "--device", "/dev/sdb", "--cpu-start", "0",
                "--cpu-end", "7", "--output", str(output)]
        generator.main(argv)
        cases = runner.load_cases(output)
        self.assertEqual(len(cases), 8)
        self.assertEqual({c["runtime_s"] for c in cases}, {30, 60})
        self.assertEqual({j["rwmixread"] for c in cases for j in c["job_list"]}, {70})

    def test_limits_and_cpu_validation(self):
        with self.assertRaisesRegex(runner.ConfigError, "exceeding"):
            generator.main(self.base() + ["--max-cases", "10", "--output", str(self.root/"x.xlsx")])
        with self.assertRaisesRegex(runner.ConfigError, "CPU range size"):
            generator.main(self.base() + ["--numjobs", "17", "--output", str(self.root/"y.xlsx")])


if __name__ == "__main__":
    unittest.main()
