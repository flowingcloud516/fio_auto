"""Portable tests: run with python3.11 -m unittest discover -s tests -v.

These test planning, XLSX parsing and subprocess handling; they do NOT perform
block-device I/O. Linux/fio integration must be checked on the target machine.
"""
import ast
import copy
import configparser
import json
from pathlib import Path
import re
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

HERE = Path(__file__).resolve().parent
PROJECT_ROOT = HERE.parent
sys.path.insert(0, str(PROJECT_ROOT))
import fio_excel_runner as runner

SAMPLE = PROJECT_ROOT / "示例.xlsx"


class RunnerTests(unittest.TestCase):
    def setUp(self):
        self.cases = runner.load_cases(SAMPLE)
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.doCleanups()  # Close log handlers before deleting temp files on Windows.
        self.temp.cleanup()

    def workbook_edit(self, member, change):
        output = self.root / "edited.xlsx"
        with zipfile.ZipFile(SAMPLE) as source, zipfile.ZipFile(output, "w") as target:
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename == member:
                    data = change(data)
                target.writestr(entry, data)
        return output

    def test_python_311_syntax(self):
        ast.parse((PROJECT_ROOT / "fio_excel_runner.py").read_text(encoding="utf-8"), feature_version=(3, 11))

    def test_real_workbook_and_defaults(self):
        self.assertEqual(len(self.cases), 6)
        self.assertEqual([c["enabled"] for c in self.cases], [True, True, False, False, False, False])
        self.assertEqual(self.cases[0]["runtime_s"], 60)
        self.assertTrue(self.cases[0]["time_based"])
        self.assertIsNone(self.cases[0]["timeout_s"])
        self.assertEqual(self.cases[1]["total_workers"], 8)
        self.assertEqual(self.cases[0]["requested_device"], "/dev/nvme1n1")

    def test_legacy_workbook_defaults_time_based(self):
        cells = runner.read_xlsx(SAMPLE)
        for coordinate in [key for key in cells["Cases"] if key.startswith(("N", "O"))]:
            del cells["Cases"][coordinate]
        with patch.object(runner, "read_xlsx", return_value=cells):
            cases = runner.load_cases(SAMPLE)
        self.assertTrue(all(case["time_based"] for case in cases))
        self.assertTrue(all(case["timeout_s"] is None for case in cases))

    def test_watchdog_selection(self):
        self.assertEqual(runner.watchdog_seconds({"case_id":"timed", "time_based":True,
                                                  "runtime_s":60, "timeout_s":None}, 30), 90)
        self.assertEqual(runner.watchdog_seconds({"case_id":"full", "time_based":False,
                                                  "runtime_s":None, "timeout_s":7200}, 30), 7200)

    def test_device_excel_and_cli_precedence(self):
        runner.resolve_case_devices(self.cases)
        self.assertEqual(self.cases[0]["effective_device"], "/dev/nvme1n1")
        self.assertEqual(self.cases[0]["device_source"], "excel")
        runner.resolve_case_devices(self.cases, "/dev/sdb")
        self.assertTrue(all(c["effective_device"] == "/dev/sdb" for c in self.cases))
        self.assertEqual(self.cases[0]["device_source"], "command_line")

    def test_system_disk_partition_parent_is_protected(self):
        sys_block = self.root / "sys_dev_block"
        partition = sys_block / "8_2"
        disk = sys_block / "8_0"
        partition.mkdir(parents=True)
        disk.mkdir()
        (partition / "dev").write_text("8:2\n")
        (partition / "partition").write_text("2\n")
        # In real sysfs the resolved partition's parent contains the whole-disk dev file.
        (sys_block / "dev").write_text("8:0\n")
        (disk / "dev").write_text("8:0\n")
        mountinfo = self.root / "mountinfo"
        mountinfo.write_text("36 25 8:2 / / rw,relatime - ext4 /dev/sda2 rw\n")
        resolver = lambda root, device_id: root / device_id.replace(":", "_")
        with patch.object(runner, "block_device_entry", side_effect=resolver):
            self.assertEqual(runner.system_disk_ids(mountinfo, sys_block), {"8:2", "8:0"})
            with self.assertRaisesRegex(runner.ConfigError, "Refusing to test system disk"):
                runner.check_not_system_disk({"path":"/dev/sda", "major":8, "minor":0},
                                             mountinfo, sys_block)

    def test_system_disk_traces_device_mapper_slaves(self):
        sys_block = self.root / "sys_dev_block"
        dm = sys_block / "253_0"
        slave = dm / "slaves" / "sda2"
        backing = sys_block / "8_2"
        dm.mkdir(parents=True)
        slave.mkdir(parents=True)
        backing.mkdir()
        (dm / "dev").write_text("253:0\n")
        (slave / "dev").write_text("8:2\n")
        (backing / "dev").write_text("8:2\n")
        mountinfo = self.root / "mountinfo"
        mountinfo.write_text("36 25 253:0 / / rw,relatime - ext4 /dev/mapper/root rw\n")
        resolver = lambda root, device_id: root / device_id.replace(":", "_")
        with patch.object(runner, "block_device_entry", side_effect=resolver):
            self.assertEqual(runner.system_disk_ids(mountinfo, sys_block), {"253:0", "8:2"})

    def test_system_disk_check_requires_root_mount(self):
        mountinfo = self.root / "mountinfo"
        mountinfo.write_text("36 25 0:5 / /tmp rw - tmpfs tmpfs rw\n")
        with self.assertRaisesRegex(runner.ConfigError, "root filesystem"):
            runner.system_disk_ids(mountinfo, self.root / "sys")

    def test_missing_or_invalid_device(self):
        self.cases[0]["requested_device"] = ""
        with self.assertRaisesRegex(runner.ConfigError, "missing device"):
            runner.resolve_case_devices(self.cases)
        runner.resolve_case_devices(self.cases, "/dev/sdb")
        for value in ("sdb", "/dev/sdb\nrw=write", "/dev/../etc/passwd"):
            self.cases[0]["requested_device"] = value
            with self.assertRaises(runner.ConfigError):
                runner.resolve_case_devices(self.cases)

    def test_cli_without_device_reads_excel(self):
        code = runner.main(["--excel", str(SAMPLE), "--dry-run", "--device-size-bytes",
                            "1000000000000", "--output", str(self.root)])
        self.assertEqual(code, 0)
        archive = next(self.root.iterdir())
        for path in archive.glob("*/jobs.fio"):
            self.assertIn("filename=/dev/nvme1n1\n", path.read_text())

    def test_each_device_uses_own_capacity(self):
        cases = self.cases[:2]
        cases[0]["requested_device"] = "/dev/sdb"
        cases[1]["requested_device"] = "/dev/sdc"
        def probe(name):
            return dict(path=name, major=8, minor=16 if name == "/dev/sdb" else 32,
                        capacity_bytes=10**12 if name == "/dev/sdb" else 2*10**12,
                        logical_sector_bytes=512)
        with patch.object(runner, "load_cases", return_value=cases), \
             patch.object(runner.sys, "platform", "linux"), \
             patch.object(runner, "check_cpus"), patch.object(runner, "check_device_idle"), \
             patch.object(runner, "probe_device", side_effect=probe):
            runner.main(["--excel", str(SAMPLE), "--dry-run", "--output", str(self.root)])
        archive = next(self.root.iterdir())
        first = json.loads((archive/"001_seq_read_1m"/"run_info.json").read_text())
        second = json.loads((archive/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(first["device"], "/dev/sdb")
        self.assertEqual(second["device"], "/dev/sdc")
        self.assertEqual(first["capacity_bytes"], 10**12)
        self.assertEqual(second["capacity_bytes"], 2*10**12)
        self.assertLessEqual(first["region_end_exclusive"], 840000000000)
        self.assertEqual(second["region_end_exclusive"], 1680000000000)

    def test_job_overrides(self):
        custom = next(c for c in self.cases if c["case_id"] == "custom_jobs")
        self.assertEqual(custom["total_workers"], 4)
        self.assertEqual([j["numjobs"] for j in custom["job_list"]], [1, 3])
        self.assertEqual([j["rw"] for j in custom["job_list"]], ["read", "randwrite"])
        self.assertEqual([j["iodepth"] for j in custom["job_list"]], [16, 32])

    def test_cpu_distribution_property(self):
        for start in (0, 2, 251):
            for length in range(1, 33):
                counts_options = [[1], [length], [min(4,length)], [1]*min(8,length),
                                  [1]*(length+3), [max(1,length//3)]*4,
                                  [1,min(length,4),min(length,2)]]
                for counts in counts_options:
                    end = start + length - 1
                    groups = runner.allocate_job_cpus(start, end, counts)
                    self.assertEqual(len(groups),len(counts))
                    for cpus, count in zip(groups,counts):
                        self.assertEqual(len(cpus),count)
                        self.assertEqual(cpus,list(range(cpus[0],cpus[0]+count)))
                        self.assertTrue(all(start <= x <= end for x in cpus))
                    if sum(counts) <= length:
                        flattened = [cpu for group in groups for cpu in group]
                        self.assertEqual(len(flattened),len(set(flattened)))
                        self.assertEqual(groups[0][0],start)
                        if len(groups)>1:
                            self.assertEqual(groups[-1][-1],end)

    def test_contiguous_range_251_example(self):
        case = copy.deepcopy(self.cases[0])
        case["cpu_start"],case["cpu_end"] = 251,258
        for count, expected in ((1,"251"),(4,"251-254")):
            case["job_list"][0]["numjobs"] = count
            case["total_workers"] = count
            plan = runner.make_plan(case,"/dev/test",10**12,512)
            self.assertIn(f"cpus_allowed={expected}\n",plan["fio_text"])
            self.assertEqual([w["expected_cpu"] for w in plan["workers"]],list(range(251,251+count)))

    def test_one_job_needs_enough_distinct_cpus(self):
        with self.assertRaisesRegex(runner.ConfigError,"numjobs"):
            runner.allocate_job_cpus(251,253,[4])

    def test_heterogeneous_jobs_cpu_mapping(self):
        case = next(c for c in self.cases if c["case_id"] == "custom_jobs")
        plan = runner.make_plan(case,"/dev/test",10**12,512)
        self.assertEqual([r["cpus_allowed"] for r in plan["cpu_assignments"]],["2","15-17"])
        # fio split uses the global instance index, so the second job rotates.
        self.assertEqual([w["expected_cpu"] for w in plan["workers"]],[2,16,17,15])

    def test_cpu_mapping_across_distinct_jobs(self):
        case = copy.deepcopy(self.cases[1])
        case["cpu_start"], case["cpu_end"] = 2, 4
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        self.assertEqual([w["expected_cpu"] for w in plan["workers"]], [2,3,3,4,2,3,3,4])
        self.assertEqual([r["cpus_allowed"] for r in plan["cpu_assignments"]],
                         ["2-3","3-4","2-3","3-4"])
        self.assertEqual(plan["fio_text"].count("numjobs=2"), 1)
        self.assertEqual(plan["fio_text"].count("cpus_allowed_policy=split"), 4)
        self.assertEqual(plan["fio_text"].count("[job_"), 4)

    def test_common_job_options_are_promoted_to_global(self):
        case = copy.deepcopy(self.cases[1])
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        parsed = configparser.ConfigParser(interpolation=None)
        parsed.read_string(plan["fio_text"])
        for name in ("rw", "bs", "iodepth", "numjobs"):
            self.assertIn(name, parsed["global"])
            self.assertEqual(len(re.findall(rf"(?m)^{name}=", plan["fio_text"])), 1)
        self.assertEqual(plan["estimated_io_buffer_bytes"],
                         sum(j["numjobs"] * j["iodepth"] * j["bs_bytes"]
                             for j in case["job_list"]))

    def test_only_heterogeneous_options_stay_in_jobs(self):
        case = copy.deepcopy(self.cases[1])
        case["job_list"][0]["iodepth"] = 99
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        parsed = configparser.ConfigParser(interpolation=None)
        parsed.read_string(plan["fio_text"])
        self.assertNotIn("iodepth", parsed["global"])
        self.assertEqual(plan["fio_text"].count("iodepth="), case["jobs"])
        for name in ("rw", "bs", "numjobs"):
            self.assertIn(name, parsed["global"])

    def test_range_and_fixed_parameters(self):
        for case in self.cases:
            for capacity in (123456789, 10**12, 4000787030016):
                plan = runner.make_plan(case, "/dev/test", capacity, 4096)
                self.assertLessEqual(plan["region_end_exclusive"] * 100, capacity * 84)
                self.assertEqual(plan["region_end_exclusive"] % 4096, 0)
                for job in case["job_list"]:
                    self.assertEqual(plan["region_end_exclusive"] % job["bs_bytes"], 0)
                for option in runner.FIXED:
                    self.assertIn(option + "\n", plan["fio_text"])
                if case["time_based"]:
                    self.assertIn("time_based=1\n", plan["fio_text"])
                else:
                    self.assertNotIn("time_based=", plan["fio_text"])
                    self.assertNotIn("runtime=", plan["fio_text"])
                self.assertIn("offset=0\n", plan["fio_text"])
                self.assertNotIn("offset_increment", plan["fio_text"])

    def test_invalid_range_and_sector(self):
        with self.assertRaises(runner.ConfigError):
            runner.make_plan(self.cases[0], "/dev/test", 1024, 512)
        case = copy.deepcopy(self.cases[0])
        case["job_list"][0]["bs_bytes"] = 512
        with self.assertRaises(runner.ConfigError):
            runner.make_plan(case, "/dev/test", 10**12, 4096)

    def test_readable_percent_ranges(self):
        single = runner.make_plan(self.cases[0], "/dev/test", 10**12 + 37, 512)
        self.assertIn("size=84%\n", single["fio_text"])
        case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "custom_jobs"))
        for job in case["job_list"]:
            job["bs_bytes"] = 4096
        plan = runner.make_plan(case, "/dev/test", 4096 * 100 * 1000000, 4096)
        self.assertEqual([r["fio_size"] for r in plan["job_regions"]], ["42%", "42%"])
        self.assertEqual([r["fio_offset"] for r in plan["job_regions"]], ["0", "42%"])

    def test_unaligned_percent_uses_bytes(self):
        case = next(c for c in self.cases if c["case_id"] == "custom_jobs")
        plan = runner.make_plan(case, "/dev/test", 10**12 + 1, 512)
        self.assertTrue(all("%" not in r["fio_size"] for r in plan["job_regions"]))
        self.assertTrue(all("%" not in r["fio_offset"] for r in plan["job_regions"]))
        self.assertEqual(plan["job_regions"][0]["nominal_range"], "[0%, 42%)")
        self.assertEqual(plan["job_regions"][1]["nominal_range"], "[42%, 84%)")

    def test_partition_boundaries_and_fio_arguments_property(self):
        for count in (1, 2, 3, 4, 5, 7, 16, 85):
            for capacity in (123456789, 10**12 + 37, 4096*100*1000000):
                case = copy.deepcopy(self.cases[0])
                case["jobs"] = count
                case["job_list"] = [dict(job_index=i+1, rw="read", bs_bytes=4096 if i%2 else 8192,
                                         iodepth=32, numjobs=2, rwmixread=None) for i in range(count)]
                case["total_workers"] = count * 2
                plan = runner.make_plan(case, "/dev/test", capacity, 4096)
                parsed = configparser.ConfigParser(interpolation=None)
                parsed.read_string(plan["fio_text"])
                self.assertNotIn("size", parsed["global"])
                prior_end = 0
                lengths = []
                for region, job in zip(plan["job_regions"], case["job_list"]):
                    options = parsed[region["job"]]
                    bs = job["bs_bytes"]
                    offset = options["offset"]
                    size = options["size"]
                    if offset.endswith("%"):
                        offset = capacity * int(offset[:-1]) // 100
                        offset = (offset + bs - 1) // bs * bs
                    else:
                        offset = int(offset)
                    if size.endswith("%"):
                        size = capacity * int(size[:-1]) // 100
                        size = min(size, capacity-offset) // bs * bs
                    else:
                        size = int(size)
                    self.assertEqual(offset, prior_end)
                    self.assertEqual(offset, region["offset_bytes"])
                    self.assertEqual(size, region["size_bytes"])
                    self.assertGreaterEqual(size, bs)
                    self.assertEqual(offset % bs, 0)
                    self.assertEqual(size % bs, 0)
                    self.assertLessEqual((offset+size)*100, capacity*84)
                    prior_end = offset+size
                    lengths.append(size)
                self.assertEqual(prior_end, plan["region_end_exclusive"])
                self.assertLessEqual(max(lengths)-min(lengths), plan["region_alignment_bytes"])

    def test_numjobs_does_not_change_independent_job_regions(self):
        case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "custom_jobs"))
        first = runner.make_plan(case, "/dev/test", 10**12, 512)
        case["job_list"][0]["numjobs"] = 9
        case["total_workers"] = 12
        second = runner.make_plan(case, "/dev/test", 10**12, 512)
        self.assertEqual(first["job_regions"], second["job_regions"])
        self.assertEqual(len(second["workers"]), 12)

    def test_reject_partition_too_small(self):
        case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "custom_jobs"))
        for job in case["job_list"]:
            job["bs_bytes"] = 4096
        with self.assertRaisesRegex(runner.ConfigError, "every independent job"):
            runner.make_plan(case, "/dev/test", 8192, 4096)

    def test_block_size_units_and_invalid_values(self):
        for text, expected in (("4k",4096),("1m",1048576),("4KiB",4096),("4096",4096)):
            self.assertEqual(runner.block_size(text,"bs"),expected)
        for value in ("0", "-4k", "4k,8k", "4k\nrw=write", "1.5k"):
            with self.assertRaises(runner.ConfigError):
                runner.block_size(value,"bs")

    def test_numa_memory_policy_validation_and_job_output(self):
        for value in ("default", "local", "prefer:1", "bind:0", "bind:0-1", "interleave:0,2-3", "bind:all"):
            self.assertEqual(runner.numa_mem_policy(value, "policy"), value)
        for value in ("bind", "prefer:0-1", "local:0", "bind:2-1", "unknown:0", "bind:0,,1"):
            with self.assertRaises(runner.ConfigError):
                runner.numa_mem_policy(value, "policy")
        case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "custom_jobs"))
        case["job_list"][0]["numa_mem_policy"] = "bind:0"
        case["job_list"][1]["numa_mem_policy"] = "prefer:1"
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        self.assertEqual(plan["fio_text"].count("numa_mem_policy="), 2)
        self.assertIn("numa_mem_policy=bind:0\n", plan["fio_text"])
        self.assertIn("numa_mem_policy=prefer:1\n", plan["fio_text"])

    def test_numa_nodes_must_be_online(self):
        case = copy.deepcopy(self.cases[0])
        case["job_list"][0]["numa_mem_policy"] = "bind:0-1"
        online = self.root / "numa_online"
        online.write_text("0-1\n", encoding="ascii")
        runner.check_numa_nodes([case], online)
        online.write_text("0\n", encoding="ascii")
        with self.assertRaisesRegex(runner.ConfigError, "not online"):
            runner.check_numa_nodes([case], online)

    def test_fio_numa_build_capability_check(self):
        case = copy.deepcopy(self.cases[0])
        case["job_list"][0]["numa_mem_policy"] = "bind:0"
        supported = subprocess.CompletedProcess([], 0, "numa_mem_policy: NUMA memory policy\n", "")
        with patch.object(runner.subprocess, "run", return_value=supported):
            runner.check_fio_numa_support("fio", [case])
        unsupported = subprocess.CompletedProcess([], 1, "", "No such command: numa_mem_policy")
        with patch.object(runner.subprocess, "run", return_value=unsupported):
            with self.assertRaisesRegex(runner.ConfigError, "built with libnuma"):
                runner.check_fio_numa_support("fio", [case])

    def test_numactl_policy_translation(self):
        self.assertEqual(runner.numactl_policy_args("bind:2"), ["--membind=2"])
        self.assertEqual(runner.numactl_policy_args("prefer:1"), ["--preferred=1"])
        self.assertEqual(runner.numactl_policy_args("interleave:0-1"), ["--interleave=0-1"])
        self.assertEqual(runner.numactl_policy_args("local"), ["--localalloc"])
        self.assertEqual(runner.numactl_policy_args("default"), [])

    def test_auto_numa_backend_falls_back_to_numactl(self):
        case = copy.deepcopy(self.cases[0])
        for job in case["job_list"]:
            job["numa_mem_policy"] = "bind:2"
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        with patch.object(runner, "fio_supports_numa", return_value=False), \
             patch.object(runner.shutil, "which", return_value="/usr/bin/numactl"):
            runner.configure_numa_backend([plan], "fio")
        self.assertEqual(plan["numa_backend"], "numactl")
        self.assertEqual(plan["numa_command_prefix"], ["/usr/bin/numactl", "--membind=2"])
        self.assertNotIn("numa_mem_policy=", plan["fio_text"])

    def test_numactl_rejects_different_job_policies(self):
        case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "custom_jobs"))
        case["job_list"][0]["numa_mem_policy"] = "bind:0"
        case["job_list"][1]["numa_mem_policy"] = "bind:1"
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        with patch.object(runner.shutil, "which", return_value="/usr/bin/numactl"):
            with self.assertRaisesRegex(runner.ConfigError, "whole fio process"):
                runner.configure_numa_backend([plan], "fio", "numactl")

    def test_forced_numactl_requires_executable(self):
        case = copy.deepcopy(self.cases[0])
        case["job_list"][0]["numa_mem_policy"] = "bind:0"
        plan = runner.make_plan(case, "/dev/test", 10**12, 512)
        with patch.object(runner.shutil, "which", return_value=None):
            with self.assertRaisesRegex(runner.ConfigError, "not found"):
                runner.configure_numa_backend([plan], "fio", "numactl")

    def test_formula_rejected_even_with_cached_value(self):
        def edit(data):
            root = ET.fromstring(data)
            cell = root.find(".//{*}c[@r='J2']")
            if cell is None:
                row = root.find(".//{*}row[@r='2']")
                ns = row.tag.split("}")[0] + "}"
                cell = ET.SubElement(row, ns+"c", {"r":"J2"})
            ns = cell.tag.split("}")[0] + "}"
            ET.SubElement(cell, ns+"f").text = "30+30"
            ET.SubElement(cell, ns+"v").text = "60"
            return ET.tostring(root)
        path = self.workbook_edit("xl/worksheets/sheet1.xml", edit)
        with self.assertRaisesRegex(runner.ConfigError,"formulas"):
            runner.load_cases(path)

    def test_inline_strings_and_blank_styled_rows(self):
        def edit(data):
            root = ET.fromstring(data)
            cell = root.find(".//{*}c[@r='A2']")
            ns = cell.tag.split("}")[0] + "}"
            cell.clear()
            cell.attrib.update(r="A2",t="inlineStr")
            ET.SubElement(ET.SubElement(cell, ns+"is"),ns+"t").text = "inline_case"
            row = ET.SubElement(root.find("{*}sheetData"),ns+"row",{"r":"999"})
            ET.SubElement(row,ns+"c",{"r":"A999","s":"0"})
            return ET.tostring(root)
        cases = runner.load_cases(self.workbook_edit("xl/worksheets/sheet1.xml", edit))
        self.assertEqual(len(cases),6)
        self.assertEqual(cases[0]["case_id"],"inline_case")

    def test_duplicate_case_and_unknown_override_rejected(self):
        cells = runner.read_xlsx(SAMPLE)
        duplicate = copy.deepcopy(cells)
        duplicate["Cases"]["A3"] = duplicate["Cases"]["A2"]
        with patch.object(runner,"read_xlsx",return_value=duplicate):
            with self.assertRaisesRegex(runner.ConfigError,"duplicate case_id"):
                runner.load_cases(SAMPLE)
        invalid = copy.deepcopy(cells)
        invalid["JobOverrides"]["A2"] = "missing"
        with patch.object(runner,"read_xlsx",return_value=invalid):
            with self.assertRaisesRegex(runner.ConfigError,"unknown case_id"):
                runner.load_cases(SAMPLE)

    def test_missing_optional_sheet_and_reordered_headers(self):
        cells = runner.read_xlsx(SAMPLE)
        del cells["JobOverrides"]
        rows = cells["Cases"]
        for i in range(1,8):
            a,b = rows.get(f"A{i}"),rows.get(f"B{i}")
            rows[f"A{i}"],rows[f"B{i}"] = b,a
        with patch.object(runner,"read_xlsx",return_value=cells):
            cases = runner.load_cases(SAMPLE)
        self.assertEqual(next(c for c in cases if c["case_id"] == "custom_jobs")["total_workers"],4)
        self.assertTrue(cases[0]["enabled"])

    def test_offline_cli_and_archive(self):
        command = [sys.executable,str(PROJECT_ROOT/"fio_excel_runner.py"),"--excel",str(SAMPLE),
                   "--device","/dev/test","--dry-run","--device-size-bytes","1000000000000",
                   "--output",str(self.root),"--case","randread_4k"]
        result = subprocess.run(command,capture_output=True,text=True)
        self.assertEqual(result.returncode,0,result.stderr)
        archive = next(self.root.iterdir())
        info = json.loads((archive/"001_randread_4k"/"run_info.json").read_text())
        self.assertEqual(info["status"],"planned")
        self.assertEqual(info["selected_cpus"],[2,3,6,7,11,12,16,17])
        self.assertTrue((archive/"input.xlsx").exists())
        command = (archive/"001_randread_4k"/"fio.cmd").read_text(encoding="utf-8")
        self.assertEqual(command.strip(), "fio --output-format=normal,json --eta=never --readonly jobs.fio")
        self.assertFalse(list(archive.glob("**/result.txt")))

    def test_workload_command_includes_numactl(self):
        plan = runner.make_plan(self.cases[0], "/dev/test", 10**12, 512)
        plan["numa_command_prefix"] = ["/usr/bin/numactl", "--membind=2"]
        self.assertEqual(runner.workload_command(plan, "/usr/bin/fio"), [
            "/usr/bin/numactl", "--membind=2", "/usr/bin/fio",
            "--output-format=normal,json", "--eta=never", "--readonly", "jobs.fio"])

    def test_offline_capacity_forbidden_for_execution(self):
        with self.assertRaisesRegex(runner.ConfigError,"only with --dry-run"):
            runner.main(["--excel",str(SAMPLE),"--device","/dev/test","--device-size-bytes","1024",
                         "--output",str(self.root)])

    def test_subprocess_plain_output_and_nonzero(self):
        log = self.root/"result.txt"
        code,timed_out = runner.execute([sys.executable,"-c", "import sys; print('plain fio-like output'); print('stderr detail', file=sys.stderr); sys.exit(7)"],self.root,log,10)
        self.assertEqual(code,7)
        self.assertFalse(timed_out)
        self.assertIn("plain fio-like output",log.read_text())
        self.assertIn("stderr detail",log.read_text())

    def test_subprocess_timeout(self):
        code,timed_out = runner.execute([sys.executable,"-c","import time; print('started', flush=True); time.sleep(30)"], self.root,self.root/"timeout.txt",0.5)
        self.assertTrue(timed_out)
        self.assertIsNotNone(code)

    def test_resource_sample_collects_process_system_and_numa_memory(self):
        proc = self.root / "proc"
        process = proc / "123"
        process.mkdir(parents=True)
        (process / "stat").write_text("123 (fio worker) S 1 123 123 0 0\n")
        (process / "status").write_text("VmRSS:\t1024 kB\nVmHWM:\t2048 kB\n")
        (proc / "meminfo").write_text("MemAvailable: 9000 kB\nMemFree: 8000 kB\nSwapFree: 7000 kB\n")
        (proc / "self").mkdir()
        (proc / "self" / "status").write_text("VmRSS: 512 kB\n")
        nodes = self.root / "nodes" / "node2"
        nodes.mkdir(parents=True)
        (nodes / "meminfo").write_text("Node 2 MemFree: 6000 kB\nNode 2 MemUsed: 4000 kB\n")
        sample = runner.resource_sample(123, 2.5, proc, nodes.parent)
        self.assertEqual(sample["process_count"], 1)
        self.assertEqual(sample["fio_group_rss_kb"], 1024)
        self.assertEqual(sample["fio_group_vmhwm_kb"], 2048)
        self.assertEqual(sample["runner_rss_kb"], 512)
        self.assertEqual(sample["mem_available_kb"], 9000)
        self.assertEqual(json.loads(sample["numa_nodes"])["node2"]["MemFree_kb"], 6000)
        output = self.root / "memory.csv"
        runner.append_resource_sample(output, sample)
        self.assertIn("fio_group_rss_kb", output.read_text(encoding="utf-8"))

    def prepare_suite(self):
        plans = [runner.make_plan(c,"/dev/test",10**12,512) for c in self.cases[:2]]
        for i,p in enumerate(plans,1):
            folder = self.root/f"{i:03}_{p['case_id']}"
            folder.mkdir()
            runner.write_json(folder/"run_info.json",{"status":"pending","case_id":p["case_id"],
                                                      "case_index":i,"case_count":len(plans)})
        return plans

    def new_progress(self):
        progress = runner.RunLog(self.root)
        self.addCleanup(progress.close)
        return progress

    def test_suite_always_fail_fast(self):
        plans = self.prepare_suite()
        with patch.object(runner,"execute",side_effect=[(0,False),(0,False),(1,False),(0,False)]) as execute, \
             patch.object(runner,"probe_device",return_value={"path":"/dev/test"}), \
             patch.object(runner,"check_device_idle"), patch.object(runner,"check_cpus"):
            code = runner.run_suite(plans,self.root,"fio",60,{"/dev/test":{"path":"/dev/test"}},self.new_progress())
            self.assertEqual(code,1)
            self.assertEqual(execute.call_count,3)
        second = json.loads((self.root/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(second["status"],"skipped")
        self.assertIn("seq_read_1m",second["skip_reason"])

    def test_write_to_system_disk_is_skipped_and_later_case_runs(self):
        write_case = copy.deepcopy(next(c for c in self.cases if c["case_id"] == "precondition_write_84"))
        read_case = copy.deepcopy(self.cases[0])
        plans = [runner.make_plan(write_case, "/dev/test", 10**12, 512),
                 runner.make_plan(read_case, "/dev/test", 10**12, 512)]
        for index, plan in enumerate(plans, 1):
            folder = self.root / f"{index:03}_{plan['case_id']}"
            folder.mkdir()
            runner.write_json(folder / "run_info.json", {"status":"pending", "case_id":plan["case_id"]})
        device = {"path":"/dev/test"}
        with patch.object(runner, "execute", side_effect=[(0,False),(0,False),(0,False)]) as execute, \
             patch.object(runner, "probe_device", return_value=device), \
             patch.object(runner, "is_system_disk", return_value=True), \
             patch.object(runner, "check_device_idle"), patch.object(runner, "check_cpus"), \
             patch.object(runner, "extract_fio_json", return_value={}):
            self.assertEqual(runner.run_suite(plans, self.root, "fio", 60,
                                             {"/dev/test":device}, self.new_progress()), 0)
        self.assertEqual(execute.call_count, 3)  # two parse checks, then only the read workload
        skipped = json.loads((self.root / "001_precondition_write_84" / "run_info.json").read_text())
        passed = json.loads((self.root / "002_seq_read_1m" / "run_info.json").read_text())
        self.assertEqual(skipped["status"], "skipped")
        self.assertEqual(skipped["skip_code"], "system_disk")
        self.assertIn("write command was not executed", skipped["skip_reason"])
        self.assertEqual(passed["status"], "passed")

    def test_parse_failure_prevents_all_workloads(self):
        plans = self.prepare_suite()
        with patch.object(runner,"execute",side_effect=[(0,False),(1,False)]) as execute:
            with self.assertRaisesRegex(runner.ConfigError,"cannot parse"):
                runner.run_suite(plans,self.root,"fio",60,{"/dev/test":{"path":"/dev/test"}},self.new_progress())
            self.assertEqual(execute.call_count,2)
        first = json.loads((self.root/"001_seq_read_1m"/"run_info.json").read_text())
        second = json.loads((self.root/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(first["status"],"skipped")
        self.assertEqual(second["status"],"preflight_failed")
        self.assertEqual(second["failure_stage"],"PREFLIGHT")

    def test_interrupt_status_retained(self):
        plans = self.prepare_suite()
        with patch.object(runner,"execute",side_effect=[(0,False),(0,False),KeyboardInterrupt]), \
             patch.object(runner,"probe_device",return_value={"path":"/dev/test"}), \
             patch.object(runner,"check_device_idle"), patch.object(runner,"check_cpus"):
            with self.assertRaises(KeyboardInterrupt):
                runner.run_suite(plans,self.root,"fio",60,{"/dev/test":{"path":"/dev/test"}},self.new_progress())
        info = json.loads((self.root/"001_seq_read_1m"/"run_info.json").read_text())
        self.assertEqual(info["status"],"interrupted")
        second = json.loads((self.root/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(second["status"],"skipped")

    def test_main_logs_early_excel_failure(self):
        with self.assertRaises(FileNotFoundError):
            runner.main(["--excel",str(self.root/"missing.xlsx"),"--output",str(self.root/"logs")])
        archive = next((self.root/"logs").iterdir())
        log = (archive/"runner.log").read_text(encoding="utf-8")
        self.assertIn("CONFIG_READ",log)
        self.assertIn("Traceback",log)
        status = json.loads((archive/"run_status.json").read_text())
        self.assertEqual(status["status"],"failed")
        self.assertEqual(status["stage"],"CONFIG_READ")
        self.assertEqual(status["exit_code"],2)

    def test_main_logs_preview_progress_and_summary(self):
        runner.main(["--excel",str(SAMPLE),"--dry-run","--device-size-bytes","1000000000000",
                     "--output",str(self.root)])
        archive = next(self.root.iterdir())
        log = (archive/"runner.log").read_text(encoding="utf-8")
        for stage in ("STARTUP","CONFIG_READ","SELECT","DEVICE_CHECK","PLAN_GENERATE","ARCHIVE","PREVIEW_DONE"):
            self.assertIn(f"[{stage}]",log)
        self.assertIn("[case 2/2 id=randread_4k]",log)
        status = json.loads((archive/"run_status.json").read_text())
        self.assertEqual(status["status"],"preview_complete")
        self.assertEqual(status["counts"],{"planned":2})

    def test_archive_failure_records_current_case_as_error(self):
        def broken_prepare(args, root, progress):
            progress.select(self.cases[:2])
            progress.step("ARCHIVE","saving second case",2,"randread_4k")
            folder = root/"002_randread_4k"
            folder.mkdir()
            runner.write_json(folder/"run_info.json",dict(case_id="randread_4k",status="pending"))
            raise OSError("archive write failed")
        with patch.object(runner,"prepare_and_run",side_effect=broken_prepare):
            with self.assertRaisesRegex(OSError,"archive write failed"):
                runner.main(["--excel",str(SAMPLE),"--output",str(self.root)])
        archive = next(self.root.iterdir())
        status = json.loads((archive/"run_status.json").read_text())
        info = json.loads((archive/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(info["status"],"error")
        self.assertEqual(status["case_index"],2)
        self.assertEqual(status["counts"],{"skipped":1,"error":1})

    def test_second_case_failure_preserves_completed_first(self):
        plans = self.prepare_suite()
        progress = self.new_progress()
        with patch.object(runner,"execute",side_effect=[(0,False),(0,False),(0,False),(9,False)]), \
             patch.object(runner,"extract_fio_json",return_value={}), \
             patch.object(runner,"probe_device",return_value={"path":"/dev/test"}), \
             patch.object(runner,"check_device_idle"), patch.object(runner,"check_cpus"):
            code = runner.run_suite(plans,self.root,"fio",60,{"/dev/test":{"path":"/dev/test"}},progress)
        progress.finish("failed",code)
        status = json.loads((self.root/"run_status.json").read_text())
        self.assertEqual(status["case_index"],2)
        self.assertEqual(status["case_id"],"randread_4k")
        self.assertEqual(status["counts"],{"passed":1,"failed":1})
        log = (self.root/"runner.log").read_text(encoding="utf-8")
        self.assertIn("[CASE_RUN][case 2/2 id=randread_4k]",log)
        self.assertIn("exit_code=9",log)

    def test_timeout_stops_following_cases(self):
        plans = self.prepare_suite()
        with patch.object(runner,"execute",side_effect=[(0,False),(0,False),(-9,True)]) as execute, \
             patch.object(runner,"probe_device",return_value={"path":"/dev/test"}), \
             patch.object(runner,"check_device_idle"), patch.object(runner,"check_cpus"):
            self.assertEqual(runner.run_suite(plans,self.root,"fio",60,
                             {"/dev/test":{"path":"/dev/test"}},self.new_progress()),1)
        self.assertEqual(execute.call_count,3)
        first = json.loads((self.root/"001_seq_read_1m"/"run_info.json").read_text())
        second = json.loads((self.root/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(first["status"],"timeout")
        self.assertEqual(second["status"],"skipped")

    def test_preflight_spawn_error_stops_suite(self):
        plans = self.prepare_suite()
        with patch.object(runner,"execute",side_effect=OSError("spawn failed")) as execute:
            with self.assertRaisesRegex(OSError,"spawn failed"):
                runner.run_suite(plans,self.root,"fio",60,{},self.new_progress())
        self.assertEqual(execute.call_count,1)
        first = json.loads((self.root/"001_seq_read_1m"/"run_info.json").read_text())
        second = json.loads((self.root/"002_randread_4k"/"run_info.json").read_text())
        self.assertEqual(first["status"],"preflight_failed")
        self.assertEqual(second["status"],"skipped")

    def test_heartbeat_and_pid_are_logged(self):
        progress = self.new_progress()
        progress.step("CASE_RUN","test",1,"fake_case")
        process = unittest.mock.Mock(pid=123,returncode=0)
        process.wait.side_effect = [subprocess.TimeoutExpired("fake",10),0]
        with patch.object(runner.subprocess,"Popen",return_value=process), \
             patch.object(runner.time,"monotonic",side_effect=[0,0,10,10,10,11]):
            self.assertEqual(runner.execute(["fake"],self.root,self.root/"result.txt",60,progress),(0,False))
        log = (self.root/"runner.log").read_text(encoding="utf-8")
        self.assertIn("Process started: pid=123",log)
        self.assertIn("Process running: pid=123; elapsed=10.0s",log)
        self.assertIn("Process exited: pid=123; exit_code=0",log)

    def test_continue_on_error_is_rejected(self):
        with patch.object(sys,"stderr"), self.assertRaises(SystemExit) as error:
            runner.arguments(["--excel",str(SAMPLE),"--continue-on-error"])
        self.assertEqual(error.exception.code,2)


if __name__ == "__main__":
    unittest.main()
