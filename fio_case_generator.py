#!/usr/bin/env python3
"""Generate a fio runner input workbook from Cartesian-product parameters."""
from __future__ import annotations

import argparse
import itertools
import math
from pathlib import Path
import re
import sys

from fio_excel_runner import CASE_COLUMNS, JOB_COLUMNS, ConfigError, RW, block_size, integer, load_cases
from fio_result_report import write_xlsx

VERSION = "1.0.0"


def csv_values(value, name):
    values = [item.strip() for item in value.split(",") if item.strip()]
    if not values:
        raise ConfigError(f"--{name} must contain at least one value")
    return values


def int_values(value, name, minimum=1, maximum=None):
    return [integer(item, f"--{name}", minimum=minimum, maximum=maximum)
            for item in csv_values(value, name)]


def bs_values(value):
    values = csv_values(value, "bs")
    for item in values:
        block_size(item, "--bs")
    return values


def rw_values(value):
    values = [item.lower() for item in csv_values(value, "rw")]
    unknown = [item for item in values if item not in RW]
    if unknown:
        raise ConfigError(f"--rw contains unsupported values: {unknown}; expected {sorted(RW)}")
    return values


def token(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).lower()) or "x"


def make_case_id(prefix, index, rw, bs, depth, jobs, numjobs, runtime):
    descriptive = f"{token(prefix)}_{index:04}_{token(rw)}_{token(bs)}_qd{depth}_j{jobs}_n{numjobs}_t{runtime}"
    return descriptive[:64].rstrip("_-")


def generate_rows(args):
    rws = rw_values(args.rw)
    sizes = bs_values(args.bs)
    depths = int_values(args.iodepth, "iodepth", maximum=65536)
    jobs_values = int_values(args.jobs, "jobs", maximum=65536)
    numjobs_values = int_values(args.numjobs, "numjobs", maximum=65536)
    runtimes = int_values(args.runtime, "runtime")
    time_based = bool(args.time_based)
    timeout_s = integer(args.timeout, "--timeout") if args.timeout is not None else None
    if not time_based:
        if args.runtime != "60":
            raise ConfigError("--runtime is not used when --time-based=0; omit it")
        if timeout_s is None:
            raise ConfigError("--timeout is required when --time-based=0")
    cpu_start = integer(args.cpu_start, "--cpu-start", minimum=0, maximum=65535)
    cpu_end = integer(args.cpu_end, "--cpu-end", minimum=cpu_start, maximum=65535)
    cpu_count = cpu_end - cpu_start + 1
    if any(n > cpu_count for n in numjobs_values):
        raise ConfigError(f"Each numjobs value must be <= CPU range size {cpu_count}")
    dimensions = (rws, sizes, depths, jobs_values, numjobs_values, runtimes)
    count = math.prod(len(values) for values in dimensions)
    if count > args.max_cases:
        raise ConfigError(f"Cartesian product creates {count} cases, exceeding --max-cases={args.max_cases}")
    rows = []
    for index, (rw, bs, depth, jobs, numjobs, runtime) in enumerate(itertools.product(*dimensions), 1):
        if jobs * numjobs > 65536:
            raise ConfigError(f"jobs={jobs} × numjobs={numjobs} exceeds the 65536 total-worker limit")
        mixed = rw in {"rw", "randrw"}
        runtime_value = runtime if time_based else ""
        runtime_id = runtime if time_based else "full"
        rows.append([
            make_case_id(args.case_prefix, index, rw, bs, depth, jobs, numjobs, runtime_id),
            1, args.device, rw, bs, depth, jobs, numjobs, cpu_start, cpu_end, runtime_value,
            args.rwmixread if mixed else "", args.notes, 1 if time_based else 0, timeout_s or "",
        ])
    return rows


def arguments(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rw", required=True, help="comma-separated values, e.g. read,write,randread")
    parser.add_argument("--bs", required=True, help="comma-separated values, e.g. 8K,128K,1M")
    parser.add_argument("--iodepth", required=True, help="comma-separated values, e.g. 1,2,4,8")
    parser.add_argument("--device", required=True, help="target block device written to Cases.device")
    parser.add_argument("--cpu-start", required=True, type=int)
    parser.add_argument("--cpu-end", required=True, type=int)
    parser.add_argument("--jobs", default="1", help="comma-separated independent job counts; default 1")
    parser.add_argument("--numjobs", default="1", help="comma-separated clones per job; default 1")
    parser.add_argument("--runtime", default="60", help="comma-separated runtimes in seconds; default 60")
    parser.add_argument("--time-based", type=int, choices=(0, 1), default=1,
                        help="1: run for runtime seconds; 0: complete the range once")
    parser.add_argument("--timeout", type=int,
                        help="per-case watchdog seconds; required for --time-based=0")
    parser.add_argument("--rwmixread", type=int, default=50, help="read percentage for rw/randrw; default 50")
    parser.add_argument("--case-prefix", default="auto", help="ASCII prefix used in generated case_id")
    parser.add_argument("--notes", default="generated Cartesian product")
    parser.add_argument("--max-cases", type=int, default=10000, help="safety limit; default 10000")
    parser.add_argument("--output", type=Path, default=Path("生成用例.xlsx"))
    return parser.parse_args(argv)


def main(argv=None):
    if sys.version_info < (3, 11):
        raise ConfigError("Python 3.11 or newer is required")
    args = arguments(argv)
    if not re.fullmatch(r"[A-Za-z0-9_-]+", args.case_prefix):
        raise ConfigError("--case-prefix must use ASCII letters, digits, '_' or '-'")
    if not 0 <= args.rwmixread <= 100:
        raise ConfigError("--rwmixread must be within 0..100")
    if args.max_cases < 1:
        raise ConfigError("--max-cases must be positive")
    rows = generate_rows(args)
    write_xlsx(args.output, [("Cases", list(CASE_COLUMNS), rows),
                             ("JobOverrides", list(JOB_COLUMNS), [])])
    # Validate the generated file with exactly the same parser used by the runner.
    cases = load_cases(args.output)
    if len(cases) != len(rows):
        raise ConfigError("Generated workbook validation count mismatch")
    print(f"Generated {len(rows)} cases: {args.output.resolve()}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
