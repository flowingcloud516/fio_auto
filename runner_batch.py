#!/usr/bin/env python3
"""Run one named fio Excel suite and immediately build its result report."""
from __future__ import annotations

import argparse
from pathlib import Path
import shlex
import subprocess
import sys


VERSION = "1.0.0"
RESERVED_RUNNER_OPTIONS = {"--excel", "--output"}


def parser():
    value = argparse.ArgumentParser(
        description=("Run fio_excel_runner.py and fio_result_report.py using paths derived "
                     "from one case_name. Unrecognized options are passed to the fio runner."))
    value.add_argument(
        "case_name",
        help=("suite name or path prefix: NAME.xlsx -> NAME/ -> NAME_result.xlsx; "
              "an optional .xlsx suffix is accepted"),
    )
    value.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    return value


def suite_paths(case_name):
    base = Path(case_name)
    if base.suffix.lower() == ".xlsx":
        base = base.with_suffix("")
    if not base.name or base.name in {".", ".."}:
        raise ValueError("case_name must include a file name")
    excel = Path(f"{base}.xlsx").resolve()
    results = base.resolve()
    report = Path(f"{base}_result.xlsx").resolve()
    return excel, results, report


def reject_reserved_options(value, runner_args):
    for token in runner_args:
        option = token.split("=", 1)[0]
        if option in RESERVED_RUNNER_OPTIONS:
            value.error(f"{option} is derived from case_name and cannot be overridden")


def show_command(stage, command):
    print(f"[batch] {stage}: {shlex.join(command)}", flush=True)


def main(argv=None):
    if sys.version_info < (3, 11):
        print("ERROR: Python 3.11 or newer is required", file=sys.stderr)
        return 2
    value = parser()
    args, runner_args = value.parse_known_args(argv)
    reject_reserved_options(value, runner_args)
    try:
        excel, results, report = suite_paths(args.case_name)
    except ValueError as exc:
        value.error(str(exc))

    if not excel.is_file():
        print(f"ERROR: input Excel does not exist: {excel}", file=sys.stderr)
        return 2

    script_dir = Path(__file__).resolve().parent
    runner_script = script_dir / "fio_excel_runner.py"
    report_script = script_dir / "fio_result_report.py"
    for script in (runner_script, report_script):
        if not script.is_file():
            print(f"ERROR: required script does not exist: {script}", file=sys.stderr)
            return 2

    print(f"[batch] case_name={args.case_name}", flush=True)
    print(f"[batch] Excel={excel}", flush=True)
    print(f"[batch] results={results}", flush=True)
    print(f"[batch] report={report}", flush=True)

    runner_command = [sys.executable, str(runner_script), "--excel", str(excel),
                      "--output", str(results), *runner_args]
    show_command("1/2 run fio suite", runner_command)
    runner_code = subprocess.run(runner_command, check=False).returncode
    if runner_code:
        print(f"[batch] fio runner exited with code {runner_code}; "
              "continuing to summarize retained results", file=sys.stderr, flush=True)

    report_command = [sys.executable, str(report_script), "--excel", str(excel),
                      "--results", str(results), "--output", str(report)]
    show_command("2/2 build result report", report_command)
    report_code = subprocess.run(report_command, check=False).returncode

    if report_code:
        print(f"ERROR: result reporter exited with code {report_code}", file=sys.stderr)
    else:
        print(f"[batch] report ready: {report}", flush=True)
    if runner_code:
        return runner_code
    return report_code


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("Interrupted. Existing output files were retained.", file=sys.stderr)
        raise SystemExit(130)
