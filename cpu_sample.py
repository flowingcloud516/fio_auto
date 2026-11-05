#!/usr/bin/env python3
"""Sample Linux CPU execution time without mpstat or third-party packages.

Default: all CPUs visible in /proc/stat, including CPUs outside process
affinity. CPU IDs are logical CPUs, not physical cores. ``selected_sum`` is
the aggregate of selected CPUs. ``core_equiv`` is busy CPU seconds divided by
wall seconds.
"""

import argparse
import csv
import datetime
import math
import os
import sys
import time

FIELDS = ("user", "nice", "system", "idle", "iowait", "irq", "softirq", "steal")
BUSY = (0, 1, 2, 5, 6)


def cpu_ids(value):
    selected = set()
    for part in value.split(","):
        bounds = part.strip().split("-")
        if not 1 <= len(bounds) <= 2:
            raise ValueError("Expected CPU IDs such as 0-7,16-23")
        start = int(bounds[0])
        end = int(bounds[-1])
        if start < 0 or end < start:
            raise ValueError("CPU IDs must be nonnegative ascending ranges")
        selected.update(range(start, end + 1))
    return selected


def snapshot():
    result = {}
    with open("/proc/stat") as source:
        for line in source:
            parts = line.split()
            if parts and parts[0].startswith("cpu") and parts[0][3:].isdigit():
                values = [int(x) for x in parts[1:9]]
                result[int(parts[0][3:])] = (values + [0] * 8)[:8]
    return result


def metrics(delta, interval, hz):
    total = sum(delta)
    busy = sum(delta[index] for index in BUSY)
    output = {field + "_pct": round(100.0 * delta[index] / total, 3)
              for index, field in enumerate(FIELDS)}
    output["busy_pct"] = round(100.0 * busy / total, 3)
    output["busy_cpu_s"] = round(busy / hz, 6)
    output["core_equiv"] = round(busy / hz / interval, 6)
    return output


def sample_rows(previous, current, selected, interval, hz, stamp, elapsed):
    rows = []
    sums = [0] * 8
    valid_count = 0
    for cpu in sorted(selected):
        row = {"timestamp_utc": stamp, "elapsed_s": round(elapsed, 6),
               "interval_s": round(interval, 6), "cpu": "cpu" + str(cpu),
               "selected_cpus": len(selected), "valid_cpus": 0, "status": "ok"}
        if cpu not in previous or cpu not in current:
            row["status"] = "missing_or_hotplug"
        else:
            delta = [after - before for before, after in zip(previous[cpu], current[cpu])]
            if any(value < 0 for value in delta):
                row["status"] = "counter_decreased"
            elif sum(delta) <= 0:
                row["status"] = "no_ticks"
            else:
                row.update(metrics(delta, interval, hz))
                row["valid_cpus"] = 1
                valid_count += 1
                sums = [left + right for left, right in zip(sums, delta)]
        rows.append(row)
    aggregate = {"timestamp_utc": stamp, "elapsed_s": round(elapsed, 6),
                 "interval_s": round(interval, 6), "cpu": "selected_sum",
                 "selected_cpus": len(selected), "valid_cpus": valid_count,
                 "status": "ok" if valid_count == len(selected) else "partial"}
    if valid_count:
        aggregate.update(metrics(sums, interval, hz))
    return rows + [aggregate]


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--cpus", help="Logical CPU IDs; default all in /proc/stat")
    parser.add_argument("--interval", type=float, default=1)
    parser.add_argument("--duration", type=float, default=0,
                        help="Seconds; 0 means until Ctrl+C")
    parser.add_argument("--output", default="cpu.csv")
    args = parser.parse_args(argv)
    if (not math.isfinite(args.interval) or args.interval <= 0 or
            not math.isfinite(args.duration) or args.duration < 0):
        parser.error("interval must be finite and >0; duration finite and >=0")
    try:
        selected = cpu_ids(args.cpus) if args.cpus else None
    except ValueError as error:
        parser.error(str(error))
    hz = os.sysconf("SC_CLK_TCK")
    columns = ["timestamp_utc", "elapsed_s", "interval_s", "cpu", "selected_cpus",
               "valid_cpus", "status", "busy_pct"]
    columns += [field + "_pct" for field in FIELDS]
    columns += ["busy_cpu_s", "core_equiv"]
    previous = snapshot()
    started = last = time.monotonic()
    deadline = started + args.duration if args.duration else None
    with open(args.output, "w", newline="") as target:
        writer = csv.DictWriter(target, fieldnames=columns)
        writer.writeheader()
        try:
            while True:
                now = time.monotonic()
                if deadline is not None and now >= deadline:
                    break
                delay = args.interval if deadline is None else min(args.interval, deadline - now)
                time.sleep(delay)
                current = snapshot()
                now = time.monotonic()
                active = selected if selected is not None else set(previous) | set(current)
                stamp = datetime.datetime.now(datetime.timezone.utc).isoformat()
                writer.writerows(sample_rows(previous, current, active, now-last,
                                             hz, stamp, now-started))
                target.flush()
                previous, last = current, now
        except KeyboardInterrupt:
            pass
    print("Saved " + args.output, file=sys.stderr)


if __name__ == "__main__":
    main()
