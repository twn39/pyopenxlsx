#!/usr/bin/env python3
"""Compare two pytest-benchmark JSON exports.

Usage::

    uv run pytest tests/test_benchmark.py -m benchmark_fast \\
        --benchmark-only --benchmark-json=baseline.json
    # ... make changes ...
    uv run pytest tests/test_benchmark.py -m benchmark_fast \\
        --benchmark-only --benchmark-json=current.json
    uv run python scripts/compare_benchmarks.py baseline.json current.json

Exit code 1 if any compared mean regresses beyond --threshold (default 25%).
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict


def _load(path: Path) -> Dict[str, Dict[str, Any]]:
    data = json.loads(path.read_text(encoding="utf-8"))
    benchmarks = data.get("benchmarks") or []
    by_name: Dict[str, Dict[str, Any]] = {}
    for b in benchmarks:
        name = b.get("fullname") or b.get("fullname")
        if not name:
            continue
        stats = b.get("stats") or {}
        by_name[name] = {
            "mean": stats.get("mean"),
            "min": stats.get("min"),
            "rounds": stats.get("rounds"),
            "group": (b.get("group") or ""),
        }
    return by_name


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("baseline", type=Path, help="Baseline --benchmark-json file")
    parser.add_argument("current", type=Path, help="Current --benchmark-json file")
    parser.add_argument(
        "--threshold",
        type=float,
        default=0.25,
        help="Relative mean regression threshold (default 0.25 = 25%%)",
    )
    parser.add_argument(
        "--min-ms",
        type=float,
        default=1.0,
        help="Ignore cases with baseline mean below this many ms",
    )
    args = parser.parse_args(argv)

    base = _load(args.baseline)
    cur = _load(args.current)

    if not base or not cur:
        print("No benchmarks found in one or both files.", file=sys.stderr)
        return 2

    names = sorted(set(base) & set(cur))
    only_base = sorted(set(base) - set(cur))
    only_cur = sorted(set(cur) - set(base))

    print(f"{'name':<48} {'base_ms':>10} {'cur_ms':>10} {'delta%':>10}")
    print("-" * 82)

    regressions = []
    for name in names:
        b_mean = base[name]["mean"]
        c_mean = cur[name]["mean"]
        if b_mean is None or c_mean is None:
            continue
        b_ms = b_mean * 1000.0
        c_ms = c_mean * 1000.0
        if b_ms < args.min_ms:
            delta_pct = 0.0
            flag = ""
        else:
            delta_pct = (c_mean - b_mean) / b_mean * 100.0
            flag = "  << REGRESS" if (c_mean - b_mean) / b_mean > args.threshold else ""
            if flag:
                regressions.append((name, b_ms, c_ms, delta_pct))
        print(f"{name:<48} {b_ms:10.3f} {c_ms:10.3f} {delta_pct:9.1f}%{flag}")

    if only_base:
        print("\nOnly in baseline:", ", ".join(only_base))
    if only_cur:
        print("\nOnly in current:", ", ".join(only_cur))

    if regressions:
        print(f"\n{len(regressions)} regression(s) beyond {args.threshold:.0%}:")
        for name, b_ms, c_ms, pct in regressions:
            print(f"  - {name}: {b_ms:.3f}ms -> {c_ms:.3f}ms ({pct:+.1f}%)")
        return 1

    print("\nNo regressions beyond threshold.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
