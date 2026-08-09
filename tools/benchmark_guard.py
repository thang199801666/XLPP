#!/usr/bin/env python3
"""Reject benchmark regressions using same-machine median JSON results.

Input format: {"scenario": milliseconds, ...}. This intentionally compares a
candidate run with an explicit baseline rather than machine-dependent absolute
numbers checked into source control.
"""
import argparse, json, sys
from pathlib import Path

p = argparse.ArgumentParser()
p.add_argument("baseline", type=Path)
p.add_argument("candidate", type=Path)
p.add_argument("--max-regression", type=float, default=15.0)
a = p.parse_args()
base = json.loads(a.baseline.read_text())
cand = json.loads(a.candidate.read_text())
failures = []
for key, old in base.items():
    if key not in cand or not isinstance(old, (int, float)) or old <= 0:
        continue
    new = cand[key]
    if not isinstance(new, (int, float)):
        failures.append(f"{key}: candidate is not numeric")
        continue
    delta = (new / old - 1.0) * 100.0
    print(f"{key}: {old:.3f} -> {new:.3f} ms ({delta:+.2f}%)")
    if delta > a.max_regression:
        failures.append(f"{key}: +{delta:.2f}% > +{a.max_regression:.2f}%")
if failures:
    print("PERFORMANCE REGRESSION:\n  " + "\n  ".join(failures), file=sys.stderr)
    raise SystemExit(2)
print("performance guard: PASS")
