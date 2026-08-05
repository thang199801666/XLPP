import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET

root = pathlib.Path(__file__).resolve().parents[1]
version = (root / "VERSION").read_text(encoding="utf-8").strip()

checks = {
    "CMakeLists.txt": [f"VERSION {version}"],
    "vcpkg.json": [f'"version": "{version}"'],
    # setup.py reads the canonical VERSION file at build time.
    "setup.py": ["version=package_version"],
    "bindings/python/setup.py": ["version=package_version"],
    "bindings/c/xlpp_capi.cpp": [f'return "{version}";'],
    "bindings/csharp/XlppNet.csproj": [f"<Version>{version}</Version>"],
}

errors = []
for relative, expected_values in checks.items():
    path = root / relative
    content = path.read_text(encoding="utf-8")
    if not any(expected in content for expected in expected_values):
        errors.append(f"{relative}: expected one of {expected_values!r}")

if errors:
    print("Version mismatch:")
    print("\n".join(errors))
    sys.exit(1)

print(f"XLPP version metadata is consistent: {version}")
