import json
import pathlib
import re
import sys
import xml.etree.ElementTree as ET

root = pathlib.Path(__file__).resolve().parents[1]
version = (root / "VERSION").read_text(encoding="utf-8").strip()

checks = {
    "CMakeLists.txt": f"VERSION {version}",
    "vcpkg.json": f'"version": "{version}"',
    "setup.py": f'version="{version}"',
    "bindings/python/setup.py": f'version="{version}"',
    "bindings/c/xlpp_capi.cpp": f'return "{version}";',
    "bindings/csharp/XlppNet.csproj": f"<Version>{version}</Version>",
}

errors = []
for relative, expected in checks.items():
    path = root / relative
    if expected not in path.read_text(encoding="utf-8"):
        errors.append(f"{relative}: expected {expected!r}")

if errors:
    print("Version mismatch:")
    print("\n".join(errors))
    sys.exit(1)

print(f"XLPP version metadata is consistent: {version}")
