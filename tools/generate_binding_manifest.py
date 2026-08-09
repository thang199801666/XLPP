#!/usr/bin/env python3
"""Generate/check a high-level binding manifest from native public headers.

The manifest is intentionally source-derived: every public Workbook/Worksheet
method name and overload count is captured from the current C++ headers.  It
also records heuristic Python/C# exposure so reviewers can see binding drift.
The checked-in manifest makes any native public-surface change an explicit
review event even when the method is intentionally C++-only.
"""
from __future__ import annotations

import argparse
import difflib
import json
import re
import sys
from pathlib import Path

CLASSES = {
    "Workbook": "include/XLPP/Workbook/Workbook.h",
    "Worksheet": "include/XLPP/Worksheet/Worksheet.h",
}


def find_class_body(text: str, class_name: str) -> str:
    m = re.search(rf"(?m)^\s*class\s+{re.escape(class_name)}\s*{{", text)
    if not m:
        raise RuntimeError(f"class {class_name} not found")
    start = m.end()
    depth = 1
    i = start
    while i < len(text) and depth:
        if text.startswith("//", i):
            nl = text.find("\n", i + 2)
            i = len(text) if nl < 0 else nl + 1
            continue
        if text.startswith("/*", i):
            end = text.find("*/", i + 2)
            i = len(text) if end < 0 else end + 2
            continue
        if text[i] in ('"', "'"):
            quote = text[i]
            i += 1
            while i < len(text):
                if text[i] == "\\":
                    i += 2
                    continue
                if text[i] == quote:
                    i += 1
                    break
                i += 1
            continue
        if text[i] == "{":
            depth += 1
        elif text[i] == "}":
            depth -= 1
        i += 1
    if depth:
        raise RuntimeError(f"unterminated class {class_name}")
    return text[start : i - 1]


def public_chunks(body: str) -> str:
    # Keep only class-depth-zero public declarations. Inline method bodies are
    # retained so overloads declared/defined in the header are both visible.
    visibility = "private"
    depth = 0
    out: list[str] = []
    token = []
    i = 0
    while i < len(body):
        if body.startswith("//", i):
            nl = body.find("\n", i + 2)
            i = len(body) if nl < 0 else nl + 1
            continue
        if body.startswith("/*", i):
            end = body.find("*/", i + 2)
            i = len(body) if end < 0 else end + 2
            continue
        ch = body[i]
        if depth == 0:
            vm = re.match(r"\s*(public|private|protected)\s*:\s*", body[i:])
            if vm:
                visibility = vm.group(1)
                i += vm.end()
                token.clear()
                continue
        if ch == "{":
            if visibility == "public" and depth == 0:
                out.append("".join(token))
                out.append(" {")
                token.clear()
            depth += 1
        elif ch == "}":
            depth -= 1
            if visibility == "public" and depth == 0:
                out.append(" }\n")
        elif depth == 0:
            if visibility == "public":
                token.append(ch)
                if ch == ";":
                    out.append("".join(token))
                    out.append("\n")
                    token.clear()
            else:
                token.clear()
        i += 1
    return "".join(out)


def extract_methods(header: Path, class_name: str) -> dict[str, int]:
    body = find_class_body(header.read_text(encoding="utf-8"), class_name)
    public = public_chunks(body)
    # Collapse inline bodies; declaration prefixes remain before the first '{'.
    decls: list[str] = []
    for line in public.splitlines():
        candidate = line.split("{", 1)[0].strip()
        if candidate:
            decls.append(candidate)
    text = "\n".join(decls)
    counts: dict[str, int] = {}
    # Match callable identifiers immediately before '('. Operators/ctors are
    # excluded because high-level bindings expose them via idiomatic protocols.
    for m in re.finditer(r"\b([A-Za-z_][A-Za-z0-9_]*)\s*\(", text):
        name = m.group(1)
        if name == class_name or name.startswith("operator"):
            continue
        # Ignore control-flow fragments accidentally retained from inline bodies.
        if name in {"if", "for", "while", "switch", "return", "sizeof", "static_cast", "move"}:
            continue
        counts[name] = counts.get(name, 0) + 1
    return dict(sorted(counts.items()))


def snake(name: str) -> str:
    s1 = re.sub(r"(.)([A-Z][a-z]+)", r"\1_\2", name)
    s2 = re.sub(r"([a-z0-9])([A-Z])", r"\1_\2", s1)
    return s2.lower()


def pascal(name: str) -> str:
    return name[:1].upper() + name[1:]


PYTHON_ALIASES = {
    ("Workbook", "worksheet"): ["worksheet", "__getitem__"],
    ("Workbook", "sheetNames"): ["sheet_names"],
    ("Workbook", "sheetCount"): ["sheet_count"],
    ("Workbook", "namedStyle"): ["named_style"],
    ("Workbook", "namedStyles"): ["named_styles"],
    ("Workbook", "definedName"): ["defined_name"],
    ("Workbook", "definedNames"): ["defined_names"],
    ("Workbook", "customProperties"): ["custom_properties"],
    ("Workbook", "calcProperties"): ["calc_properties"],
    ("Workbook", "date1904"): ["date1904", "date_1904"],
    ("Workbook", "setDate1904"): ["date1904", "date_1904"],
    ("Workbook", "preservedRelationships"): ["preserved_relationships"],
    ("Workbook", "worksheets"): ["worksheets", "__iter__"],
    ("Worksheet", "cell"): ["cell", "__getitem__"],
    ("Worksheet", "append"): ["append", "append_strings"],
    ("Worksheet", "charts"): ["charts", "chart_count"],
    ("Worksheet", "images"): ["images", "image_count"],
    ("Worksheet", "tables"): ["tables", "table_count"],
    ("Worksheet", "pivotTables"): ["pivot_tables", "pivot_count"],
    ("Worksheet", "rename"): ["rename"],
    ("Worksheet", "setPrintArea"): ["set_print_area"],
    ("Worksheet", "setSheetView"): ["set_sheet_view"],
    ("Worksheet", "tryCell"): ["try_cell", "cell"],
    ("Worksheet", "name_"): ["name"],
    ("Worksheet", "mergedRanges"): ["merged_ranges"],
    ("Worksheet", "rowDimensions"): ["row_dimensions"],
    ("Worksheet", "columnDimensions"): ["column_dimensions"],
}

CSHARP_ALIASES = {
    ("Workbook", "worksheet"): ["GetWorksheet"],
    ("Workbook", "sheetNames"): ["SheetNames"],
    ("Workbook", "sheetCount"): ["SheetCount"],
    ("Workbook", "namedStyle"): ["GetNamedStyle"],
    ("Workbook", "definedName"): ["GetDefinedName"],
    ("Workbook", "definedNames"): ["DefinedNames"],
    ("Workbook", "namedStyles"): ["NamedStyles"],
    ("Workbook", "preservedRelationships"): ["PreservedRelationships"],
    ("Workbook", "worksheets"): ["Worksheets"],
    ("Workbook", "index"): ["IndexOf"],
    ("Workbook", "customProperties"): ["CustomProperties"],
    ("Workbook", "date1904"): ["Date1904"],
    ("Workbook", "setDate1904"): ["Date1904"],
    ("Worksheet", "cell"): ["Cell"],
    ("Worksheet", "append"): ["AppendRow"],
    ("Worksheet", "charts"): ["Charts"],
    ("Worksheet", "images"): ["Images"],
    ("Worksheet", "tables"): ["Tables"],
    ("Worksheet", "pivotTables"): ["PivotTables"],
    ("Worksheet", "rename"): ["Rename"],
    ("Worksheet", "setPrintArea"): ["PrintArea"],
    ("Worksheet", "setPrintTitlesRows"): ["PrintTitlesRows"],
    ("Worksheet", "setPrintTitlesCols"): ["PrintTitlesCols"],
    ("Worksheet", "rowDimensions"): ["RowDimension"],
    ("Worksheet", "columnDimensions"): ["ColumnDimension"],
    ("Worksheet", "mergedRanges"): ["MergedRanges"],
}


def py_symbols(text: str) -> set[str]:
    return set(re.findall(r"\.def(?:_property(?:_readonly)?)?\(\s*\"([^\"]+)\"", text))


def cs_symbols(text: str) -> set[str]:
    symbols = set(re.findall(r"\bpublic\s+(?:static\s+)?(?:[\w<>,?\[\].]+\s+)+([A-Za-z_][A-Za-z0-9_]*)\s*(?:\(|\{)", text))
    symbols.update(re.findall(r"\bpublic\s+[\w<>,?\[\].]+\s+([A-Za-z_][A-Za-z0-9_]*)\s*=>", text))
    return symbols


def generate(root: Path) -> dict:
    py_text = (root / "bindings/python/src/xlpp_bindings.cpp").read_text(encoding="utf-8")
    cs_text = "\n".join(
        (root / rel).read_text(encoding="utf-8")
        for rel in ("bindings/csharp/XlppNet.cs", "bindings/csharp/XlppNet.Advanced.cs")
    )
    pys = py_symbols(py_text)
    css = cs_symbols(cs_text)
    version = (root / "VERSION").read_text(encoding="utf-8").strip()
    result = {
        "schema": 1,
        "version": version,
        "native_sources": CLASSES,
        "classes": {},
    }
    for class_name, rel in CLASSES.items():
        methods = extract_methods(root / rel, class_name)
        entries = {}
        for name, overloads in methods.items():
            py_candidates = PYTHON_ALIASES.get((class_name, name), [snake(name)])
            cs_candidates = CSHARP_ALIASES.get((class_name, name), [pascal(name)])
            py_found = sorted(set(py_candidates) & pys)
            cs_found = sorted(set(cs_candidates) & css)
            entries[name] = {
                "overloads": overloads,
                "python_candidates": py_candidates,
                "python_found": py_found,
                "csharp_candidates": cs_candidates,
                "csharp_found": cs_found,
            }
        result["classes"][class_name] = entries
    return result


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", type=Path, default=Path(__file__).resolve().parents[1])
    ap.add_argument("--output", type=Path)
    ap.add_argument("--check", type=Path)
    args = ap.parse_args()
    root = args.root.resolve()
    manifest = generate(root)
    rendered = json.dumps(manifest, indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(rendered, encoding="utf-8")
    if args.check:
        expected = args.check.read_text(encoding="utf-8") if args.check.exists() else ""
        if expected != rendered:
            sys.stderr.write("Binding manifest is stale. Regenerate with tools/generate_binding_manifest.py.\n")
            diff = difflib.unified_diff(expected.splitlines(True), rendered.splitlines(True), fromfile=str(args.check), tofile="generated")
            sys.stderr.writelines(diff)
            return 2
    if not args.output and not args.check:
        sys.stdout.write(rendered)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
