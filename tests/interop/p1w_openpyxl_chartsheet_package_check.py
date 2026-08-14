#!/usr/bin/env python3
"""P1W external-host/package gate for Chartsheet auxiliary ownership.

Usage:
  python p1w_openpyxl_chartsheet_package_check.py <printer.xltx> <legacy-hf.xltx> <printer-regenerated.xltx>
"""
from pathlib import Path
import sys
import zipfile
from openpyxl import load_workbook


def require(cond, label):
    if not cond:
        raise SystemExit(f"P1W openpyxl host check failed: {label}")


def chartsheet_xml_name(zf: zipfile.ZipFile):
    names = sorted(n for n in zf.namelist() if n.startswith("xl/chartsheets/sheet") and n.endswith(".xml"))
    require(len(names) == 1, "exactly one Chartsheet XML part")
    return names[0]


def rels_name(part: str):
    directory, name = part.rsplit("/", 1)
    return f"{directory}/_rels/{name}.rels"


def check_openpyxl(path: Path):
    wb = load_workbook(path)
    require(wb.template is True, f"template identity for {path.name}")
    require(len(wb.chartsheets) == 1, f"Chartsheet discovery for {path.name}")
    require(len(wb.chartsheets[0]._charts) == 1, f"chart discovery for {path.name}")
    wb.close()


def check_printer(path: Path):
    check_openpyxl(path)
    with zipfile.ZipFile(path) as zf:
        printer = [n for n in zf.namelist() if n.startswith("xl/printerSettings/printerSettings") and n.endswith(".bin")]
        require(len(printer) == 1, "printerSettings binary part")
        sheet = chartsheet_xml_name(zf)
        xml = zf.read(sheet).decode("utf-8")
        rels = zf.read(rels_name(sheet)).decode("utf-8")
        require("<pageSetup" in xml and "r:id=" in xml, "pageSetup printer relationship owner")
        require("/printerSettings" in rels, "printerSettings relationship type")
        require("spreadsheetml.printerSettings" in zf.read("[Content_Types].xml").decode("utf-8"),
                "printerSettings content type")


def check_legacy_hf(path: Path):
    check_openpyxl(path)
    with zipfile.ZipFile(path) as zf:
        sheet = chartsheet_xml_name(zf)
        xml = zf.read(sheet).decode("utf-8")
        rels = zf.read(rels_name(sheet)).decode("utf-8")
        require("<legacyDrawingHF" in xml and "rIdHF" in xml, "legacyDrawingHF owner survives regeneration")
        require("/vmlDrawing" in rels and "rIdHF" in rels, "legacyDrawingHF relationship survives regeneration")
        require("xl/drawings/vmlDrawingHF1.vml" in zf.namelist(), "header/footer VML part survives regeneration")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        raise SystemExit(__doc__)
    check_printer(Path(sys.argv[1]))
    check_legacy_hf(Path(sys.argv[2]))
    check_printer(Path(sys.argv[3]))
    print("P1W openpyxl Chartsheet-package host check: PASS")
