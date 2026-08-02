# XL++

[![Windows CI](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/windows-ci.yml)
[![Benchmark](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml/badge.svg)](https://github.com/thang199801666/XLPP/actions/workflows/benchmark.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![C++20](https://img.shields.io/badge/C%2B%2B-20-blue)](https://en.cppreference.com/w/cpp/20)

XL++ is a dependency-light C++20 `.xlsx` read/write library targeting an API breadth comparable to Python openpyxl. It uses the C++ standard library and zlib only.

## Current milestone
Milestone 20 completes core compatibility: package-part preservation (unknown parts survive load/save round-trips), strict/transitional OOXML namespace emission (`SaveOptions::strictNamespace`), malformed-file diagnostics and lenient recovery (`Workbook::load(path, LoadOptions)`), ZIP64 read/write with large-entry support, and cancellation/progress/limits on both `open` and `save`. See `ROADMAP.md` and `MILESTONE_20.md`.

Open `XL++.sln` with Visual Studio 2026, select x64 Debug or Release, then build with Platform Toolset v145.

## Project organization

XL++ now uses matching physical folders and Visual Studio filters. See `PROJECT_STRUCTURE.md`. Existing root-level public includes remain available through forwarding headers.


## Core styles example

```cpp
auto& cell = sheet.cell("A1");
cell.setValue(1234.5);
cell.font().setBold(true);
cell.font().color().setArgb("FFFF0000");
cell.fill().setPatternType("solid");
cell.fill().foregroundColor().setArgb("FFFFFF00");
cell.border().bottom().setStyle("thin");
cell.alignment().setHorizontal("center");
cell.setNumberFormat("#,##0.00");
```
