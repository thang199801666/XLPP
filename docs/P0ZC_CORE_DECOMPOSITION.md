# P0Z-C — Fine-Grained Core Decomposition / v1.6.0

P0Z-C continues the architecture-only refactor without changing the public semantic feature baseline. The scope-defined editing-core score remains 90.7/100. This milestone removes the two remaining build-time monoliths identified in P0Z-B, sharpens package/model dependency direction, and makes the native build graph match the physical source tree at a finer granularity.

## Chart ownership

The former `ChartWriterPatcher` responsibility is now separated into:

- `ChartSerializer.*` — generated ChartML serialization;
- `ChartFormatCodec.*` — fill/line/marker format primitives;
- `ChartMutationSupport.*` — low-level XML mutation primitives;
- `ImportedChartXmlEditor.*` — pure preservation-aware ChartML semantic mutation;
- `ImportedChartPatcher.*` — package/drawing/relationship orchestration only;
- `ChartReader.*` — imported ChartML semantic read path;
- `ChartXmlSupport.*` — shared ChartML navigation support.

`ImportedChartPatcher.cpp` is approximately 175 lines and no longer knows how individual axis, label, trendline or series edits are represented in XML. Package ownership and ChartML mutation are therefore independently testable/fuzzable boundaries.

## Worksheet model ownership

The former ~2.3k-line `Model/Worksheet/Worksheet.cpp` is split into:

- `Worksheet.cpp` — cell/range/row/column/table/core model operations (~264 lines);
- `WorksheetStructural.cpp` — row/column structural/reference rewrites (~770 lines);
- `WorksheetImages.cpp` — imported image lifecycle (~184 lines);
- `WorksheetCharts.cpp` — chart-level/axis/layout mutations (~601 lines);
- `WorksheetChartSeries.cpp` — series/data-label/trendline/error-bar mutations (~480 lines);
- `WorksheetChartValidation.*` — reusable chart edit validation;
- `WorksheetReferenceSupport.*` — reusable worksheet-range parsing.

This prevents edits to Chart/Image behavior from recompiling unrelated cell/range model implementation.

## Package/model decoupling

`PreservedPart` and `PreservedRelationship` are now declared in the package-neutral public header `XLPP/Package/Preservation.h`. `Package/Opc/RelationshipGraph` and generic `Preservation` code therefore no longer include `Workbook.h` merely to obtain package metadata types.

A new `ArchitectureBoundaryTests` CTest gate enforces two rules:

1. `Package/` must not include Workbook, Worksheet, Chart or Pivot semantic headers.
2. `Model/` must not include OOXML or Package implementation headers.

The gate is intentionally cheap and runs on every native CTest invocation.

## Fine-grained CMake graph

P0Z-C compiles the core through 27 private object modules before aggregation into the same `xlpp_static` / `XLPP::xlpp` target. In particular:

- Workbook and Worksheet model objects are separate;
- ZIP, OPC and XML package objects are separate;
- Chart support, reader, serializer, pure XML editor and package patcher are separate;
- OOXML Workbook/Worksheet/Drawings/Pivot/Styles/Tables/Comments remain separate domains.

This does not alter consumer linkage or ABI; it improves parallel/incremental compilation and exposes dependency ownership.

## Native regression layout

The former ~7.6k-line `RegressionTests.cpp` no longer exists. Regression coverage is split into:

- `PackageStreamingTests.cpp`;
- `WorkbookTests.cpp`;
- `DrawingsTests.cpp`;
- `ChartsTests.cpp`;
- `PivotTests.cpp`;
- `CompatibilityTests.cpp`;
- `RegressionRegistry.cpp`.

Suite registration order and coverage remain unchanged. The largest regression TU is now roughly 2.3k lines and the domain files compile in parallel.

## Architectural invariants

- Public API lives under `include/XLPP/`.
- Model code is format-neutral.
- Package code is semantic-model-neutral.
- OOXML adapters own model/package translation.
- Preservation owns generic unknown-part/XML merge policy.
- Feature codecs own semantic mutation of their OOXML subtrees.
- IO owns load/save transaction mechanics.
- Architecture releases do not inflate Excel feature-parity scores.

## Verification baseline

Final P0Z-C source verification:

- clean Release core/build graph: **PASS**;
- standalone public-header compilation: **64/64 PASS**;
- native regression: **171/171 suites, 3,072/3,072 checks PASS**;
- full CTest registry: **4/4 PASS** (`XLPP_UnitTests`, architecture boundaries, package consumers, C ABI smoke);
- fresh `find_package()` and `add_subdirectory()` consumers: **PASS**;
- package-consumer gate after smoke-build optimization: approximately **30–35 seconds** on the verification host;
- package validator + generated sample: **PASS**;
- strict safety-warning core build: **PASS**;
- Clang ASan + UBSan full native suite: **PASS**, sanitizer stderr empty;
- Visual Studio source/test inventories regenerated for the physical tree with Platform Toolset **v145** (MSVC compilation itself was not available on the Linux verification host).
