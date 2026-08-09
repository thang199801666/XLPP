> **Superseded by P0Z-C/v1.6.0:** this document records the v1.5.0 intermediate decomposition. See `P0ZC_CORE_DECOMPOSITION.md` for the current tree.

# P0Z-B — Core Decomposition / v1.5.0

P0Z-B continues the P0Z architecture refactor without changing the public C++ API or the scope-defined 90.7/100 editing-core feature baseline. The milestone removes the remaining catch-all Workbook OOXML codec, gives read/write/preservation paths explicit internal owners, and makes the build graph mirror the source tree more closely.

## Workbook codec removal

`OOXML/Workbook/WorkbookCodec.cpp` no longer exists. Its responsibilities are now split across:

- `OOXML/Workbook/WorkbookReader.cpp` — package-to-workbook orchestration;
- `OOXML/Workbook/WorkbookWriter.cpp` — workbook-to-package orchestration;
- `OOXML/Workbook/WorkbookReferenceSupport.*` — print/name reference helpers;
- `OOXML/Worksheet/WorksheetReader.*` and `WorksheetWriter.*` — worksheet model codecs;
- `OOXML/Charts/ChartReader.*` — imported ChartML semantic read path;
- `OOXML/Charts/ChartWriterPatcher.*` — generated ChartML plus preservation-aware selective mutation;
- `OOXML/Drawings/DrawingReader.*`, `DrawingWriter.*`, `DrawingPackageSupport.*` and `PreservedDrawingEditor.*` — drawing ownership and imported-image mutation;
- `OOXML/Pivot/PivotCodec.*` — imported/generated Pivot semantic/package logic;
- `Preservation/XmlMergeSupport.*` and `PackageClosure.*` — OOXML merge and exclusive-part closure policies;
- `OOXML/Common/PackageRelationships.*` — relationship merge/allocation shared across package writers;
- `IO/BinaryFile.*` — binary file primitives at the IO layer.

The Workbook class remains the public facade; format-specific parsing and mutation are private implementation services.

## OOXML build modules

P0Z-B replaces the former single OOXML object target with subdomain object targets:

- `xlpp_ooxml_common_obj`
- `xlpp_ooxml_workbook_obj`
- `xlpp_ooxml_worksheet_obj`
- `xlpp_ooxml_charts_obj`
- `xlpp_ooxml_drawings_obj`
- `xlpp_ooxml_pivot_obj`
- `xlpp_ooxml_styles_obj`
- `xlpp_ooxml_tables_obj`
- `xlpp_ooxml_comments_obj`

Together with the non-OOXML domains, the core now builds through 20 internal object targets before aggregation into the same `xlpp_static` / `XLPP::xlpp` library.

This is a build-time/internal boundary only and does not change consumer linkage.

## Compile-path optimization

Two changes reduce avoidable development/CI cost:

1. `drawingTags()` caches XML source positions before sorting instead of repeatedly rescanning the full XML document from the comparator.
2. The large native `RegressionTests.cpp` is intentionally compiled without optimizer-heavy Release transformations and without the PCH, while production library code remains fully optimized. This avoids spending Release optimizer time on thousands of test assertions and removes the mismatched-PCH warning that would otherwise result from source-specific optimization settings.

`WorkbookWriter.cpp` also drops read-side Chart/Drawing/Worksheet includes and unused aliases after the codec split, reducing translation-unit coupling.

## Current large internal files

After P0Z-B decomposition:

- `OOXML/Workbook/WorkbookWriter.cpp`: ~795 lines;
- `OOXML/Workbook/WorkbookReader.cpp`: ~235 lines;
- `OOXML/Charts/ChartReader.cpp`: ~1.1k lines;
- `OOXML/Charts/ChartWriterPatcher.cpp`: ~2.9k lines;
- `tests/unit/Regression/RegressionTests.cpp`: ~7.6k lines.

The previous ~6.3k-line `WorkbookCodec.cpp` is eliminated. The next meaningful decomposition target is `ChartWriterPatcher.cpp`; it should be split only when serializer and selective-patcher helper ownership can remain clean. The regression TU should also continue splitting by Package/Chart/Pivot/Streaming/Compatibility domains.

## Architectural invariants

1. Public API stays in `include/XLPP/`; all codec/support headers under `src/XLPP/` remain private.
2. `Package/` owns bytes, ZIP/ZIP64, XML and OPC semantics and must not acquire Workbook-domain responsibilities.
3. `Model/` must not serialize OOXML.
4. `OOXML/` owns model/package translation.
5. `Preservation/` owns generic preservation/merge policy; feature codecs own semantic edits.
6. `IO/` owns file/stream transaction mechanics and binary file primitives.
7. Users still link one XL++ library.
8. Architecture-only releases do not inflate the Excel feature score.
