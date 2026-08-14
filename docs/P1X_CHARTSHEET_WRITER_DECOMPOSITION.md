# P1X-A Chartsheet Writer Decomposition & Relationship-Collision Hardening

P1X-A is an architectural/core-hardening milestone for `.xltx/.xltm` Chartsheets. It intentionally avoids adding a new public object family. The goal is to move package-writing ownership out of the monolithic `Workbook.cpp` and to harden relationship-ID conflict repair around preservation-backed Chartsheets.

## Changes

- Added `WorkbookChartsheetPackage.cpp/.h` and moved the complete Chartsheet package-write loop out of `Workbook.cpp`.
- The extracted writer owns generated/imported Chartsheet XML relationship assembly, drawing/chart closure replacement, `printerSettings` emission, preservation sibling merge and repeated-save dirty-state behavior.
- `Workbook.cpp` falls from ~9,547 to ~9,440 lines in this milestone.
- Added a dedicated internal access shim; no new public API or `Cell`/`Style` layout change is required.

## Relationship collision defect fixed

P1X regression constructed a valid preservation-backed Chartsheet where:

- imported chart drawing used `rIdDrawing`;
- a preserved `legacyDrawingHF` VML relationship intentionally used `rId1`;
- regenerated chart drawing initially requested the normal generated `rId1`.

The pre-P1X merge logic repaired the relationship collision by globally replacing every `r:id="rId1"` in the owner XML. That also changed the preserved `legacyDrawingHF` owner reference, leaving it inconsistent with the preserved VML relationship.

P1X makes collision repair owner/type-aware:

- generated `drawing` IDs are patched only inside the `<drawing>` owner tag;
- generated `printerSettings` IDs are patched only inside `<pageSetup>`;
- preservation-backed `legacyDrawingHF`/unknown nodes are never rewritten by a collision belonging to another relationship type.

The dedicated regression validates the package graph after the collision repair and again after a repeated save.

## Verification

- Main regression: 198/198 suites, 3,632 checks PASS.
- P1P through P1X dedicated tests: PASS.
- C API smoke: PASS.
- Standalone public-header build: PASS.
- `find_package(XLPP CONFIG)` consumer: PASS.
- `add_subdirectory()` consumer: PASS (verification build used `-O0 -DNDEBUG -g0` to avoid the known optimized clean-build compiler-time bottleneck).
- Focused Clang ASan+UBSan+leak detection on `XLPP_P1X_ChartsheetWriterTests`: PASS.

## Next target

P1Y should continue the same direction rather than regrowing `Workbook.cpp`: extract generic OPC relationship merge/closure helpers and then model explicit header/footer image-media replacement/removal above the preserved VML layer.
