# XL++ Compatibility Matrix

_Last updated: 2026-08-07_

This matrix records automated compatibility evidence. `PASS` means the checked scenario is covered by an automated XL++ regression test or package-validator test. It is not a claim that the entire feature family is complete.

| Feature / producer | Generate | Read model | Unrelated-edit round-trip | Object graph validated | Untouched object bytes preserved | External host re-save | Status |
|---|---:|---:|---:|---:|---:|---:|---|
| Basic XLSX cells/styles — XL++ | PASS | PASS | PASS | PASS | N/A | Not run in this batch | Stable baseline |
| Image + chart — OpenPyXL 3.1.5 | N/A | PASS for image anchors and common chart metadata | PASS | PASS | PASS for untouched siblings | PASS via LibreOffice headless after selective chart edits | Image mutation + deeper selective chart editing PASS |
| Image + chart — LibreOffice Calc | N/A | PASS for image anchors and common chart metadata | PASS | PASS | PASS for untouched siblings | PASS via LibreOffice headless after selective chart edits | Image mutation + deeper selective chart editing PASS |
| Combined bar + line / secondary axis — OpenPyXL 3.1.5 | N/A | PASS for plot/axis structure | PASS | PASS | PASS for selective axis-title edits | PASS via LibreOffice headless | Native axId/crossAx + secondary-axis model PASS |
| Scatter + 2 series + sibling image — OpenPyXL 3.1.5 | N/A | PASS including dual value axes and series titles | PASS | PASS | PASS for unsupported line formatting and sibling image | Host validation in P0I | Multi-series scatter selective editing PASS |
| Scatter labels + trendlines + error bars + sibling image — OpenPyXL 3.1.5 | N/A | PASS for plot/series labels, trendlines and error bars | PASS | PASS | PASS for sibling image/relationship bytes and unsupported series formatting | PASS via LibreOffice headless; LO normalizes plot labels to series labels | P0K selective chart-feature foundation PASS |
| Scatter per-point labels + custom error ranges + rich series formatting — OpenPyXL 3.1.5 | N/A | PASS for `dLbl/idx`, custom plus/minus ranges, line/fill/marker and trendline/error-bar line formatting | PASS | PASS | PASS for sibling image and drawing relationships | PASS via LibreOffice headless; Calc preserves object graph/ranges but normalizes inherited point-label flags and some marker/fill styling | P0L selective formatting foundation PASS |
| Scatter `dPt` + rich text + advanced fill/line — OpenPyXL-derived fixture | N/A | PASS for rich title/point labels, `dPt`, gradient/pattern fill, color transforms, custom dash/cap/compound/join | PASS | PASS | PASS for sibling image and drawing relationships | PASS via LibreOffice headless for object graph/rich text/`dPt`; Calc normalizes gradient/pattern/color transforms and may materialize default `dPt/dLbl` nodes | P0M selective advanced-formatting foundation PASS |
| Chart style/theme + series caches + sibling image — LibreOffice/OpenPyXL-derived fixture | PASS for generated style IDs and series caches | PASS for workbook theme palette, scheme colors/transforms, chart-style/color-style resources, `strCache` and `numCache` | PASS | PASS | PASS for untouched style/color-style relationship parts and sibling drawing/image resources | Partial: LibreOffice re-computes caches from worksheet formulas and removes style/color-style relationship resources/scheme styling on re-save | P0T style/theme/cache foundation PASS; Calc normalization documented |
| Chart cache synchronization + theme transform engine — LibreOffice/OpenPyXL-derived fixture | PASS for generated cache synchronization | PASS for title/category/value caches, sparse indexes, theme font/effect metadata and final transformed RGB/alpha | PASS | PASS | PASS for untouched chart-style/color-style resources and sibling image/drawing relationships | PASS via LibreOffice headless; Calc may materialize blank string categories and normalize numeric cache format codes | P0U cache synchronization/theme transform foundation PASS |
| Projected pie + Doughnut + Radar + sibling image — OpenPyXL 3.1.5 | PASS for Pie-of-Pie, Bar-of-Pie, Doughnut and Radar generation | PASS for `ofPieType`, split controls, first-slice angle, hole size, radar style and series-line formatting | PASS | PASS | PASS for untouched sibling chart/image/drawing relationship bytes | Partial: LibreOffice keeps all four charts but normalizes projected-pie split settings and doughnut hole size on re-save | P0S projected-pie/doughnut/radar foundation PASS |
| 3D/surface charts + sibling image — OpenPyXL 3.1.5 | PASS for Bar3D/Surface3D generation | PASS for Bar3D, Line3D, Area3D, Pie3D, Surface, Surface3D + `view3D`/walls | PASS | PASS | PASS for five untouched sibling chart parts + image/drawing relationships | Partial: LibreOffice preserves Bar3D/Line3D/Area3D/Pie3D but converts Surface/Surface3D to Bar3D on re-save | P0R preservation/generation foundation PASS |
| Table/comments/external link — OpenPyXL + injected external-link OOXML | N/A | Partial (legacy comments; table may remain opaque) | PASS | PASS | PASS for opaque table/external-link parts | Partial: table/comments survive LibreOffice re-save; LibreOffice removes the external link | P0G preservation PASS; external-link host retention not claimed |
| Pivot — LibreOffice Calc DataPilot | N/A | Opaque preservation | PASS | PASS | PASS | PASS (LibreOffice headless) | Preservation PASS |
| Preserved LibreOffice pivot + newly generated XL++ pivot | PASS for new pivot | Existing pivot opaque | PASS | PASS | PASS for original pivot/cache | PASS (LibreOffice headless) | Mixed preservation PASS |
| Pivot logical `cacheId` different from part suffix | PASS | N/A | PASS | PASS | PASS for original pivot/cache | Not separately host-tested | ID decoupling PASS |
| VBA source-generated project | Experimental | Partial | Covered by existing automated tests | Package checks | External project preservation covered | Microsoft Excel not available here | Experimental |
| Microsoft Excel-created image/chart/pivot fixtures | N/A | Not yet in corpus | Not yet | Not yet | Not yet | Not yet | Missing fixture |
| Excel Desktop recovery-log / COM validation | N/A | N/A | N/A | N/A | N/A | Not available in current Linux environment | Pending |

## Current P0 validator guarantees

The package validator currently checks:

- malformed relationship records (empty Id/Type/Target, invalid TargetMode);
- duplicate relationship IDs;
- dangling internal relationships;
- orphaned internal parts;
- `[Content_Types].xml` consistency;
- workbook-to-worksheet ownership;
- worksheet-to-drawing ownership;
- drawing-to-image/chart ownership;
- generic DrawingML `r:id` / `r:embed` / `r:link` ownership for unsupported relationships;
- multiple worksheet drawing owner nodes and relationships;
- worksheet table-part ownership;
- legacy comments and VML-note ownership;
- workbook external-reference/external-link ownership;
- drawing inventory for shapes, text boxes, connectors, groups and non-chart graphic objects;
- workbook-to-pivot-cache ownership;
- worksheet-to-pivot-table ownership, including LibreOffice's relationship-only worksheet serialization;
- pivot-table `cacheId` cross-linking to the workbook cache declaration;
- pivot-table-to-cache-definition relationship consistency;
- pivot-cache-definition-to-cache-records ownership;
- before/after reachable object-count regressions for drawings/images/charts/shapes/tables/comments/external links/pivots.

## Known preservation gaps

- Existing worksheet images have a read model for one-cell, two-cell and absolute anchors. Imported images can be moved, resized, removed or replaced selectively by stable ID, and new images can be appended without regenerating untouched sibling image/chart XML.
- Selective mutation covers imported embedded pictures and a controlled imported-chart subset. Chart stable-ID edits support chart/axis/series titles, legend state/position, category/value or xVal/yVal references, move/resize, selective remove, and additive new charts on a preserved drawing while retaining unsupported ChartML.
- Shapes, text boxes, groups, crop/rotation/flip and other DrawingML extensions remain preservation-only, but their relationship-bearing XML is now owner-validated and preserved during selective image edits.
- Existing charts have a namespace-tolerant read model for common chart metadata and safe selective editing for title/axis/series text, legend, references and geometry. Native `axId`/`crossAx`, plot membership and primary/secondary-axis structure are exposed for imported charts; combined charts are preservation-aware and secondary-axis titles can be edited selectively by axis ID. P0K models plot-level and series-level data labels, multiple trendlines, and X/Y error bars. P0L adds per-point label overrides, custom plus/minus error-bar formula/range editing, and controlled series line/fill/marker plus trendline/error-bar line formatting. P0M adds imported `dPt`, rich title/label text, gradient/pattern/color-transform metadata and advanced line metadata; P0N adds manual layout plus axis/legend formatting; P0O adds axis min/max/log/orientation, numeric crosses-at, display units, gridline lifecycle, and chart/plot-area fill-line formatting. P0P adds chart data tables, drop/high-low lines, up/down bars, and plot/series leader-line formatting with selective add/remove lifecycle. Imported chart remove and additive append remain supported without regenerating sibling DrawingML. Combined-chart generation/restructuring, full 3D/surface editing fidelity, deeper theme/effect fidelity and complex chart dependency editing remain incomplete. P0R provides read/preserve coverage for six 3D/surface plot types plus Bar3D/Surface3D generation and selective `view3D`/wall editing. P0S adds Pie-of-Pie/Bar-of-Pie split controls, Doughnut/Pie slice-angle and hole-size metadata, and Radar style read/generate/selective editing. P0T adds chart style/theme resource discovery, scheme-color base resolution, and first-class title/category/value cache read/generate/selective editing while preserving untouched style resources. P0U adds worksheet-driven cache synchronization, sparse cache diagnostics, theme font/effect inspection and sequential final RGB/alpha transform resolution.
- Existing pivots are preserved and graph-validated but are not yet exposed as a complete read/edit public model.
- Microsoft Excel Desktop repair/recovery validation remains required before P0/M1 can be considered fully closed.
- P0G adds preservation fallback for source table/comment parts that are not fully represented in the object model; unrelated edits keep the original package graph connected rather than deleting the feature.
