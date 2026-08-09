> **P0Z-I reliability/enterprise baseline:** v1.12.0 adds dirty formula recalculation, advanced AutoFilter round-trip, External Data/Data Model inspection, C ABI v2 capability negotiation, package SDK workflows, and corpus-driven preservation gates. Native regression is 177/177 suites with 3,225 checks; enterprise corpus is 16/16 scenarios with zero unexpected removed parts; 1,000 seeded Clang/libFuzzer runs are clean on the release host. Excel Desktop remains an external Windows+Excel gate.

> **P0Z-H VBA baseline:** v1.11.0 adds source-authoring for Standard/Class/Document modules, stable worksheet VBA code names, VBA project properties/conditional constants, raw project export, and preserve-only safety for external/signed projects. Microsoft Excel Desktop macro execution/recovery-log validation is still external.

> **P0Z-G hardening baseline:** v1.10.0 retains the P0Z-F Pivot/Chart feature coverage and adds strict-warning-clean builds, durable save semantics, dependency-aware worksheet copy, deeper Pivot/Chart validation and sanitizer/libFuzzer release gates. Microsoft Excel Desktop recovery-log validation is still external.

> **P0Z-E binding/core baseline:** v1.8.0 preserves the 90.7/100 scope-defined native editing-core compatibility surface, stabilizes child-handle identity across collection growth, decomposes Formula function families, and makes native-derived binding-manifest drift a release gate. No additional Excel Desktop compatibility claim is implied by the version bump.

# XL++ Compatibility Matrix

_Last updated: 2026-08-09_

This matrix records automated compatibility evidence. `PASS` means the checked scenario is covered by an automated XL++ regression test or package-validator test. It is not a claim that the entire feature family is complete.


> **P0Y core reliability:** v1.3.0 adds strong-load rollback, atomic path saves, model/A1 preflight validation, hardened ZIP/ZIP64 materialized and streaming readers, bounded-memory ZIP64 file-backed writes, worksheet topology rewrite/invalidation, and an ASan+UBSan full-suite gate. These are reliability guarantees around the P0X compatibility surface rather than new Excel UI feature claims.

| Feature / producer | Generate | Read model | Unrelated-edit round-trip | Object graph validated | Untouched object bytes preserved | External host re-save | Status |
|---|---:|---:|---:|---:|---:|---:|---|
| Basic XLSX cells/styles — XL++ | PASS | PASS | PASS | PASS | N/A | Not run in this batch | Stable baseline |
| P0Z-I dirty formula recalculation — XL++ | PASS for supported calculation families | PASS | PASS for cached results | Dependency fan-out validated | N/A | Excel comparison corpus pending | Dirty roots select direct + transitive dependent formulas only |
| P0Z-I advanced AutoFilter — XL++ | PASS for Top10/dynamic/color/icon/date-group filters | PASS | PASS | OOXML reader/writer round-trip validated | N/A | Excel Desktop not available on this host | Advanced filter generation/reload PASS |
| P0Z-I External Data/Data Model inspection | Preservation-first; semantic generation not claimed | PASS for connection/query/external-link/model inventory subset | PASS for corpus scenarios | Part inventory/relationships inspected | PASS required for unknown enterprise payloads | Excel Desktop not available on this host | Inspection boundary + preservation-first policy |
| P0Z-I enterprise corpus — OpenPyXL/LibreOffice | N/A | PASS for exercised files | PASS: 16/16 operations | Package delta gate PASS | 0 unexpected removed parts | LibreOffice/OpenPyXL producers represented | Corpus foundation PASS |
| P0Z-H generated XLSM/VBA source project — XL++ | PASS for Standard/Class/Document modules, event procedures, project metadata and stable worksheet code names | PASS for XL++-generated project source/metadata; external projects remain opaque-preserved | PASS for generated project/model topology; signed/external destructive rewrites are blocked | PASS for content types, relationships, module topology and code names | PASS for untouched external/signed project bytes | Excel Desktop not available on this host | P0Z-H source authoring + preservation safety PASS |
| P0Z-F modern Excel ChartEx families — XL++ | PASS: Histogram, Pareto, Box & Whisker, Waterfall, Funnel, Treemap, Sunburst, Filled Map | PASS for generated ChartEx semantic type/cache subset | PASS for generated save/load | PASS | N/A for newly generated parts | Excel Desktop not available on this host | P0Z-F ChartEx generation/read-back PASS |
| P0Z-F Combo + Scatter/Bubble — XL++ | PASS for Combo/secondary axes, XY Scatter and Bubble size payload | PASS for generated semantic model | PASS | PASS | N/A for newly generated parts | Excel Desktop not available on this host | P0Z-F classic chart generation PASS |
| P0Z-F generated non-OLAP PivotTable — XL++ | PASS for row/column/page/multiple-data, grouping, filters, x14 Show-Values-As | PASS including cache records and field/options subset | PASS for generated semantic model | PASS | N/A for newly generated parts | Excel Desktop not available on this host | P0Z-F Pivot generation/read-back PASS |
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
| Pivot — LibreOffice Calc DataPilot | N/A | PASS for common semantic model (source/location/fields/layout/style/cache/filter/data options) | PASS | PASS | PASS when untouched; mutable access opts into semantic regeneration | PASS (LibreOffice headless preservation corpus) | Preservation + controlled semantic edit PASS |
| Preserved LibreOffice pivot + newly generated XL++ pivot | PASS for new pivot | Existing pivot semantic common subset | PASS | PASS | PASS for untouched original pivot/cache | PASS (LibreOffice headless preservation corpus) | Mixed preservation + semantic regeneration PASS |
| Pivot logical `cacheId` different from part suffix | PASS | N/A | PASS | PASS | PASS for original pivot/cache | Not separately host-tested | ID decoupling PASS |
| P0X formula calculation — XL++ core | PASS for supported formulas | PASS | Cached/spill results survive supported save/load workflows | N/A | N/A | Host comparison corpus pending | 26/26 scope-defined core capability families; dynamic arrays, structured refs, iterative cycles, dependency graph and external resolver included |
| Agile AES-256/SHA-512 encryption — XL++ → LibreOffice | PASS | PASS after password decrypt | PASS through load/save/password rotation | CFB/DataSpaces/HMAC validated | N/A | PASS: LibreOffice opened XL++ encrypted workbook and read expected cells | Modern OOXML encryption core PASS |
| Standard AES-128/192/256 + SHA-1 encryption — XL++ / LibreOffice | PASS for XL++ writer | PASS including independent LibreOffice AES-128 fixture | PASS through decrypt/re-save/password lifecycle | CFB/EncryptionInfo/DataSpaces validated | N/A | PASS: LibreOffice 25.2 opened XL++ Standard AES-128 output and read expected cells | Standard read/write interoperability PASS |
| Large Agile encrypted CFB/DIFAT — XL++ | PASS | PASS | PASS byte-for-byte for >9 MiB plaintext | FAT/DIFAT/HMAC validated | N/A | N/A | Large-container regression PASS |
| VBA source-generated project (legacy summary row) | PASS | PASS for XL++-generated source/metadata | PASS | PASS | External/signed project preservation covered | Microsoft Excel not available here | Superseded by P0Z-H VBA row above |
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
- Existing charts have a namespace-tolerant read model for common chart metadata and safe selective editing for title/axis/series text, legend, references and geometry. Native `axId`/`crossAx`, plot membership and primary/secondary-axis structure are exposed for imported charts; combined charts are preservation-aware and secondary-axis titles can be edited selectively by axis ID. P0K models plot-level and series-level data labels, multiple trendlines, and X/Y error bars. P0L adds per-point label overrides, custom plus/minus error-bar formula/range editing, and controlled series line/fill/marker plus trendline/error-bar line formatting. P0M adds imported `dPt`, rich title/label text, gradient/pattern/color-transform metadata and advanced line metadata; P0N adds manual layout plus axis/legend formatting; P0O adds axis min/max/log/orientation, numeric crosses-at, display units, gridline lifecycle, and chart/plot-area fill-line formatting. P0P adds chart data tables, drop/high-low lines, up/down bars, and plot/series leader-line formatting with selective add/remove lifecycle. Imported chart remove and additive append remain supported without regenerating sibling DrawingML. Combined-chart generation is now supported for the covered classic plot families; restructuring arbitrary imported combined charts, full 3D/surface editing fidelity, deeper theme/effect fidelity and complex chart dependency editing remain incomplete. P0R provides read/preserve coverage for six 3D/surface plot types plus Bar3D/Surface3D generation and selective `view3D`/wall editing. P0S adds Pie-of-Pie/Bar-of-Pie split controls, Doughnut/Pie slice-angle and hole-size metadata, and Radar style read/generate/selective editing. P0T adds chart style/theme resource discovery, scheme-color base resolution, and first-class title/category/value cache read/generate/selective editing while preserving untouched style resources. P0U adds worksheet-driven cache synchronization, sparse cache diagnostics, theme font/effect inspection and sequential final RGB/alpha transform resolution.
- Existing and generated worksheet-source non-OLAP pivots now have a substantial semantic read/edit model. Remaining gaps are OLAP/Data Model hierarchies, slicers/timelines, calculated fields/items, PivotCharts and deeper unsupported-extension editing.
- Microsoft Excel Desktop repair/recovery validation remains required before P0/M1 can be considered fully closed.
- P0G adds preservation fallback for source table/comment parts that are not fully represented in the object model; unrelated edits keep the original package graph connected rather than deleting the feature.
