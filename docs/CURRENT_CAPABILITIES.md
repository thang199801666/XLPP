# XL++ Current Capabilities

**Release:** v1.1.2  
**Status vocabulary:** `Supported`, `Supported with limitations`, `Preservation-first`, `Experimental`, `Unsupported`.

This document is the active capability summary. Historical gap lists and milestone notes are useful for provenance but must not be used as the current source of truth.

| Area | Current status | Notes |
|---|---|---|
| Workbook/worksheet/cell core | Supported | Read/write, structural edits, defined names, properties, views, protection and durable path saves |
| Styles/number formats | Supported with limitations | Broad common Excel styling; uncommon extension/theme fidelity remains corpus-dependent |
| Tables | Supported | Read/write/edit with structural reference updates |
| AutoFilter/sorting | Supported | Includes value/custom plus v1.12 Top10, dynamic, color/icon and date-group filters |
| Data validation | Supported | Common Excel validation types and round-trip |
| Conditional formatting | Supported with limitations | Common formula/cell/data-bar/color-scale/icon-set rules; uncommon Office extensions remain preservation-oriented |
| Formula storage | Supported | Formula text, cached values, shared/array/dynamic-array metadata |
| Formula calculation | Supported with limitations | Broad in-process engine, dependency graph, iterative/external resolution and dirty recalculation; includes working-day/ISO week functions, percentile/quartile, rank/mode and correlation/covariance/linear-regression families; not Excel-function complete |
| Charts | Supported with limitations | Broad classic/3D/ChartEx generation and selective imported editing; arbitrary ChartML extension fidelity is preservation-oriented |
| Non-OLAP PivotTables | Supported with limitations | Strong worksheet-source Pivot coverage; enterprise inventory identifies PivotCharts, slicers/timelines and OLAP caches, exposes incoming/outgoing OPC topology plus PivotTable/cache metadata, and supports selective PivotChart-source and OLAP refresh edits; calculated fields/items and full enterprise authoring remain gaps |
| Images | Supported | Common image anchoring/editing and preservation |
| General DrawingML/SmartArt | Preservation-first + inventory | SmartArt/diagram parts are explicitly inventoried and preserved; broad semantic shape/SmartArt authoring is not complete |
| VBA generated source projects | Supported with limitations | Standard/Class/Document modules, events, code names and project metadata; external/signed projects are safety/preservation surfaces |
| UserForms/FRX/ActiveX | Preservation-first + inventory | UserForm/FRX and ActiveX parts are explicitly inventoried; designer/control payloads are not semantically regenerated |
| Encryption | Supported | Agile AES-256/SHA-512 plus Office Standard AES/SHA-1 profiles covered by existing matrix |
| Streaming | Supported | Streaming read/write, limits, ZIP64 and large-sheet workflows |
| External workbook links | Preservation-first + inspection | v1.12 exposes metadata inventory; semantic rewrite is intentionally limited |
| Connections/query tables | Selective editing + preservation | Connection/query metadata is inspectable and refresh-on-load can be selectively edited without regenerating opaque sibling payloads |
| Power Query | Preservation-first + inventory | Mashup/query parts are explicitly inventoried; no M engine or semantic regeneration |
| Data Model/OLAP | Selective metadata editing + preservation | Model parts/relationships and OLAP Pivot caches are inventoried; OLAP cache refresh-on-load is editable, while proprietary model binaries remain preservation-only |
| C++ API | Supported | Primary API, C++20 |
| C ABI | Supported | ABI version 2 with additive capability negotiation |
| Python binding | Supported with limitations | Source surface and wheel workflow present; release-host binary verification requires pybind11/toolchain |
| C# binding | Supported | SafeHandle ownership, complete stable C ABI projection (869/869 exports), typed Chart/Pivot/Formula/Drawing/Streaming models, NuGet workflow, and Windows binary tests |
| Excel Desktop interoperability | Experimental/external gate | Automation/corpus infrastructure belongs on a Windows host with Microsoft Excel installed |

## Reliability baseline

- 177/177 native unit suites; 3,279/3,279 checks.
- Strict-warning native build passes on the release host.
- Full suite covered under ASan + UBSan in three bounded slices.
- Clang 17 libFuzzer load/validate/resave smoke: 1,000 runs from 16 third-party fixture seeds with no finding.
- Enterprise preservation corpus foundation: 16/16 scenarios, 0 unexpected removed parts.

For planned work, see the active
[roadmap](https://github.com/thang199801666/XLPP/blob/main/ROADMAP.md). Automated
verification is summarized in the repository's
[feature coverage](https://github.com/thang199801666/XLPP/blob/main/tests/FEATURE_COVERAGE.md).
