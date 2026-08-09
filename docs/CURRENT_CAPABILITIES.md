# XL++ Current Capabilities

**Release:** v1.12.0 / P0Z-I  
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
| Formula calculation | Supported with limitations | Broad in-process engine, dependency graph, iterative/external resolution and v1.12 dirty recalculation; not Excel-function complete |
| Charts | Supported with limitations | Broad classic/3D/ChartEx generation and selective imported editing; arbitrary ChartML extension fidelity is preservation-oriented |
| Non-OLAP PivotTables | Supported with limitations | Strong worksheet-source Pivot coverage; calculated fields/items, PivotCharts, slicers/timelines and OLAP/Data Model authoring remain gaps |
| Images | Supported | Common image anchoring/editing and preservation |
| General DrawingML/SmartArt | Preservation-first | Unknown drawings are preserved where possible; broad semantic shape/SmartArt authoring is not complete |
| VBA generated source projects | Supported with limitations | Standard/Class/Document modules, events, code names and project metadata; external/signed projects are safety/preservation surfaces |
| UserForms/FRX/ActiveX | Preservation-first | Not fully semantically authored |
| Encryption | Supported | Agile AES-256/SHA-512 plus Office Standard AES/SHA-1 profiles covered by existing matrix |
| Streaming | Supported | Streaming read/write, limits, ZIP64 and large-sheet workflows |
| External workbook links | Preservation-first + inspection | v1.12 exposes metadata inventory; semantic rewrite is intentionally limited |
| Connections/query tables | Preservation-first + inspection | v1.12 exposes connection/query metadata without regenerating opaque enterprise payloads |
| Power Query | Preservation-first + detection | No M engine or semantic authoring |
| Data Model/OLAP | Preservation-first + inspection | v1.12 detects model parts/relationships and OLAP Pivot caches; no proprietary model rewrite |
| C++ API | Supported | Primary API, C++20 |
| C ABI | Supported | ABI version 2 with additive capability negotiation |
| Python binding | Supported with limitations | Source surface and wheel workflow present; release-host binary verification requires pybind11/toolchain |
| C# binding | Supported with limitations | SafeHandle ownership and NuGet workflow present; release-host binary verification requires .NET SDK |
| Excel Desktop interoperability | Experimental/external gate | Automation/corpus infrastructure belongs on a Windows host with Microsoft Excel installed |

## Reliability baseline

- 177/177 native unit suites; 3,225/3,225 checks.
- Strict-warning native build passes on the release host.
- Full suite covered under ASan + UBSan in three bounded slices.
- Clang 17 libFuzzer load/validate/resave smoke: 1,000 runs from 16 third-party fixture seeds with no finding.
- Enterprise preservation corpus foundation: 16/16 scenarios, 0 unexpected removed parts.

For release-specific details, see `docs/P0ZI_PHASE28_TO_37_REFINEMENT.md` and `PACKAGE_STATUS.md`.
