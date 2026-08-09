# P0X — 90% General-purpose Excel Editing Engine Gate

> Historical acceptance gate: **P0Y/v1.3.0** preserves this 90.7/100 feature-scope result and supersedes its reliability baseline with stronger I/O, ZIP/ZIP64, topology, validation and sanitizer guarantees. See `P0Y_CORE_HARDENING.md`.

P0X broadens the acceptance target from the P0W Formula/Encryption subsystems to
the complete **general-purpose XLSX editing core**. The score below is deliberately
weighted by practical editing risk. It is **not** a claim that XL++ implements 90%
of every Microsoft Excel UI feature or every Excel function.

## P0X acceptance result

| Capability domain | Weight | P0X score | Evidence / remaining gap |
|---|---:|---:|---|
| OOXML package, ZIP, OPC, streaming and large-file robustness | 12 | 11.5 | ZIP64, mmap/streaming paths, relationship graph, validators, package consumers; exotic producer corruption remains open-ended |
| Workbook/worksheet/cell/style core | 13 | 12.0 | Strong read/write model, dates/errors/rich text/styles/names/protection; some rare workbook-view metadata remains preservation-first |
| Common Excel features | 12 | 10.0 | Tables, filters/sorts, validation, conditional formatting, comments, print/page setup, hyperlinks, images; advanced DrawingML ecosystem remains incomplete |
| Formula storage, calculation and dependency analysis | 15 | 14.7 | 26/26 core capability families, dynamic arrays, structured refs, external resolver, iterative cycles, public dependency graph; specialized Excel catalog still incomplete |
| Reference-safe structural editing | 15 | 13.8 | Transactional insert/delete rows/columns rewrites formulas/names/tables/filters/CF/DV/print/charts/pivots/anchors/hyperlinks with rollback; move/copy/rename topology can go deeper |
| Charts / drawings / media | 10 | 8.5 | Broad chart generation/read/selective edit and preservation-aware DrawingML; complex shapes/effects and arbitrary chart restructuring remain incomplete |
| Pivot tables | 8 | 6.9 | Imported semantic model/edit, regeneration, layout/style/cache/data-field/report-filter options; grouping/calculated fields/slicers/timelines remain |
| OOXML password encryption and security | 5 | 4.8 | Agile AES-256/SHA-512 read/write, Standard AES-128/192/256 SHA-1 read/write, integrity/hardening/interoperability; certificate key encryptors remain |
| Preservation and independent interoperability | 6 | 5.2 | Relationship/object graph regression and LibreOffice/OpenPyXL fixtures; Microsoft Excel Desktop repair-log/COM gate remains |
| Bindings, packaging, samples and CI | 4 | 3.3 | C ABI smoke is automated, Python/C# parity expanded, install consumers pass, source sample and cross-platform CI added; Python/C# builds still require their SDK dependencies |

**Weighted P0X result: 90.7 / 100.**

The target is therefore closed for the stated **general-purpose XLSX editing core**.
Advanced ecosystem features (ActiveX, slicers/timelines, PowerPivot/Data Model,
Cube functions, arbitrary OfficeArt effects, full VBA designer semantics, etc.) are
tracked outside this 90% core gate and must not be presented as complete.

## 1. Workbook dependency graph

`Workbook::dependencyGraph()` produces a non-mutating graph over formula cells and
classifies dependencies as cells/ranges, defined names, tables, external references
or volatile references.

```cpp
const auto graph = workbook.dependencyGraph();
auto inputs = graph.precedentsOf("Summary", "D10");
auto users = graph.dependentsOf("Data", "B7");
const bool linked = graph.dependsOn("Summary", "D10", "Data", "B7");
```

Defined names and structured tables are also expanded to their concrete worksheet
ranges where possible, making the graph useful for invalidation and editor tooling.

## 2. Transactional structural editing

The workbook-level structural API is the preferred path for edits that can affect
objects outside the target sheet:

```cpp
xlpp::StructuralEditOptions options;
options.transactional = true;
options.failOnInvalidReference = true;

auto report = workbook.insertRows("Data", 10, 2, options);
```

The transaction updates:

- moved/removed cells and row/column dimensions;
- normal/shared/array/dynamic-array formula references and metadata;
- workbook/local defined names;
- merged ranges and freeze panes;
- AutoFilter/sort ranges;
- conditional formatting ranges/formulas;
- data validation ranges/formulas;
- table references and table-column models;
- print area / print titles;
- internal hyperlinks;
- chart series/error-bar references and chart/image anchors;
- pivot source/location references and cache invalidation.

If `failOnInvalidReference` is enabled, the full edit is preflighted on a private
copy. A transaction that would produce an invalid `#REF!` dependency is rejected
without replacing the live workbook, so previously acquired worksheet/cell handles
remain stable.

## 3. Formula engine after P0X

The original P0W 21/26 matrix is now 26/26 for the defined core families:

- dynamic-array spill calculation (`SEQUENCE`, `SORT`, `UNIQUE`, `FILTER`,
  `TRANSPOSE`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `HSTACK`, `VSTACK`,
  `TOROW`, `TOCOL`);
- structured table references including data/header/total selectors, current-row
  selectors and rectangular multi-column selectors;
- `INDIRECT`, `OFFSET`, `ROWS`, `COLUMNS`, `ROW`, `COLUMN`, `ADDRESS` and `LET`;
- fixed-point iterative calculation for circular numerical models using
  `maxIterations` and `maxChange`;
- external workbook formulas through an explicit `externalReferenceResolver`
  callback. XL++ never opens an external workbook implicitly.

External formulas support quoted and unquoted forms such as:

```text
'[Budget.xlsx]Rates'!A1:A10
[Budget.xlsx]Rates!B2
```

The resolver receives the workbook token, worksheet and canonical A1 address. A
missing value becomes `#REF!` and is counted in the calculation report.

The 26/26 score measures **engine capability families**, not the entire Excel
function catalog. Specialized database/cube/engineering/statistical/financial
functions can still be added without changing this core-family result.

## 4. Pivot semantic editing after P0X

Imported pivots remain byte-preserved until mutable pivot access explicitly opts in
to semantic regeneration. The model now round-trips common report semantics:

- compact / outline / tabular layout;
- row/column grand totals and formatting flags;
- PivotTable style name/header/stripe flags;
- cache refresh/save/enable flags;
- field sorting, subtotal placement, blank rows and new-item filter behavior;
- report-filter selected item and multi-selection metadata;
- data-field caption, number format, aggregation, show-data-as, base field/item.

The writer snapshots its pivot generation plan before worksheet dirty-state cleanup,
so imported-pivot regeneration cannot disappear midway through package creation.

## 5. Encryption after P0X

Password-to-open OOXML support now includes both major profiles used by modern Office
containers:

- Agile AES-256-CBC / SHA-512 — read + write + HMAC integrity;
- Standard AES-128/192/256 / SHA-1 — read + write.

The Standard writer uses the required 50,000-round password derivation and emits
Office CFB/DataSpaces metadata. Independent host evidence includes:

- LibreOffice opening XL++ Agile output;
- XL++ opening a LibreOffice Standard AES-128 fixture;
- LibreOffice opening XL++ Standard AES-128 output and reading expected cells.

The modern password-encryption matrix is now **19/21 = 90.5%**. The remaining two
items are alternate Agile cipher/hash profiles and certificate/private-key key
encryptors.

## 6. Binding and packaging closure

P0X expands Python/C/C# surfaces for calculation, structural editing and encryption.
The C ABI has an automated smoke target covering formula calculation, iterative
calculation, structural edits with stable handles, Agile encryption and Standard
encryption inspection. CMake now registers that test whenever `XLPP_BUILD_TESTS` is
enabled.

A buildable `samples/xlpp_p0x_sample` target removes the previous default-sample
warning, and `.github/workflows/ci.yml` defines native Linux/Windows/macOS configure,
build and CTest gates.

## 7. Final verification gate

The final v1.2.0 source gate completed with **169/169 native test suites and 2,992/2,992 checks passing**. CTest registered and passed all three native/package tests: the unit suite, the install/consumer test (`find_package()` + `add_subdirectory()`), and the C ABI smoke test. Public headers, the package validator, C ABI library and the P0X source sample all build successfully.

The verification host also reopened XL++ Standard AES-128/SHA-1 output with LibreOffice 25.2.3.2 and read the expected cells after password authentication. Python and C# binding source parity was expanded, but those two extension projects were not compiled on this host because `pybind11` and the .NET SDK are not installed.

## 8. Remaining work beyond the 90% gate

Highest-value follow-on work is no longer basic XLSX I/O. It is compatibility depth:

1. Microsoft Excel Desktop fixture corpus, repair-log detection and COM validation.
2. Pivot grouping, calculated fields/items, richer filters, slicers/timelines.
3. Advanced DrawingML shapes/effects/groups and arbitrary chart restructuring.
4. Worksheet rename/move/copy dependency translation across every object family.
5. More specialized Excel functions and richer LAMBDA semantics.
6. Certificate-based Agile encryption and alternate crypto profiles.
7. Continuous Python/C# build/test jobs with their full SDK dependencies.

