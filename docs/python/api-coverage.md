# Python feature and API coverage

This page maps the Python binding to native XL++ domains. It describes XL++
binding coverage, not the complete feature set of Microsoft Excel.

## Parity status

The generated parity manifest currently maps all public Workbook and Worksheet
methods tracked by the repository:

| Native class | Tracked public methods | Python mapping |
| --- | ---: | ---: |
| `Workbook` | 64 | 64 |
| `Worksheet` | 168 | 168 |

Python uses `snake_case`, properties, overload adapters, and Python-native
`bytes`/lists/callables where direct C++ signatures are not idiomatic.

## Domain matrix

| Domain | Python binding | Semantic level | Guide |
| --- | --- | --- | --- |
| Workbook/worksheet lifecycle | Full public model | Author/read/write | [Guide](workbooks-and-worksheets.md) |
| File and memory I/O | Full | Author/read/write | [Guide](workbooks-and-worksheets.md#memory-io) |
| Cell values and types | Full | Author/read/write | [Guide](cells-ranges-rows.md) |
| Formulas and metadata | Full binding | Broad evaluator; not all Excel behavior | [Guide](formulas.md) |
| Rich text/comments/hyperlinks | Full public model | Author/read/write | [Guide](cells-ranges-rows.md) |
| Styles | Full public model | Author/read/write | [Guide](styles-and-layout.md) |
| Dimensions/layout/printing/views | Full public model | Author/read/write | [Guide](styles-and-layout.md) |
| Merges and panes | Full | Author/read/write | [Guide](styles-and-layout.md) |
| Named styles and names | Full public model | Author/read/write | [Guide](workbooks-and-worksheets.md) |
| Tables | Full public model | Author/read/write | [Guide](data-features.md) |
| AutoFilter | Full public model | Author/read/write/inspect | [Guide](data-features.md) |
| Conditional formatting | Full public model | Author/read/write | [Guide](data-features.md) |
| Data validation | Full public model | Author/read/write | [Guide](data-features.md) |
| Workbook/sheet protection | Full public model | Author/read/write | [Guide](styles-and-layout.md) |
| Structural editing | Full public helpers/reports | Reference-aware edit | [Guide](workbooks-and-worksheets.md) |
| Charts | Broad public model and imported editors | Broad authoring + targeted preservation edits | [Guide](charts-and-drawings.md) |
| Images/drawing anchors | Full public image model | Author/read/write/inspect | [Guide](charts-and-drawings.md) |
| Pivot tables | Full public model | Author/read/write/inspect | [Guide](pivot-tables.md) |
| Formula dependency graph | Full | Build/query/report | [Guide](formulas.md) |
| Reference translation | Full public helpers | Translate/invalidate | [Guide](formulas.md) |
| Streaming reader/writer | Full primary API | Sequential read/write | [Guide](streaming-and-performance.md) |
| NumPy/records | Full helper API | Bulk read/write | [Guide](streaming-and-performance.md) |
| VBA | Full public XL++ API | Preserve/binary/module authoring | [Guide](vba-encryption-enterprise.md) |
| Password-to-open encryption | Full options/inspection | Supported profiles read/write | [Guide](vba-encryption-enterprise.md) |
| External data | Main inspection fields | Inspect/preserve/targeted metadata | [Guide](vba-encryption-enterprise.md) |
| Data Model and OLAP | Main inspection fields | Inspect/preserve/targeted refresh | [Guide](vba-encryption-enterprise.md) |
| PivotCharts/slicers/timelines | Inventory + selected edits | Preservation-first | [Guide](vba-encryption-enterprise.md) |
| Power Query | Inventory + selected refresh metadata | Preservation-first | [Guide](vba-encryption-enterprise.md) |
| SmartArt/ActiveX/UserForms | Inventory | Preservation-first | [Guide](vba-encryption-enterprise.md) |

## Public Python model groups

### Core

`Workbook`, `Worksheet`, `Row`, `Cell`, `CellRange`, `CellReference`,
`DateTime`, `CellError`, `SaveOptions`, `LoadOptions`, and validation reports.

### Style and layout

`Color`, `Font`, `Fill`, `BorderSide`, `Border`, `Alignment`, `Style`,
`NamedStyle`, `RowDimension`, `ColumnDimension`, `SheetView`, `PageSetup`,
`PageMargins`, `PrintOptions`, `HeaderFooter`, worksheet/workbook protection.

### Data model

`Table`, `TableColumn`, `TableStyleInfo`, `AutoFilter` and filter models,
`ConditionalFormattingCollection` and rule models, `DataValidationCollection`,
`DefinedName`, document/custom properties.

### Charts and drawings

`Chart`, `ChartSeries`, plot/axis models, trendlines, error bars, data labels,
caches, fills, lines, theme colors, rich text, manual layouts, 3D/walls,
auxiliary chart objects, `Image`, and drawing-anchor metadata.

### Pivots

`PivotTable`, `PivotCache`, `PivotField`, `PivotFieldReference`, `PivotFilter`,
`PivotGrouping`, and `PivotLayout`.

### Calculation and editing

Calculation options/reports, dependency graph models, structural edit
options/reports, worksheet rename options/reports, chart-cache sync
options/reports, and reference translation results.

### Streaming and enterprise

Streaming reader/writer models, VBA module/project models, encryption info,
external-data inspection, Data Model inspection, enterprise feature inventory,
and targeted enterprise edit reports.

## Intentional non-bindings

Internal storage containers and loader implementation types are not public
Python API. Examples include native stable-vector implementation details and
private worksheet row-source plumbing. Python receives lists, stable child
handles, iterators, or domain objects instead.

## How parity is checked

`tools/generate_binding_manifest.py` parses native public headers and maps them
to Python/C# binding names. CI verifies `bindings/PARITY_MANIFEST.json` and
native architecture tests prevent accidental public-surface drift.

Parity does not replace runtime testing. The Python suite exercises module
construction, round trips, openpyxl interoperability, chart/image output,
NumPy paths, native report models, lifetime behavior, and public parity aliases.

## Known semantic gaps

- Complete Excel formula function catalog and exact calculation behavior
- Complete PivotChart/slicer/timeline semantic authoring
- Complete OLAP/Data Model and Power Query authoring/execution
- Complete SmartArt authoring
- Complete ActiveX and UserForms authoring
- Excel application automation, rendering, refresh, and recalculation behavior

See [Known gaps and development roadmap](../MISSING_FEATURES_AND_DEVELOPMENT_ROADMAP.md)
for the detailed backlog.
