# OPC Preservation Core

This document describes the current P0 preservation work for XL++.

## What is implemented

- Package-wide parsing of every OPC `.rels` part into `PreservedRelationship` records.
- Relative target resolution for root, workbook, worksheet, drawing, chart, pivot and other package parts.
- Package-graph validation for:
  - malformed relationship records (empty Id/Type/Target and invalid TargetMode);
  - duplicate relationship IDs per source part;
  - dangling internal relationships;
  - orphaned internal parts;
  - duplicate or malformed content-type declarations;
  - content-type overrides that point to missing parts;
  - package parts without a matching default or override.
- Owner/object-graph validation for:
  - workbook sheet nodes and their worksheet relationships;
  - worksheet drawing nodes and drawing relationships;
  - embedded and externally linked image relationships;
  - chart references inside DrawingML;
  - worksheet pivot-table references;
  - workbook pivot-cache references;
  - pivot-table `cacheId` consistency;
  - pivot-cache-record relationships;
  - worksheet table-part ownership;
  - legacy comment relationships and comment counts;
  - legacy VML drawing ownership;
  - workbook external-link ownership and external-link-path relationships;
  - object relationships that remain in `.rels` but are no longer referenced by owner XML.
- Reachable object inventory for worksheets, drawings, images, charts, shapes, text boxes, connectors, groups, other DrawingML objects, tables, comments, external links, pivot tables and pivot caches.
- Package comparison that reports object-count regressions, not only added/removed ZIP parts.
- Copy-on-write preservation for untouched worksheet drawings and pivot references.
- Preservation fallback for imported table/comment package objects that are not representable by the current read model; unrelated cell edits keep their owner XML/relationships/raw parts connected.
- Stale source comment/VML parts are suppressed when XL++ regenerates comments, preventing orphaned legacy drawings.
- Byte-preserving round-trip of untouched drawing, chart, image and pivot parts.
- Collision-safe relationship ID allocation. Original IDs are retained; generated IDs are moved to `rIdXLPP<n>` only when needed, and generated owner XML is updated consistently.
- Worksheet source tracking by sheet name, so advanced objects remain attached when another worksheet is removed and sheet indexes shift.
- Preservation of workbook-level pivot cache references.
- Correct OPC relationship and content-type declarations for custom document properties.
- A command-line package validator and package diff report.

## Validator

Build with the default CMake configuration or explicitly enable tools:

```bash
cmake -S . -B build -DXLPP_BUILD_TOOLS=ON
cmake --build build --target xlpp-package-validator
```

Validate one workbook:

```bash
xlpp-package-validator workbook.xlsx
```

The output includes a reachable object inventory. A chart or image is counted only when the complete owner-reference chain remains valid. Preservation-only DrawingML shapes/text boxes/connectors/groups are inventoried as well.

Compare a source workbook with a round-tripped workbook:

```bash
xlpp-package-validator before.xlsx --compare after.xlsx
```

The comparison lists added, removed and changed parts, validates both package graphs, and reports decreases in visible/reachable object counts. The command returns failure when the output package is invalid or an object-count regression is detected.

For CI, append `--json` to either validation mode. The tool emits one JSON object containing inventory, validation errors, package diffs and object-count regressions while preserving the same exit-code semantics.

## Test coverage

The XL++ preservation regression fixture contains:

- an embedded PNG image;
- a worksheet drawing;
- a bar chart;
- a pivot table;
- a pivot cache definition and cache records;
- two worksheets, where the leading worksheet is removed before save;
- an unrelated cell edit on the object-bearing worksheet.

The test verifies that all untouched visual and pivot parts remain byte-identical, all owner `r:id` references resolve, no parts are orphaned, no relationships dangle, and content types remain consistent.

Independent fixture coverage now also includes:

- an OpenPyXL-created workbook with one image and one chart using default-namespace DrawingML and one-cell anchors;
- a LibreOffice Calc round-tripped workbook with one image and one chart using prefixed DrawingML and two-cell anchors.

Both independent workbooks are loaded by XL++, modified in an unrelated cell, saved, and checked for byte-identical drawing/chart/media parts and unchanged visible object counts.

Negative fixtures verify that the validator detects:

- owner XML referencing a missing relationship ID;
- object relationships that remain in `.rels` after their worksheet drawing node is removed;
- table/external-link owner nodes removed while their relationships remain;
- comment-count regressions even when the comments part remains valid;
- malformed relationship records and targets that escape above the package root;
- image/chart parts that remain in the ZIP but are no longer visible;
- duplicate IDs, missing targets, orphaned parts, stale overrides, duplicate defaults and missing content types.

## Current limitations

- This work prioritizes preservation. It does not yet provide a complete DrawingML, chart or pivot object model.
- Imported embedded pictures support selective stable-ID mutation across their original drawing parts. Imported charts now expose a read model plus selective title/series-reference/move/resize editing; unsupported chart properties and general DrawingML objects remain preservation-first.
- LibreOffice-authored pivot preservation is covered; a Microsoft Excel-created pivot fixture is still required for P0 closure.
- Microsoft Excel Desktop recovery-log and visual validation must be run on Windows with Excel installed; it was not available in the Linux build environment used for this batch.
- General shapes, SmartArt, slicers, timelines, ActiveX and OLE remain preservation targets for later P0/P3 work.

## Next P0 batch

1. Add a Microsoft Excel-created pivot fixture and Excel Desktop recovery-log coverage.
2. Add Excel COM open/save/recovery-log checks on Windows.
3. Expand corruption fuzzing beyond deterministic relationship-target and namespace-prefix cases.
4. Add external-link/table/comment fixtures from Microsoft Excel and LibreOffice for host-level preservation evidence.
5. Expand chart selective editing from the current title/series-reference/geometry subset while keeping unsupported chart XML preservation-first.

## Drawing preservation foundation (P0D)

XL++ loads existing embedded images into the worksheet read model without marking the source drawing dirty. `Image::anchorInfo()` exposes one-cell, two-cell and absolute anchor metadata in EMU units together with package-origin identifiers (`stableId`, drawing part, media part and relationship ID).

When `Worksheet::addImage()` is used on a loaded worksheet that already owns a drawing, XL++ performs an additive package edit: it injects a new one-cell anchor into the preserved DrawingML part, allocates a collision-free image relationship/media part, and leaves existing chart XML and existing media bytes untouched.

## Safe imported-image mutation (P0E)

Imported embedded pictures can now be edited without regenerating the surrounding drawing. The selective APIs are `imageByStableId`, `moveImage`, `moveImageAbsolute`, `resizeImage`, `replaceImage` and `removeImage`. Save applies the pending edits directly to the target anchor/relationship/media while preserving sibling DrawingML verbatim.

Important preservation behavior:

- one-cell and absolute anchors patch only position/extent fields;
- two-cell moves preserve producer-native sub-cell offsets;
- two-cell resize updates both the terminal `<to>` marker and the picture transform extent so Excel/LibreOffice do not normalize back to the original size;
- replace reuses an unshared media part when safe, but allocates a private media part when the original resource is shared;
- remove deletes the relationship and media part only when no remaining image references that media;
- stale media `<Override>` entries in `[Content_Types].xml` are removed with deleted media parts;
- stable drawing-object IDs are used to target the exact picture while unrelated chart/image XML remains untouched.

Mutable access through `images()` still marks the complete drawing for regeneration. Selective mutation currently covers embedded pictures, not charts, shapes, text boxes or crop/rotation/flip metadata.

## Multi-drawing and unknown DrawingML preservation (P0F)

Selective image edits are now dispatched by `Image::sourceDrawingPart()` rather than assuming the first worksheet drawing relationship. In producer packages containing multiple explicit `<drawing r:id>` owner nodes, XL++ patches only the drawing that owns the target image and keeps untouched drawing parts byte-identical.

The package validator also scans namespace-qualified relationship-bearing attributes (`*:id`, `*:embed`, `*:link`) inside DrawingML. This allows preservation-only shapes, text boxes and extension objects to keep arbitrary relationship types (for example external hyperlinks) while still detecting both missing relationship records and `.rels` entries that have become unreferenced. The reachable inventory now records shapes, text boxes, connectors, groups and non-chart graphic objects, and `--json` provides machine-readable validation/diff output for CI.

## Extended owner graph and corruption hardening (P0G)

The validator now treats tables, legacy comments/VML notes and workbook external links as first-class owner-graph objects. `--compare` reports count regressions for these objects in addition to drawings/images/charts/pivots. Relationship parsing accepts namespace-prefixed `Relationship` elements, content-type inspection accepts namespace-prefixed `Default`/`Override` elements, and malformed relationship records are reported instead of being silently discarded.

Internal OPC target resolution now normalizes `.`/`..` segments and package-absolute targets while rejecting backslash paths, URI-scheme targets in Internal mode, and paths that escape above the package root.

Writer preservation was tightened for unsupported imported table/comment cases: source table/comment parts are retained in the preservation cache, unparsed tables keep their original `tableParts` owner and relationship, and regenerated comments suppress obsolete source comment/VML parts so they cannot remain orphaned. A regression fixture verifies `Workbook::load()` -> unrelated cell edit -> `save()` keeps one table, two comments and one external link reachable, with the external-link XML byte-identical.


## Chart preservation and selective-edit foundation (P0H)

Existing charts are now loaded with stable package identity and their owning drawing anchor. The reader accepts both `c:`-prefixed chart XML and default-namespace chart XML used by OpenPyXL, and records chart type/grouping, chart title, category/value axis titles, legend state, series formulas, one/two/absolute anchor geometry, source drawing/chart parts and relationship ID.

Selective mutation APIs are `chartByStableId`, `setChartTitle`, `setChartSeriesReferences`, `moveChart`, `moveChartAbsolute` and `resizeChart`. Save patches the preserved chart/drawing parts directly rather than serializing them through XL++'s smaller generated-chart model. Title edits retain the existing title formatting when a text run is available; series-reference edits replace only the selected formulas and remove only stale category/value caches. Unsupported chart XML such as data-label extensions, line/fill styling, axis IDs and producer-specific metadata remains preserved. Worksheet `sheetFormatPr` from the source package is also preserved verbatim where available; this keeps producer default row/column metrics stable so two-cell chart geometry does not drift when a host such as LibreOffice re-saves the workbook.

Mutable `chart()` / `charts()` access still intentionally marks the full drawing dirty and therefore remains the opt-in full-regeneration path. P0I adds safe append/remove plus deeper text/legend edits; advanced styling and complex chart structures remain later work.


## Deeper selective chart editing and lifecycle (P0I)

P0I extends the imported-chart path without widening the destructive serializer. Stable-ID edits now cover chart title, X/Y axis titles, legend visibility/position, individual series titles, series references, move and resize. Scatter and bubble charts treat the first and second `valAx` nodes as X and Y axes, so dual-value-axis titles are read and patched correctly. Multi-series regression uses an independently generated OpenPyXL scatter chart and confirms unsupported line/marker formatting plus a sibling image remain intact.

Imported charts can now be removed by stable ID. Save removes only the owning anchor and drawing relationship, suppresses the chart part plus exclusively-owned internal dependency closure, and cleans obsolete content-type overrides. Conversely, `Worksheet::addChart()` on a loaded worksheet uses an additive path: a collision-free `chartN.xml`, relationship ID and one-cell anchor are injected into the preserved drawing instead of forcing full DrawingML regeneration. Repeated saves rebuild from the original preservation baseline and do not duplicate the appended chart. Pure metadata edits do not rewrite the drawing `.rels`, preserving its original bytes.

The full-regeneration path remains available through mutable `chart()` / `charts()` access. P0J adds read-only combined-plot/native-axis structure plus selective secondary-axis title editing. P0K adds data-label/trendline/error-bar inspection and controlled mutation; combined-chart generation/restructuring, chart sheets and advanced formatting remain later work.


## Chart axis and series structure foundation (P0J)

P0J adds a preservation-oriented structural model for imported ChartML. `Chart::plots()` records each supported plot in document order with its chart type, grouping, series span and native `axId` references. `Chart::axes()` records category/value/date/series axis kind, `axId`, `crossAx`, position, title and whether the axis belongs exclusively to a later/secondary plot. `combined()`, `hasSecondaryAxes()`, `primaryXAxisId()`, `primaryYAxisId()` and `axisById()` provide read-only inspection without flattening the producer XML.

Primary X/Y title edits now resolve through the first plot's native axis IDs rather than positional `catAx`/`valAx` assumptions. `Worksheet::setChartAxisTitle(stableId, axisId, title)` targets a specific imported axis, including secondary value axes. A producer-independent OpenPyXL fixture combines a clustered bar plot (`axId` 10/100) with a line plot on a right-side secondary value axis (`axId` 200, `crossAx` 10). Regression verifies that primary and secondary titles can be edited while all native IDs, cross-axis links, the `crosses=max` setting and the combined bar+line structure remain intact.

LibreOffice headless host validation successfully opens and re-saves the edited workbook. LibreOffice normalizes the numeric axis IDs and may split a shared category axis into separate primary/secondary pairs, but XL++ reloads that normalized structure as two plots with secondary-axis classification still intact and the package graph remains valid.


## Chart labels, trendlines and error bars foundation (P0K)

P0K expands the imported-chart model without replacing producer ChartML. `Chart::Plot::dataLabels` records plot-level `<dLbls>`, while `ChartSeries::dataLabels()` records series-level labels used by hosts such as LibreOffice. `ChartSeries::trendlines()` reads multiple trendline nodes and records type, polynomial order, moving-average period, forward/backward projection and equation/R-squared flags. `ChartSeries::errorBars()` records X/Y direction, plus/minus/both behavior, value type, fixed/percentage/deviation value and end-cap state.

Selective APIs patch only the target subtree: `setChartPlotDataLabels`, `setChartSeriesDataLabels`, `setChartSeriesTrendline`, `addChartSeriesTrendline`, `removeChartSeriesTrendline`, `setChartSeriesErrorBars` and `removeChartSeriesErrorBars`. Existing trendline formatting/labels and unrelated series formatting remain preserved when a trendline is edited. Non-custom error bars can be added or edited; custom error bars are read and preserved, but writes are rejected until XL++ models the custom plus/minus reference formulas safely.

The independent OpenPyXL fixture contains a two-series Scatter chart with plot and series labels, linear/polynomial trendlines, fixed-value Y error bars and a sibling image. Regression confirms selective edits keep the drawing relationship part and sibling image byte-identical. LibreOffice host validation opens and re-saves the result successfully. LibreOffice normalizes plot-level labels into per-series `<dLbls>` nodes; the series-label model allows XL++ to read the normalized workbook without losing label metadata, while trendlines and X/Y error bars remain readable and the package graph stays valid.
