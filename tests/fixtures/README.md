# External OOXML preservation fixtures
These files are immutable regression inputs. They are intentionally produced by applications other than XL++ so preservation tests do not validate only XL++'s own XML conventions.

## `openpyxl/image_chart.xlsx`

Generated with OpenPyXL 3.1.5. It contains:

- one worksheet named `Objects`;
- one embedded PNG image;
- one bar chart;
- one-cell DrawingML anchors;
- default-namespace spreadsheet drawing elements.

## `libreoffice/image_chart.xlsx`

Created by opening and saving the OpenPyXL fixture through LibreOffice Calc in headless mode. It contains the same visible image and chart but uses LibreOffice's serialization, including:

- prefixed `xdr:` DrawingML elements;
- two-cell anchors with `editAs="oneCell"`;
- LibreOffice workbook extensions and document properties.


## `libreoffice/pivot.xlsx`

Created directly with LibreOffice Calc's DataPilot API. It contains:

- one worksheet named `Data`;
- one pivot table named `SalesPivot`;
- one workbook pivot cache with cache records;
- a worksheet-to-pivot relationship without a `pivotTableParts` container, matching LibreOffice's serialization.

The pivot fixture is used both for unrelated-cell round-trip preservation and for mixed saves where XL++ adds a second pivot while the original LibreOffice pivot remains opaque and byte-identical.

The image/chart fixtures are additionally used by the Drawing Preservation Foundation tests to verify that XL++ reads producer-native image anchors and can append a new image into the existing drawing while leaving the original chart XML and media bytes untouched.

## Test contract

A preservation test loads each fixture with XL++, edits an unrelated cell, and saves a new workbook. The test then requires:

- zero package-graph or owner-reference errors;
- unchanged reachable drawing, image, chart, pivot-table and pivot-cache counts as applicable;
- byte-identical drawing/chart/media parts and byte-identical untouched pivot table/cache parts for unrelated edits;
- for additive image edits, the original chart/media bytes stay identical while the existing drawing part is extended rather than replaced.

Do not regenerate these files as part of normal test execution. Changes to fixture bytes should be reviewed as compatibility-corpus changes.

## `openpyxl/combined_secondary_axes.xlsx`

Created with OpenPyXL 3.1.5. The workbook contains one combined chart with a clustered bar plot on the primary category/value axis pair (`axId` 10/100) plus a line plot on a secondary value axis (`axId` 200, `crossAx` 10, right-side axis). It is used to validate plot ordering, native axis IDs/cross-axis links, primary/secondary-axis classification, and selective axis-title edits by `axId` without flattening the combined ChartML structure.


## `openpyxl/chart_labels_trendline_errorbars.xlsx`

Generated with OpenPyXL 3.1.5. It contains a two-series Scatter chart with plot-level and series-level data labels, a linear trendline, a polynomial trendline, fixed-value Y error bars, and a sibling PNG image. P0K uses this fixture to verify namespace-tolerant inspection plus selective label/trendline/error-bar mutation while the drawing relationship XML and sibling media stay byte-identical.

- `openpyxl/chart_auxiliary_objects.xlsx` — line chart with data table, drop/high-low lines, up/down bars, plot-level leader lines, and an image sibling for selective preservation regression.


## `openpyxl/stock_auxiliary_datatable_text.xlsx`

Created with OpenPyXL 3.1.5 and then augmented with valid ChartML formatting that OpenPyXL does not expose directly. It contains an open-high-low-close StockChart, high-low lines, up/down bars, a plot-area data table with `txPr` default text formatting, and a sibling PNG image. P0Q uses it to validate `Chart::Type::Stock`, selective auxiliary edits, data-table text-style preservation, and byte-identical drawing/media preservation.

## `openpyxl/projected_pie_doughnut_radar.xlsx`

Generated with OpenPyXL 3.1.5. It contains Pie-of-Pie, Bar-of-Pie, Doughnut and Radar charts plus a sibling PNG image. The projected charts exercise native `ofPieType`, split type/position, custom second-plot points, gap width and second-plot size; Doughnut exercises first-slice angle and hole size; Radar exercises native radar style and marker preservation. P0S uses the fixture for read-model, selective-edit, generation-parity and byte-preservation regression.
