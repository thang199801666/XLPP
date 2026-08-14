# Cross-library benchmarks

## Native regression target

With `-DXLPP_BUILD_BENCHMARKS=ON`, `xlpp_native_benchmark` is always built. `libxlsxwriter` is optional: when it is discoverable CMake also builds `xlpp_external_benchmark`; otherwise configuration continues with the native target only. This keeps performance-regression CI usable without weakening the external comparison workflow.


The benchmark uses the same deterministic 10,000 x 10 worksheet for every
implementation. Timings include workbook construction, styling, and the final
`.xlsx` write. Runs print machine-readable lines in this format:

```text
BENCHMARK,<language>,<library>,<operation>,<milliseconds>,<bytes>
```

Comparisons currently included:

- C++: XLPP and libxlsxwriter. libxlsxwriter is a write-only library, so this
  comparison reports write performance only.
- C#: XLPP and ClosedXML.
- Python: XLPP, openpyxl, and XlsxWriter. XlsxWriter is write-only; the Python
  runner reports read timings for libraries that support reading.

The large-file streaming comparison is defined separately from the normal
workbook benchmark. It uses 100K, 500K, and 1M rows with 10 columns and reports
read time plus file size. `n/a` is reserved for libraries without a streaming
reader, such as ClosedXML, XlsxWriter, and libxlsxwriter.

These are indicative CI measurements, not a substitute for a controlled local
performance study. GitHub Actions uploads the raw output as an artifact.

## Core hot-path regression target

`xlpp_core_hotpath_benchmark` isolates four internal-performance-sensitive
public API paths without needing an external comparison library:

- bulk row append into a single worksheet;
- repeated `dimensions()` / `maxRow()` / `maxColumn()` queries after the
  worksheet-extents cache is warm;
- repeated `trackedCellChangeCount()` union queries, which must not clone the
  tracked-cell tree or allocate O(number-of-cells) temporary memory;
- a Store-compression save, which makes worksheet/model serialization costs
  easier to see without deflate variability.

The first optional command-line argument is the row count (default: 30,000).
Output is machine-readable, for example:

```text
BENCHMARK,C++,XLPP,core_bulk_build,<milliseconds>,<cell-count>
METRIC,C++,XLPP,sizeof_cell_bytes,<bytes>
METRIC,C++,XLPP,sizeof_style_bytes,<bytes>
BENCHMARK,C++,XLPP,geometry_queries,<milliseconds>,20
BENCHMARK,C++,XLPP,tracked_change_count,<milliseconds>,5
BENCHMARK,C++,XLPP,save_store,<milliseconds>,<xlsx-bytes>
VERIFY,<rows>,<columns>,<dimensions>,tracked=<change-count>
```

This target is useful for catching accidental O(N) geometry-query regressions
or bulk-insert tracking overhead that functional tests alone cannot detect.
