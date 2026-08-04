# Cross-library benchmarks

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
workbook benchmark. It uses 100K, 500K, and 1M rows with 10 columns and will
report cells/second plus peak memory when enabled. Its result table is kept in
the README and `docs/performance.md`; it is currently marked `n/a` rather than
mixing an unimplemented scenario with the normal-load baseline.

These are indicative CI measurements, not a substitute for a controlled local
performance study. GitHub Actions uploads the raw output as an artifact.
