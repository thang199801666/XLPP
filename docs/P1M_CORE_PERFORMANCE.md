# P1M-A Core Hot-Path & Streaming Performance Notes

Date: 2026-08-13

P1M-A continues the core-hardening work from P1L-A. The goal of this batch is
not feature breadth: it removes avoidable CPU/memory work from common worksheet
and streaming-reader paths while preserving public API behavior and increasing
negative/regression coverage.

## Implemented changes

### Worksheet geometry and bulk append

- `Worksheet::extents()`, `maxRow()`, `maxColumn()` and `dimensions()` now share
  a lazily populated extents cache.
- Single-cell insertion updates a valid cache incrementally.
- Structural row/column edits invalidate the cache before the next query.
- `Worksheet::append()` inserts cells directly instead of routing every value
  through the random-access `cell()` tracking path. Mutation revision/change
  tracking remains intact.

This keeps repeated geometry queries O(1) after the first scan and avoids a
tracking-tree insertion for every cell in bulk append workloads.

### Worksheet serialization

- Sequential sheet XML serialization now walks the row-major cell map directly.
- The default sequential path no longer builds an O(N-cells) pointer array plus
  row-span table before writing XML.
- Parallel-row mode materializes row work structures only when explicitly used.
- A regression test covers the final row/value in a 3,000-row parallel-row save.
  This test was added after development caught an unsafe stream-buffer shortcut;
  the shortcut was removed rather than trading correctness for a micro-optimization.

### Shared strings and save metadata

- Save-time shared-string indexing owns each unique string once: the ordered SST
  list references stable keys in the index instead of storing a second copy.
- Comment and external-hyperlink presence are collected during the existing
  style/SST cell scan rather than rescanning each sheet later.

### Streaming reader hardening

- Shared-string loading scans `<si>` / `<t>` elements through `string_view`
  slices and allocates only the final decoded values.
- Entity-free text follows a zero-copy input scan; XML unescaping occurs only
  when required.
- Worksheet row, style, numeric value and shared-string indexes now require
  complete strict parses. Malformed numeric prefixes and out-of-range SST
  indexes are rejected instead of falling through permissive standard-library
  conversions.

## Verification

Full regression build (Clang, Release semantics with `-O0 -DNDEBUG -g0`):

- **195 / 195 test suites PASS**
- **3,583 / 3,583 checks PASS**

Focused ASan + UBSan smoke coverage passed for a combined path containing:

- extents-cache population;
- 5,000-row bulk append;
- structural row insertion/cache invalidation;
- parallel-row save;
- streaming readback.

## Local before/after measurements

These values are local development regression measurements. They are useful for
comparing P1L-A and P1M-A on the same machine/build profile, not as universal
library claims.

### DOM/core hot path — 30,000 x 10 cells, Store save

Median P1L-A -> P1M-A using the same Clang verification profile:

| Operation | P1L-A | P1M-A | Change |
| --- | ---: | ---: | ---: |
| Bulk build | ~1,060 ms | ~889 ms | ~16% faster |
| 20 geometry query loops | ~908 ms | ~49 ms | ~18.5x faster |
| Store save | ~1,746 ms | ~1,487 ms | ~15% faster |
| Peak RSS | ~421.7 MB | ~402.5 MB | ~4.5% lower |

Output size remained identical at 18,944,481 bytes.

### Streaming read — 20,000 x 10 cells, 120,000 unique strings

Median P1L-A -> P1M-A:

| Metric | P1L-A | P1M-A | Change |
| --- | ---: | ---: | ---: |
| Read time | ~603 ms | ~546 ms | ~9.6% faster |
| Peak RSS | ~31.9 MB | ~25.4 MB | ~20% lower |

Both readers produced exactly 20,000 rows / 200,000 cells and the same numeric
checksum.

## Indicative cross-library write comparison

A 20,000 x 10 workload (six unique-string columns, four numeric columns) was
also run locally to establish a rough order of magnitude. XL++ used GCC `-O1`;
Python libraries were XlsxWriter 3.2.9 and openpyxl 3.1.5.

| Library | Median operation time | Approx. peak RSS |
| --- | ---: | ---: |
| XL++ | ~0.60 s | ~269 MB |
| XlsxWriter | ~0.97 s | ~152 MB |
| openpyxl | ~1.42 s | ~194 MB |

On this workload XL++ is about 1.6x faster than XlsxWriter and 2.4x faster than
openpyxl, but its DOM-oriented model uses materially more memory. The comparison
is intentionally labelled indicative: languages, object models, compression
pipelines and measurement boundaries differ. Memory footprint is therefore a
higher-priority optimization target for the next core batch than chasing small
additional write-latency wins.

## Persistent benchmark target

P1M-A adds `xlpp_core_hotpath_benchmark`, which reports machine-readable timing
for bulk append, repeated geometry queries and Store save. It is independent of
external libraries and can be used as a lightweight performance-regression CI
signal.

## Next core priorities

1. Reduce per-cell memory footprint and allocator pressure without public API/ABI
   churn, especially string/formula/style storage in large sparse/dense sheets.
2. Split `Workbook.cpp` into writer/reader/package/preservation translation units
   so optimized builds do not bottleneck on one very large compilation unit.
3. Add persistent fuzz targets/corpora for streaming XML/SST, ZIP/CFB and formula
   reference parsing.
4. Continue the reusable formula/reference token/AST work required for fully
   non-heuristic structural transformations.
5. Extend benchmark coverage to million-cell streaming/write workloads on a
   controlled machine and track latency + peak RSS together.
