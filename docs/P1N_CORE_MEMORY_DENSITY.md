# P1N-A Core Memory Density & Change-Tracking Hardening

P1N-A targets the dominant remaining core cost identified by P1M-A: per-cell memory footprint and allocator/cache pressure on large worksheets. The work deliberately keeps existing public method signatures and semantics, while compacting internal/default state and removing an avoidable O(N) temporary allocation from change tracking.

## Compact default-state representation

A new internal `xlpp::internal::CompactString` stores canonical/default strings without allocating a `std::string` object per field. It deep-copies materialized values and collapses assignments of the canonical default back to the null representation.

The style model now uses this representation for fields that overwhelmingly remain at Excel defaults, including font name/color, fill pattern/colors, borders, alignment strings and number format. Formula metadata reference text uses the same strategy.

Local GCC/libstdc++ size probes for this source tree:

| Type | P1M-A | P1N-A |
|---|---:|---:|
| `Style` | 616 B | 184 B |
| `Font` | 80 B | 32 B |
| `Fill` | 96 B | 24 B |
| `Border` | 320 B | 80 B |
| `Alignment` | 80 B | 32 B |
| `FormulaMetadata` | 56 B | 32 B |
| `Hyperlink` | 104 B | 8 B |
| `Comment` | 64 B | 8 B |
| `Cell` | 1,072 B | 424 B |

`sizeof` is compiler/ABI dependent; these values are regression measurements, not portable ABI constants.

## Cell packing

`Cell` now packs row/column into one 64-bit coordinate key, stores the optional style index with a sentinel instead of `std::optional<std::size_t>`, and keeps formula text in compact default storage. Public row/column/style-index/formula APIs remain unchanged.

`Hyperlink` and `Comment` use lazy deep-copy PIMPL storage, so the common cell with neither feature no longer reserves their complete string payloads. Copy construction/assignment remains value-semantic; mutating a copy does not alias the source.

## Change-tracking query

`Worksheet::trackedCellChangeCount()` previously copied the tracked-key set and inserted every mutated cell into that temporary set. P1N-A instead merge-walks the already ordered cell map and tracked-key set. Complexity becomes O(N+K) time with O(1) auxiliary memory and no temporary tree allocation.

On a local GCC `-O0` 200,000-cell numeric worksheet, five repeated tracking queries changed from about 997 ms in P1M-A to about 91 ms in P1N-A (~10.9x faster). Peak RSS for that focused workload changed from about 236 MB to about 102 MB.

## End-to-end development benchmark

A local identical GCC `-O0` 20,000 x 10 / 200,000-cell benchmark measured the following medians. These are development A/B numbers intended to detect regressions, not universal library rankings.

| Metric | P1M-A | P1N-A | Change |
|---|---:|---:|---:|
| Bulk build | ~1,816.6 ms | ~1,244.7 ms | ~31.5% faster |
| Geometry query loop | ~27.2 ms | ~20.1 ms | ~26% faster |
| Store save | ~2,010.6 ms | ~1,423.0 ms | ~29.2% faster |
| Peak RSS | ~270.8 MB | ~146.0 MB | ~46.1% lower |

P1N-A adds `sizeof_cell_bytes`, `sizeof_style_bytes`, and `tracked_change_count` output to `xlpp_core_hotpath_benchmark` so future CI/performance runs can catch representation or tracking regressions.

## Correctness verification

- Full unit regression: 196/196 suites, 3,600/3,600 checks PASS.
- New compact-model tests cover default semantics, equality/hash, deep-copy isolation, reset-to-default behavior, formula metadata, hyperlink/comment ownership and tracking-count duplicate handling.
- Focused ASan+UBSan stress with 50,000 compact cells and copies: PASS.
- Standalone public-header compile check: PASS.
- C API build + `xlpp_capi_smoke`: PASS.

## Compatibility note

P1N-A preserves the public C++ **source-level method signatures and behavior**, but intentionally changes the in-memory layout/size of public C++ value types such as `Cell`, `Style`, `Hyperlink` and `Comment`. Therefore this is a **C++ binary ABI change**. Rebuild the XL++ library and all native consumers/bindings together; do not mix P1M-A-compiled objects with a P1N-A library.

The public release version remains `1.1.2` because P1N-A is a development milestone rather than a binary-compatible release cut.

## Next core target

The largest remaining engineering debt is compile-time scalability: `src/XLPP/Workbook/Workbook.cpp` is roughly 724 KB / 12.3K lines, and optimized GCC builds spend disproportionate time in that translation unit. The next batch should split package read/write, preservation, chart/Pivot/VBA and serializer logic into cohesive translation units before pursuing deeper object-container changes. Persistent parser/mutation fuzzing and a controlled million-cell latency/RSS benchmark should follow in parallel.
