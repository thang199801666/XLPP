# P1O-A Core Decomposition & Density Scaling

P1O-A continues the P1N-A correctness/performance hardening work with two goals: reduce the compile-time cost of the monolithic workbook implementation and lower the default per-cell footprint without changing the public source API.

## Workbook translation-unit decomposition

`src/XLPP/Workbook/Workbook.cpp` previously carried unrelated VBA/UserForm, Pivot mutation, chart-cache/reference synchronization and chart-style mutation logic in the same translation unit. P1O-A moves those cohesive groups into:

- `WorkbookVba.cpp`
- `WorkbookPivot.cpp`
- `WorkbookChartCache.cpp`
- `WorkbookChartStyle.cpp`

CMake discovers the new sources automatically and the Visual Studio v145 project explicitly includes them.

Local source-size comparison:

| Metric | P1N-A | P1O-A | Change |
|---|---:|---:|---:|
| `Workbook.cpp` bytes | 724,828 | 525,433 | ~27.5% lower |
| `Workbook.cpp` lines | 12,282 | 8,773 | ~28.6% lower |

A same-profile GCC `-O0`, single-job compile of only `Workbook.cpp` changed from about **28.46 s / 969,944 KB compiler peak RSS** to about **22.51 s / 807,508 KB**, or roughly **20.9% less elapsed time** and **16.7% lower compiler peak RSS**. This is a development-machine regression measurement, not a portable compiler guarantee.

## Compact rare cell payloads

P1O-A adds the internal `xlpp::internal::CompactOptional<T>`. The empty representation is one pointer; storage is allocated only when the optional value is engaged. Copy construction/assignment deep-copies the optional payload, so public value semantics remain intact.

The cell model now uses compact optional storage for rare/default-empty payloads:

- rich text;
- named style name;
- hyperlink;
- comment.

Public accessors that return `const std::optional<T>&` remain source compatible via a stable empty optional reference when no payload is materialized.

On the local GCC/libstdc++ size probe, `sizeof(Cell)` changes from **424 B in P1N-A to 352 B in P1O-A** (~17.0% smaller). Relative to the pre-P1N 1,072-byte layout, the current cell is ~67.2% smaller. `sizeof` is compiler/ABI dependent.

## Append insertion fast path

`Worksheet::append()` inserts monotonically increasing row-major keys. P1O-A supplies `cells_.end()` as the `std::map::try_emplace` insertion hint, avoiding a redundant logarithmic tree search for each cell while retaining the existing ordered map and its stable node/reference behavior.

## Empty-sheet extents invariant

An empty worksheet already has the canonical `A1:A1` extent, so P1O-A initializes the extents cache as valid instead of invalid. `cell()` and monotonic `append()` can therefore maintain min/max extents from the first inserted cell onward. This removes the previous first-query O(N) rescan after bulk append.

On the optimized one-million-cell A/B, the 20-query geometry block changed from a median of about **79.3 ms** in P1N-A to about **0.029 ms** in P1O-A (roughly **2,700x** for this benchmark shape). Structural row/column edits still invalidate the cache and force a correctness-first rebuild on the next query.

## Benchmark observability

`xlpp_core_hotpath_benchmark` now emits a cross-platform `peak_rss_bytes` metric in addition to object-size, build, geometry, tracking and save metrics. On Windows it uses `GetProcessMemoryInfo`; on Unix-like systems it uses `getrusage`.

Same-profile GCC `-O0` A/B, 20,000 x 10 / 200,000 cells, three fresh runs per build:

| Metric | P1N-A median | P1O-A median | Change |
|---|---:|---:|---:|
| Bulk build | ~1,267.0 ms | ~1,109.2 ms | ~12.5% faster |
| Store save | ~1,386.6 ms | ~1,374.0 ms | ~0.9% faster / effectively flat |
| External peak RSS | ~145,952 KB | ~129,792 KB | ~11.1% lower |

A one-million-cell (100,000 x 10) same-profile run measured:

| Metric | P1N-A | P1O-A | Change |
|---|---:|---:|---:|
| Bulk build | ~6,340 ms | ~5,631 ms | ~11.2% faster |
| Store save | ~7,477 ms | ~7,537 ms | ~0.8% slower / within local run noise |
| External peak RSS | ~697,604 KB | ~619,332 KB | ~11.2% lower |

The million-cell save result is intentionally reported as flat/noisy rather than presented as a speedup. P1O-A's demonstrated benefit is lower cell memory, faster bulk construction and better compile scalability without a material save-throughput regression.

## Optimized `-O1` scaling check

A separate same-machine GCC `-O1` build was used to ensure the development `-O0` results were not an optimizer artifact. For one million cells, recent three-run medians were approximately:

| Metric | P1N-A | P1O-A | Change |
|---|---:|---:|---:|
| Bulk build | ~2,234 ms | ~1,706 ms | ~23.6% faster |
| Geometry query block | ~79.3 ms | ~0.029 ms | ~2,700x faster |
| Store save | ~3,753 ms | ~3,580 ms | ~4.6% faster in this run set |
| Peak RSS (representative) | ~696,800 KB | ~618,500 KB | ~11.2% lower |
| Five tracked-change scans | ~393 ms | ~413 ms | ~5% slower; retained as a P1P watch item |

The save and tracking deltas show more run-to-run variance than cell build/RSS, so they are not treated as universal guarantees. The material wins of this batch are object density, append/build throughput, geometry-cache behavior and reduced compiler peak memory.

For build scalability, a clean `-O1 -j2` build of the static core + hot-path benchmark reduced observed peak compiler/build RSS from about **1,203,644 KB to 821,964 KB (~31.7%)**, but total clean elapsed time increased from **78.32 s to 88.85 s (~13.4%)** because four new translation units each pay header/front-end overhead. Conversely, compiling only the main `Workbook.cpp` dropped from about **28.46 s / 969,944 KB** to **22.51 s / 807,508 KB** in the same development setup. P1P should preserve the lower per-TU memory while recovering clean-build wall time with better TU boundaries/PCH or other build-graph work.

## Correctness verification

- Full non-sanitized unit regression: **197/197 suites, 3,613/3,613 checks PASS**.
- C API rebuild + `xlpp_capi_smoke`: PASS.
- Standalone public-header target, including `CompactOptional.h`: PASS.
- Staged install + installed `find_package(XLPP)` consumer: PASS.
- New compact-optional regression checks deep-copy isolation, reset-to-empty behavior and the cell-size guard.
- Full ASan+UBSan regression with leak detection/halt-on-error: **197/197 suites, 3,613/3,613 checks PASS**.

## Compatibility note

Public C++ method signatures are intentionally preserved, but `Cell` layout changes again in P1O-A. This is therefore another **C++ binary ABI change** relative to P1N-A. Rebuild XL++ and all native C++ consumers/bindings together. Opaque-handle C API clients remain source-compatible after rebuilding the core/C wrapper.

The project version remains `1.1.2`; P1O-A is a development milestone rather than a binary-compatible release cut.

## Recommended P1P target

1. Continue decomposition by moving package read/write and worksheet serializer/preservation logic out of `Workbook.cpp`.
2. Split the large unit-test translation unit by subsystem so optimized test builds scale with the library.
3. Add persistent fuzz targets/corpora for ZIP, CFB, XML, SST, formula/reference and mutation sequences.
4. Introduce a reusable formula/reference token stream/AST for unions, intersections, structured references, names, dynamic arrays, 3-D and external references.
5. Investigate sparse-cell container/node overhead only behind explicit stable-handle tests; do not trade semantic stability for a benchmark-only gain.
6. Add controlled MSVC v145, C#/Python runtime and million-cell performance CI gates.
