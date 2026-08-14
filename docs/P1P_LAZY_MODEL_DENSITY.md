# P1P-A Lazy Model Density & Formula/Tracking Hardening

P1P-A continues the P1O-A scaling work by targeting the largest remaining always-resident pieces inside a default `Cell`: inline `Style` and `FormulaMetadata` storage. The guiding rule is unchanged: reduce default memory and allocator/cache pressure without weakening value semantics, stable worksheet handles or OOXML correctness.

## Lazy value storage

P1P-A adds the internal `xlpp::internal::CompactValue<T>`. Its empty representation is one owning pointer. A value is allocated only when mutable state is actually needed; copying a populated `CompactValue<T>` performs a deep copy.

`Cell` now uses lazy value storage for:

- `Style`;
- `FormulaMetadata`.

This builds on the compact strings and optional payloads introduced in P1N/P1O. Const/default reads return semantic default objects without allocation. Mutable access materializes an owned value before returning a reference.

On the local GCC/libstdc++ ABI probe:

| Metric | P1O-A | P1P-A | Change |
|---|---:|---:|---:|
| `sizeof(Cell)` | 352 B | **152 B** | **~56.8% lower** |
| `sizeof(Style)` | 184 B | 184 B | unchanged |

Relative to the 1,072-byte pre-P1N cell layout, the P1P-A default cell is about **85.8% smaller**. `sizeof` remains compiler/ABI dependent.

## Formula metadata correctness and density

`Cell::setFormula()` now explicitly returns a cell to normal-formula semantics and clears stale shared/array/dynamic metadata. Structural reference rewriting uses the separate internal-purpose `setFormulaTextPreservingMetadata()` path, so a reference rewrite does not accidentally demote a shared/array formula.

The OOXML reader was also changed so ordinary formulas do **not** materialize `FormulaMetadata` merely to store default values. Metadata is allocated only when the `<f>` element actually carries non-default state (`shared`, `array`, `dataTable`, detected dynamic-array state, `ref`, `si`, `aca`, or `ca`). A dedicated package round-trip regression loads 5,000 ordinary formulas plus shared/array formulas and verifies that ordinary formulas remain metadata-allocation-free while the special formula metadata survives.

## Mutation tracking hardening

Mutable access to state that can affect serialization now forms a mutation boundary. P1P-A covers mutable style/formula metadata, comments, hyperlinks, named-style state, rich text and raw style-index changes. This reduces the risk that differential/change-aware save paths miss an edit performed through a mutable subobject reference after `clearDirty()`.

A merely materialized but untouched default `Style` remains semantically default and is not treated as a non-default style by the serializer.

## Serializer/default-style fast path

Save-time scans use `Cell::hasNonDefaultStyle()` before consulting the style catalog. Default cells therefore avoid materializing or deeply traversing style state. Formula/reference structural scans use const metadata access first and materialize only if a rewrite actually needs to update metadata.

## Optimized million-cell A/B

Same-machine GCC `-O1`, 100,000 x 10 / 1,000,000 numeric cells:

| Metric | P1O-A | P1P-A | Change |
|---|---:|---:|---:|
| `sizeof(Cell)` | 352 B | **152 B** | **~56.8% lower** |
| Peak RSS | 633,303,040 B | **441,380,864 B** | **~30.3% lower** |
| Bulk build | 1,717.172 ms | **1,065.867 ms** | **~37.9% faster** |
| Five tracked-change scans | 456.608 ms | **244.661 ms** | **~46.4% faster** |
| Store save | 3,848.857 ms | **3,778.813 ms** | **~1.8% faster / effectively flat** |
| Store output bytes | 67,566,941 | 67,566,941 | identical |

The strongest result is memory/cache locality rather than a claimed universal save-throughput gain. Timings are local regression measurements and should not be generalized across compilers, machines or workbook shapes.

## Default vs fully styled density probe

P1P-A adds `xlpp_cell_density_benchmark [rows] [default|styled]` to make the cost of lazy style allocation visible. A current 20,000 x 10 / 200,000-cell GCC `-O1` run measured:

| Workload | Peak RSS | Build | Store save |
|---|---:|---:|---:|
| Default numeric cells | ~89.0 MB | ~335.7 ms | ~613.1 ms |
| Every cell styled | ~170.5 MB | ~660.1 ms | ~801.9 ms |

The fully styled path intentionally pays for one owned style per cell. P1P-A does not introduce shared mutable/COW style semantics solely to improve that pathological benchmark, because doing so would complicate reference/value behavior. Future style interning should be considered only with explicit mutation-isolation tests.

## Verification

- Main regression: **198/198 suites, 3,632/3,632 checks PASS** after relinking against the final core.
- Dedicated `XLPP_P1P_LazyFormulaTests`: PASS; verifies package round-trip allocation behavior for ordinary/shared/array formula metadata.
- C API rebuild + `xlpp_capi_smoke`: PASS.
- Standalone public-header target, including `CompactValue.h`: PASS.
- Staged install + external `find_package(XLPP CONFIG)` consumer: PASS.
- Focused Clang ASan+UBSan 50,000-cell stress for lazy Style/FormulaMetadata/Comment/Hyperlink deep-copy and mutation: PASS with leak detection/halt-on-error.

A full instrumented project build was attempted with both GCC and Clang ASan+UBSan. The remaining monolithic `Workbook.cpp` sanitizer compile exceeded the execution window in this environment, so P1P-A **does not claim a full sanitizer-suite pass**. The focused sanitizer result covers the new lazy value mechanism; completing the remaining translation-unit decomposition is carried forward as a build-scalability and full-sanitizer-enablement target.

## Compatibility note

Ordinary source usage remains compatible, but P1P-A is a **C++ ABI change** relative to P1O-A because `Cell` layout changes from 352 to 152 bytes. In addition, mutable `style()` / `formulaMetadata()` access may allocate and therefore is no longer treated as allocation-free/noexcept behavior. Rebuild XL++ and all native C++ consumers/bindings together. Opaque C API clients remain source compatible after rebuilding the wrapper/core.

The project version remains `1.1.2`; P1P-A is a development milestone, not a binary-compatible release cut.

## Recommended P1Q target

1. Split package reader/writer and worksheet serializer/preservation helpers out of `Workbook.cpp` so full ASan/UBSan and optimized builds become incremental and bounded in compiler RSS/time.
2. Begin decomposing the legacy unit-test monolith into subsystem translation units; new P1P formula-density regression is already isolated in its own TU.
3. Add persistent fuzz targets/corpora for ZIP, CFB, XML, SST, formula/reference parsing and mutation sequences.
4. Replace remaining heuristic reference rewriting with a reusable formula/reference token stream/AST.
5. Investigate worksheet sparse-node overhead and optional style interning only behind stable-handle and mutation-isolation tests.
6. Add controlled MSVC v145 + C#/Python/C runtime CI and million-cell latency/RSS gates.
