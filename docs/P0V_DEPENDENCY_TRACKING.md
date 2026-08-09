# P0V Dependency Tracking — P0V-A Implementation Note

## Goal

P0V-A makes chart-cache synchronization dependency-aware without introducing `Cell -> Worksheet -> Workbook` ownership callbacks into the existing object model.

## Design

`Workbook` stores an exact source snapshot for each supported chart cache reference. The key identifies owner worksheet, imported stable chart ID (or generated chart index), series index, cache kind and normalized reference. The snapshot records source worksheet/range, workbook date epoch, every source coordinate, cell value, formula text and number format.

This produces conservative deterministic invalidation:

- first synchronization: register + synchronize every supported reference;
- unchanged source range: `changedReferencesOnly=true` skips rebuilding and applying the cache;
- changed source value/formula/format: only references containing that source are rebuilt;
- unrelated worksheet edit: no tracked chart dependency changes;
- reference identity change: new key, therefore a fresh synchronization;
- tracker reset/load/clear: next changed-only synchronization behaves as a full registration pass.

## Why snapshot tracking first

The current `Cell` public API is value-like and has no owner callback. Adding owner pointers now would require auditing worksheet copy/move/erase behavior and every mutator, increasing preservation risk. P0V-A therefore establishes correct observable dependency behavior with no Cell ABI ownership change. P0V-B can introduce a reusable mutation journal/registry after copy/move semantics are explicitly designed.

## Public API

```cpp
xlpp::ChartCacheSyncOptions options;
options.changedReferencesOnly = true;
auto report = workbook.synchronizeChartCaches(options);

workbook.trackedChartCacheDependencyCount();
workbook.resetChartCacheDependencyTracking();
```

Report additions:

- `referencesChecked`
- `referencesUnchanged`
- `dependenciesRegistered`
- `dependenciesChanged`

Opt-in save preparation:

```cpp
xlpp::SaveOptions save;
save.synchronizeChartCaches = true;
save.synchronizeChangedChartCachesOnly = true;
workbook.save("output.xlsx", save);
```

Save-time synchronization occurs on a private workbook copy and does not mutate caller chart caches.

## Verified regression cases

- initial dependency registration and cache generation;
- repeated changed-only synchronization skips all unchanged caches;
- unrelated cell edit does not invalidate chart sources;
- one value-range cell edit invalidates only the value dependency;
- number-format/formula-sensitive snapshot behavior;
- explicit tracker reset;
- opt-in save-time synchronization materializes caches in output while caller state remains cache-free;
- full project baseline: 155/155 suites, 2,731 checks passing.

## P0V-B / P1 follow-up

The next layer should generalize this into a workbook dependency registry capable of representing:

`Cell/Range -> Formula -> DefinedName/Table -> Chart/Pivot/Validation/CF`

and pair it with structural reference rewriting for insert/delete row/column operations. **P0X completes that follow-up** with a public formula/name/table dependency graph plus workbook-level transactional structural editing. The P0V chart snapshot mechanism remains a specialized incremental cache optimization.
