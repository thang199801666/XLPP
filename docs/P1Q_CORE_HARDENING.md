# P1Q-A Core Input Hardening & Streaming Guards

P1Q-A focuses on rejecting malformed or adversarial package input early while keeping normal streaming paths allocation-light. It intentionally avoids another value-type layout change.

## XML streaming

`XmlPullReader` now treats `>` inside quoted attributes correctly, rejects truncated matching elements, checks a source callback that returns more bytes than requested, guards size arithmetic, and limits its buffered look-ahead. The default limit is 64 MiB and can be changed through `StreamingReadOptions::maxXmlElementBytes`; set it to 0 only when the application explicitly wants unlimited buffering.

## Streaming ZIP/package guards

`StreamingReadOptions` exposes `maxEntries`, `maxEntryBytes`, `maxTotalBytes`, and `maxFileBytes`. These are forwarded to `ZipArchiveReader` central-directory parsing, so oversized packages fail before worksheet payload materialization. Zero keeps the corresponding limit unlimited.

## OPC relationship handling

Workbook worksheet targets are resolved relative to `xl/workbook.xml`. Dot segments are normalized, `..` may walk to package root but never above it, and internal targets containing a URI scheme or backslash are rejected. Duplicate relationship IDs, dangling worksheet IDs, duplicate worksheet names, and relationships resolving to missing package parts are rejected. Minimal packages with no workbook `.rels` continue to use the legacy `xl/worksheets/sheetN.xml` fallback.

## CFB strictness

CFB regular and mini streams must now provide at least the number of bytes declared by their directory entry. FAT/miniFAT early termination is an error. Sector address arithmetic is checked before multiplication/addition. Directory-tree traversal rejects out-of-range references, cycles, invalid linked entry types, empty linked names and duplicate case-folded paths.

## Verification

- XLPP_UnitTests: 198/198 suites, 3,632/3,632 checks PASS.
- XLPP_P1P_LazyFormulaTests: PASS.
- XLPP_P1Q_CoreHardeningTests: PASS.
- Focused ASan+UBSan/leak detection for all P1Q-touched parser paths: PASS.
- XLPP_HeaderCheck: PASS.
- xlpp_capi_smoke: PASS.
- staged `find_package(XLPP CONFIG)` consumer including `StreamingReadOptions`: PASS.

The remaining large `Workbook.cpp` and monolithic legacy test TU still dominate clean/instrumented build time; decomposition remains the next architectural hardening priority.
