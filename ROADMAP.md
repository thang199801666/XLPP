# XL++ Development Roadmap

> Single source of truth for the project's development status and direction.
> Historical per-milestone notes were consolidated here on 2026-08-16.

## Project status

XL++ is a modern C++20 Excel library focused on speed, correctness and
preservation-friendly read/write. It reads and writes `.xlsx`, `.xlsm`, `.xls`
(BIFF8), `.xlsb` (BIFF12) and CSV, plus password-to-open encrypted Office
packages, without requiring Excel to be installed.

### Currently implemented

- Workbook / worksheet model with cells, ranges, styles, named styles,
  comments, hyperlinks, merged cells, tables, AutoFilters, conditional
  formatting, data validation, page setup, protection, sparklines and
  defined names.
- Full OOXML package round-trip: unknown parts/relationships are preserved
  byte-for-byte; strict (ISO 29500) and transitional namespaces are supported.
- Charts (native chart + imported chart preservation, chart-cache sync,
  style/theme matrix materialization), PivotTables (imported + generated,
  shared caches, grouping, filters, selective mutation) and chartsheets.
- Slicers, shapes/text boxes, images and drawing parts.
- Formula dependency analysis and an in-process calculation engine
  (arithmetic, ranges, 40+ functions, cross-sheet references, circular
  detection) via `Workbook::calculate()`.
- VBA project read/write: standard/class/document modules, designer
  (UserForm) modules, MS-OFORMS binary authoring through
  `Workbook::addUserForm()` and semantic UserForm inspection/editing.
- Streaming writer/reader APIs, parallel sheet serialization, transactional
  output and deterministic package ordering.
- Encryption: Agile (AES-128/192/256 + SHA-1/256/384/512) and Standard
  CryptoAPI AES password-to-open read/write.
- Structural edits with cross-sheet reference repair and transactional
  rollback, plus an in-memory model validator.
- Bindings: C API, C# and Python; packaging for CMake `find_package` and
  `add_subdirectory` consumers.
- Hardening: memory-dense `Cell`/`Style` layout, sanitizer (ASan/UBSan)
  profiles, resource limits on untrusted input and a structured fuzz harness
  for `Workbook::load()`.

### Verification baseline

- **222/222 unit-test suites PASS** (plus dedicated P1P–P1X executables).
- Full suite passes under MSVC AddressSanitizer with no memory errors.
- Interop validated against openpyxl/XlsxWriter/LibreOffice and Office-created
  files; encrypted output opens in LibreOffice.

## Development roadmap (future work)

Priorities, in rough order:

1. **VBA depth** — complete MS-OFORMS control-family coverage (SpinButton,
   ScrollBar, MultiPage, TabStrip, Image, Frame), VBA project signing and
   protection, and richer `dir`/`PROJECT` stream round-tripping.
2. **Formula engine expansion** — array/dynamic-array evaluation, more
   functions, defined-name and structured-reference resolution inside
   `calculate()`, and iterative/circular recalculation policies.
3. **Pivot depth** — selective cache-record/shared-item mutation, richer
   grouping/filter edge cases, imported PivotChart corpora and OLAP
   cache-field handling.
4. **Chart depth** — extended chart types (radar/3-D surface), error
   bars/trendline authoring and chart-style/theme matrix authoring.
5. **Performance** — controlled million-cell latency + peak-RSS benchmarks,
   per-cell allocator traffic reduction and faster clean optimized builds.
6. **Engineering** — continue translation-unit decomposition, persistent
   fuzz corpora/CI gates for ZIP/CFB/XML/formula parsers, MSVC v145 +
   C#/Python runtime CI, and bindings parity for the newest native APIs.

## Notes

- Public API is source-compatible across milestones; the P1N-era compact
  `Cell`/`Style` layout is an intentional C++ ABI change (consumers must
  rebuild).
- Preservation-first behavior: unsupported or unknown package content is kept
  byte-for-byte rather than dropped, except where an explicit opt-in
  regeneration API is used.
