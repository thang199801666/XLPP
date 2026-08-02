# XL++ Core-to-100% Roadmap

## Current focus: production-grade read/write core

1. **Milestone 15 — Streaming foundation** — complete
   - Append-only streaming writer
   - Row callback reader
   - Chunked zlib ZIP output
   - File-backed package entries

2. **Milestone 16 — Direct streaming ZIP reader** — complete
   - Central-directory index without loading the full archive
   - Incremental inflate of worksheet entries
   - Pull iterator and callback APIs
   - Shared-string lazy access and cache policy

3. **Milestone 17 — Fast XML scanner** — complete
   - Non-allocating token scanner based on `string_view`
   - Direct numeric conversion with `from_chars`
   - Reduced temporary strings and attribute maps
   - Benchmarks against the current parser

4. **Milestone 18 — Shared strings and date/time core** — complete
   - Streaming shared-string writer
   - Deduplication modes: disabled, hash, bounded LRU
   - Excel serial date/time conversion
   - ISO date cells and workbook date epoch

5. **Milestone 19 — Parallel package pipeline** — complete
   - Parallel worksheet serialization
   - Optional parallel deflate per package entry
   - Deterministic package ordering
   - Configurable compression level and strategy

6. **Milestone 20 — Core compatibility completion** — complete
   - Preserve unknown package parts and relationships
   - Strict/transitional namespace handling
   - Malformed-file diagnostics and recovery options
   - ZIP64 and large-entry support
   - Cancellation, progress callbacks, and limits

## Core 100% definition

“Core 100%” means complete and robust handling of workbook, worksheet, row, cell, value, formula metadata, styles, dimensions, relationships, shared strings, dates, errors, package preservation, and both DOM and streaming I/O. It does not mean 100% parity with every optional openpyxl chart, pivot, macro, or drawing feature.
