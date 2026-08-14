# Architecture

## High-level design

```
┌──────────────────────────────────────────────────┐
│                   Public API                      │
│  Workbook │ Worksheet │ Cell │ Style │ Range      │
│  Chart │ PivotTable │ Table │ DefinedName         │
├──────────────────────────────────────────────────┤
│                  Core Engine                      │
│  ┌─────────────┐ ┌──────────┐ ┌───────────────┐ │
│  │ XML Engine  │ │ZIP Engine│ │Cell Storage   │ │
│  │ (SIMD scan) │ │(mmap I/O)│ │(uint64_t key) │ │
│  └─────────────┘ └──────────┘ └───────────────┘ │
├──────────────────────────────────────────────────┤
│              Platform Abstraction                 │
│  ┌──────────┐ ┌───────────┐ ┌────────────────┐  │
│  │ MappedFile│ │ThreadPool │ │ Compression    │  │
│  │Win+POSIX │ │           │ │ (zlib)         │  │
│  └──────────┘ └───────────┘ └────────────────┘  │
└──────────────────────────────────────────────────┘
```

## Layer breakdown

### 1. Public API (`include/XLPP/`)

Header-only, no external dependencies beyond C++20 STL. Classes:

- `Workbook` — container for worksheets, named styles, defined names, properties
- `Worksheet` — grid of cells, layout, filters, tables, charts
- `Cell` — value (variant), formula, style, hyperlink, comment
- `Style` / `Font` / `Fill` / `Border` / `Alignment` — fluent chaining
- `CellReference` — A1 ↔ (row, col) conversion
- `CellRange` — rectangular range operations
- `DateTime` — Excel serial date conversion
- `StreamingWorkbookWriter` / `StreamingWorkbookReader` — large file I/O
- `SaveOptions` / `LoadOptions` — configuration

### 2. XML Engine (`src/XLPP/XML/`)

- `XmlScanner` — allocation-free element iteration using `string_view`
- `SimdScan.h` — SSE2-accelerated byte scanning for tag/attribute search
- `XmlPullReader` — streaming XML reader for ZIP entry content
- `XmlUtilities` — helpers: escape, unescape, tag extraction, numeric parsing

**Performance**: SIMD first-char filter scans 16 bytes/iteration for tag name matching. Benchmark shows ~6× speedup vs baseline string search.

### 3. ZIP Engine (`src/XLPP/Packaging/`)

- `ZipArchive` — in-memory ZIP creation with parallel deflate compression
- `ZipArchiveReader` — streaming ZIP reader with central directory index
- `MappedFile` — cross-platform memory-mapped I/O (Windows `CreateFileMappingW` / POSIX `mmap`)
- `ZipEntrySource` — pull-based decompressed entry reader

**Key optimizations**:
- MappedFile closes the file handle after mapping (allows temp file cleanup)
- Central directory is parsed directly from mapped memory (zero-copy)
- Entry data is read via pointer arithmetic on the mapped view

### 4. Cell Storage (`src/XLPP/Worksheet/Worksheet.cpp`)

Cells stored in `std::map<uint64_t, Cell>` with composite key:

```cpp
uint64_t key = (row << 20) | column;
```

This gives:
- **Row-major ordering** naturally (no sort needed at save)
- **Integer key comparison** (faster than string comparison)
- **No string allocation** per cell (key is 8 bytes)

### 5. Save Pipeline

```
Workbook::save()
  ├── Build StyleCatalog       (dedup styles by hash+equality)
  ├── Build SST index          (dedup shared strings)
  ├── serializeSheets()        (parallel per sheet or per row)
  │     └── sheetXml()         (serialize one sheet to XML string)
  ├── Build ZIP entries        (styles, SST, workbook, rels, etc.)
  └── ZipArchive::save()       (parallel compression + write)
```

**Multi-level parallelism**:
- `parallelSheets + parallelWorkers > 1`: sheets distributed across threads
- `parallelRows + single sheet`: rows within the sheet split across threads
- ZIP entry compression: parallel deflate in `ZipArchive::save()`

### 6. Streaming Engine (`src/XLPP/Streaming/`)

For very large files (100K+ rows):

- `StreamingWorkbookWriter` — appends rows directly to disk spool file, assembles ZIP on close
- `StreamingWorkbookReader` — pull-based XML parsing from decompressed ZIP entries
- `SharedStringsReader` — lazy-load SST with `std::once_flag` for thread safety

## Memory management

- Cells stored **by value** in map nodes
- Styles composed **by value** (no shared_ptr, no flyweight)
- Default style objects have zero cost at save-time (skipped via `isDefault()`)
- Save-time style deduplication via hash-based `StyleCatalog`

## Thread safety

- **Read**: thread-safe for concurrent read (const operations)
- **Write**: not thread-safe (single-threaded modification)
- **Save**: multi-threaded internally, safe to call from any thread
- **Streaming**: not thread-safe for write operations

## File format support

| Format | Read | Write | Notes |
|--------|------|-------|-------|
| `.xlsx` | ✅ | ✅ | Full support |
| `.xlsm` | ✅ | ✅ | Macros preserved via part preservation |
| `.xlsb` | ❌ | ❌ | Not yet supported |
| `.xls` | ❌ | ❌ | Binary format not supported |
| Strict OOXML | ✅ | ✅ | `SaveOptions::strictNamespace` |
| ZIP64 | ✅ | ✅ | >4GB packages |
