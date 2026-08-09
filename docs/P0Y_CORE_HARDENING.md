> **P0Z-B/v1.5.0 supersession note:** P0Z-B preserves the P0Y reliability/feature behavior and supersedes the P0Z architecture layout with Workbook/Chart/Drawing codec decomposition. See `P0ZB_CORE_DECOMPOSITION.md`.

# P0Y — Core Hardening / v1.3.0

P0Y does not broaden the P0X 90.7/100 Excel-feature score. It hardens the same
native editing core so failure paths, malformed packages, large ZIP64 workbooks
and workbook-topology mutations behave predictably enough for production-style
workflows.

## Acceptance baseline

- Native regression: **171/171 suites, 3,072/3,072 checks PASS**.
- Full native suite under **AddressSanitizer + UndefinedBehaviorSanitizer: clean**.
- Strict-warning build promotes high-value control-flow/memory diagnostics to errors.
- P0X's 90.7/100 general-purpose editing-core feature matrix remains the feature baseline.

## 1. Transactional workbook I/O

### Strong load guarantee

Public `Workbook::load()` stages parsing/decryption into a temporary workbook and
commits only after success. A failed load therefore leaves the caller's previous
workbook state intact instead of partially clearing/replacing it.

### Atomic path save

`SaveOptions::atomicWrite` defaults to `true`. Path saves serialize to a unique
same-directory staging file and replace the destination only after serialization
succeeds. Failure cleans the staging file and does not truncate an existing target.
Stream saves retain normal stream semantics.

The replace operation is atomic at the filesystem rename/replace boundary. P0Y
does **not** claim full power-loss durability (`fsync` of file and parent directory)
on every filesystem.

## 2. Workbook invariants and validation

`Workbook::validate()` and `SaveOptions::validateBeforeSave` enforce modeled
invariants before serialization:

- valid UTF-8 worksheet names, Excel's 31-code-point limit and forbidden characters;
- case-insensitive worksheet-name uniqueness;
- defined-name uniqueness by Excel scope and case-insensitive name matching;
- valid local defined-name sheet scopes;
- table names/ranges and duplicate table names;
- Pivot-name conflicts.

A1 parsing now enforces Excel's `1,048,576 x 16,384` grid and the exact
`[$]COLUMN[$]ROW` grammar before cell-key generation. This also prevents invalid
large columns from overlapping row bits in the compact cell key.

## 3. ZIP, ZIP64 and malformed-package hardening

Both materialized and direct streaming readers now validate package structure
before dereferencing variable-length records:

- legal EOCD and ZIP64 locator/record geometry;
- single-disk constraints;
- central/local header bounds and matching names/methods;
- ZIP64 sentinel/extra-field requirements;
- duplicate entry names;
- supported compression/encryption flags;
- CRC32;
- actual inflated bytes versus declared uncompressed size;
- configured entry/file/aggregate/count limits.

P0Y adds direct ZIP64 support to `StreamingWorkbookReader` rather than falling
back to materializing the archive.

### Memory-safety defects found by sanitizers

The hardening pass found and fixed two defects that ordinary regression did not
expose:

1. `ZipEntrySource` owned a raw `z_stream*` but previously relied on default move
   operations, permitting two moved objects to believe they owned the same inflater.
2. Mutation of central-directory variable lengths could reach `std::string`
   construction before the complete record bounds were validated, producing a real
   ASan heap-buffer-overflow.

The central-directory parser now uses cursor/bounds invariants and the inflater
has explicit move ownership.

## 4. Streaming decompression budgets

`StreamingReaderOptions` exposes:

- `maxFileBytes`;
- `maxEntryBytes`;
- `maxTotalBytes`;
- `maxEntries`;
- `validateCellReferences`.

The direct inflater enforces an output budget while bytes are produced, not only
after decompression. A malicious entry cannot declare a tiny size and stream an
unbounded payload before detection.

## 5. Bounded-memory ZIP64 writes

File-backed ZIP entries remain file-backed in the ZIP64 planning path:

- stored entries are scanned for CRC and copied directly;
- deflated entries are streamed once into a temporary prepared backing file;
- the final archive copies from the prepared/source path;
- preparation files are scoped and removed after success or failure.

This removes a previous whole-file `readFile()` path that could defeat ZIP64's
large-file support by exhausting RAM.

## 6. Workbook topology correctness

### Dependency-aware rename

`Workbook::renameWorksheet()` rewrites modeled local worksheet qualifiers in:

- cell formulas and formula metadata;
- defined names;
- conditional-formatting and data-validation formulas;
- chart series/title/error-bar references;
- Pivot source/location references;
- internal hyperlinks.

Quoted/unquoted sheet names, apostrophes and supported 3-D qualifiers are handled.
Formula string literals and external-workbook qualifiers are left untouched. Rename
is applied in place after preflight so live worksheet handles retain identity.

### Worksheet removal

Removing a worksheet now invalidates supported local dependencies to `#REF!`,
removes local defined names owned by the sheet and compacts following
`localSheetId` values. External workbook references remain unchanged. Worksheet
erase can invalidate worksheet handles because the current public storage is a
`std::deque`; callers must reacquire handles after removal.

## 7. Defined-name semantics

Defined-name matching is case-insensitive and scoped. The same local name may
exist independently on different worksheets; conflicts are rejected only within
the same scope. Workbook/global lookup remains available and explicit scoped lookup
is supported.

## 8. Build and architecture hardening

P0Y adds CMake gates for:

- `XLPP_ENABLE_SANITIZERS` (ASan + UBSan on supported GCC/Clang builds; ASan on MSVC);
- `XLPP_ENABLE_STRICT_WARNINGS` with high-value safety diagnostics promoted to errors.

The large `Workbook.cpp` started to be decomposed into dedicated translation units
for chart-cache synchronization, workbook-model CRUD, VBA lifecycle and worksheet
rename/topology code. This is intentionally incremental to avoid ABI/behavior churn.

## Remaining core debt

> **P0Z-G/v1.10.0 supersession note:** items 1-5 and the worksheet-copy portion of item 7 below have since been closed by P0Z-B through P0Z-G. The regression monolith and Workbook/WorkbookCodec monoliths were decomposed, stable insertion handles were introduced, strict-warning output is clean, durable fsync/FlushFileBuffers save was added, and copy topology is dependency-aware. Long-running external fuzz/Excel-host corpora and worksheet-removal handle ABI design remain current.


P0Y materially raises reliability but does not call the core literally perfect.
The main remaining engineering debt is:

1. `Workbook.cpp` is still roughly 8k lines and needs further extraction.
2. `tests/XLPP.UnitTests/main.cpp` remains monolithic and should be split into
   subsystem translation units despite PCH reducing build cost.
3. Worksheet deletion can invalidate previously acquired worksheet handles due to
   container erase semantics; a stable-node storage design is a future ABI-level choice.
4. Strict-warning mode is safety-clean but still emits conversion/sign/style noise;
   the source is not advertised as globally warning-free.
5. Atomic save protects against partial serialization but does not yet provide a
   portable full `fsync` power-loss durability guarantee.
6. A larger Microsoft Excel Desktop repair-log/corrupt-package corpus and continuous
   fuzzing remain valuable compatibility gates.
7. Worksheet move/copy topology and rarer cross-object references can still be made
   more dependency-aware.

P0Y should therefore be treated as a **core reliability milestone**, not an excuse
to stop compatibility testing on representative production workbooks.
