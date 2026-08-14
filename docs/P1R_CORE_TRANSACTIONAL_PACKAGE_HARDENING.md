# P1R-A Core Transactional Load & Package Hardening

P1R-A continues the P1Q malformed-input work at the materialized `Workbook::load()` boundary. The goal is a strong exception guarantee for load, bounded stream/model materialization, strict parsing for semantic numeric fields, and consistent OPC relationship handling without weakening preservation of relationship types XL++ does not model.

## Transactional load

`Workbook::load(path, options)` now constructs a candidate workbook and commits it only after the complete package/model load succeeds. A package/XML/relationship/numeric/resource exception therefore leaves the caller's existing workbook unchanged. `load(std::istream&, options)` inherits the same guarantee after its bounded temporary-file bridge.

Stream input/output now lives in `WorkbookIO.cpp` with RAII temporary-file cleanup. Input is copied in 64 KiB chunks, `maxFileBytes` is enforced before full materialization, cancellation is checked while copying, and source/destination I/O failures are surfaced. Temporary files are removed on size-limit, cancellation and destination-write failures.

## Model resource limits

`LoadOptions` adds four model-level guards (0 means unlimited):

- `maxWorksheets`
- `maxCells` (total unique materialized worksheet cells)
- `maxSharedStrings`
- `maxDefinedNames`

These complement P1Q ZIP byte/entry limits. They protect against small compressed packages that expand into very large in-memory object graphs.

## Materialized XML and numeric parsing

The in-memory XML utility is quote-aware for opening-tag termination and skips comments, CDATA and processing instructions while locating elements. Same-name nested elements are balanced so pseudo closing tags inside lexical markup cannot truncate the owning element.

Critical OOXML numeric fields now use exact `std::from_chars`-based parsing. Trailing garbage, overflow and non-finite floating values are rejected. Out-of-range shared-string/style/font/fill/border references no longer silently degrade to defaults.

Worksheet cell enumeration uses zero-copy `tagsForEach` rather than building a copied `vector<string>` for every `<c>` element. On the local 20,000 x 10 fixture this reduced peak RSS from roughly 121 MiB to 101 MiB (~16-17%) while end-to-end Debug load time stayed within roughly 1-2% of P1Q in the final same-machine check. The lexical scanner itself is intentionally stricter and is slower in isolation; the zero-copy load path offsets most of that cost.

## OPC package relationship hardening

`WorkbookPackageReader.cpp` centralizes relationship-map and content-type parsing for the materialized reader. Duplicate relationship IDs are rejected. Relationships actually consumed by workbook/sheet/table/comment/hyperlink owners must have the expected relationship type; internal relationships are normalized relative to their owner, may not escape the package root, and must reference an existing part when required.

Internal targets now reject general URI schemes such as `file:` and `mailto:`, not only strings containing `://`. `TargetMode="External"` hyperlinks remain supported and are not interpreted as ZIP paths.

Unsupported/unmodeled relationships remain preservation-friendly: an empty/unknown `Type` on an unconsumed relationship is not rejected merely because XL++ does not model it. This distinction was locked by the existing Chart-part preservation regression.

`[Content_Types].xml` now rejects duplicate `Override PartName` and duplicate `Default Extension` declarations rather than silently applying last-wins semantics.

## Verification

Final P1R-A validation:

- main regression: 198/198 suites, 3,632/3,632 checks PASS;
- P1P lazy-formula regression: PASS;
- P1Q malformed-input/ZIP/CFB hardening regression: PASS;
- P1R transactional/model/package hardening regression: PASS;
- C API smoke: PASS;
- standalone public-header build: PASS;
- installed `find_package(XLPP CONFIG)` consumer: PASS;
- source `add_subdirectory()` consumer: PASS;
- focused Clang ASan + UBSan + leak detection for P1Q/P1R hardening suites: PASS.

P1R-A does not change `Cell`/`Style` layout. The P1P-era C++ ABI rebuild requirement still applies when upgrading from a pre-P1P binary, but P1R itself does not introduce another model-layout break.

## Remaining core work

`Workbook.cpp` remains about 527 KiB / 8.8k lines. P1R extracted stream I/O and package relationship/content-type helpers, but strict materialized parsing added enough code that the main TU remains essentially the same size. P1S should move the worksheet/materialized package reader and writer/serializer blocks into dedicated translation units before adding more functionality to the monolith.
