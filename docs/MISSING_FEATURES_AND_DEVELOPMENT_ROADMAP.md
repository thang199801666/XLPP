> **Historical roadmap notice (v1.12.0):** this file contains legacy gap analysis retained for provenance. It is not the active source of truth. Use `docs/CURRENT_CAPABILITIES.md`, `docs/capabilities.json`, `ROADMAP.md`, and `PACKAGE_STATUS.md` for current status.

# XL++ — Missing Features and Development Roadmap

**Updated:** 2026-08-08  
**Historical baseline package:** `XLPP_Full_Project_P0ZH_VBA_Authoring_Preservation_v1.11.0`

**P0Z-H status note:** v1.11.0 retains P0Z-G hardening and substantially expands VBA authoring/preservation: Standard/Class/Document modules, event source, stable code names, module flags, project properties/conditional constants, binary export and signed/external-project safety. Excel Desktop host validation remains external.

**P0Z-E status note:** v1.8.0 keeps the P0Y/P0X feature/reliability baseline, further decomposes Formula/Chart internals, and makes C/Python/C# high-level parity an automated release gate; it is not a new feature-parity score.

**P0Y status note:** P0X closed the scope-defined 90% general-purpose editing-core gate at **90.7/100**; P0Y/v1.3.0 keeps that scope and hardens correctness, I/O transactions, ZIP/ZIP64, streaming limits, topology edits, model invariants and sanitizer coverage. Older P0/P0W gap lists below are retained as roadmap history where useful.
**Current automated baseline:** 174/174 test suites with the P0Z-H VBA corpus; strict/sanitizer release gates are rerun for each package. P0Z-G established the seeded Clang/libFuzzer baseline.  
**Primary objective:** evolve XL++ from a strong XLSX generator/streaming library into a dependable general-purpose Excel OOXML reader, editor, and round-trip preservation library.

---

## 1. Executive summary

XL++ already provides a substantial foundation for:

- creating new `.xlsx` and experimental `.xlsm` workbooks;
- cells, stored formulas plus the P0W calculation core, dates, errors, rich text and hyperlinks;
- common styles, named styles, dimensions, merges and worksheet views;
- tables, filters, sorting, validation and conditional formatting;
- legacy comments/notes, page setup and document properties;
- streaming reads/writes, ZIP64, progress and cancellation;
- basic images, charts, pivot tables and VBA project generation.

The library is **not yet feature-complete**, but the P0B–P0U preservation program has substantially closed the original object-loss risk for the covered image/chart/pivot relationship graph. The current engineering bottleneck has moved from raw preservation to **dependency-aware editing and structural correctness**.

P0V-A now tracks direct chart A1 source dependencies and can skip unchanged cache rebuilds or synchronize caches automatically on an opt-in save copy. The next critical work is transitive dependency tracking plus Excel-grade row/column reference rewriting. Microsoft Excel Desktop recovery-log validation remains the final external gate before declaring preservation P0 closed.

---

## 2. Current readiness

| Use case | Current status |
|---|---|
| Create reports containing data, formulas and common styles | Good |
| Import/export large tabular datasets using streaming | Good |
| Create image/chart/pivot output | Good for covered worksheet-source features; modern charts use ChartEx; advanced Pivot ecosystem remains partial |
| Add macros from VBA source text | Experimental; Excel host validation still required |
| Open an arbitrary workbook and modify a few cells without losing advanced objects | Not ready |
| Replace OpenPyXL, ClosedXML or EPPlus broadly | Not ready |
| Preserve complex enterprise workbooks | Not ready |

Approximate engineering maturity:

| Area | Estimated maturity |
|---|---:|
| Core workbook generation | 75–85% |
| Streaming tabular I/O | 80–90% |
| General workbook editing and preservation | 45–55% |
| Broad common Excel feature coverage | 50–60% |
| Advanced enterprise workbook compatibility | Below production target |

These values are engineering estimates, not formal compatibility scores.

---

# 3. Missing and incomplete features

## 3.1 Package graph and round-trip preservation

**Priority: P0 — critical**

### Current problems

- Worksheet drawing relationships can be lost after load/save.
- Existing images can disappear from the visible workbook.
- Existing charts can disappear despite their XML parts remaining in the archive.
- Existing pivot table relationships and pivot cache references can be removed.
- Unsupported package parts may remain as orphaned files.
- Unknown parts are not consistently preserved as a connected dependency graph.
- Tests that only check whether a ZIP part exists can pass even when Excel no longer sees the object.

### Required capabilities

- Parse every `.rels` part into a package-wide relationship graph.
- Represent source part, relationship ID, relationship type, target and target mode.
- Resolve relative targets according to OPC rules.
- Preserve all untouched parts and relationships byte-for-byte where practical.
- Rebuild only the subgraph affected by an edited feature.
- Detect missing targets, duplicate relationship IDs and orphaned internal parts.
- Preserve unknown extension parts and extension relationships.
- Preserve content-type overrides/defaults for untouched parts.
- Support safe relationship ID allocation without renumbering unrelated objects.

### Completion criteria

- An image-bearing workbook retains the same visible image count after round-trip.
- A chart-bearing workbook retains the same visible chart count after round-trip.
- A pivot-bearing workbook retains the same pivot count and cache references.
- No internal relationship points to a missing part.
- No formerly reachable object becomes orphaned.
- Independent readers detect all preserved objects.
- Microsoft Excel opens all fixtures without recovery logs.

---

## 3.2 Images and drawing anchors

**Priority: P0/P1**

### Existing support

- Basic image generation.
- PNG/JPEG media packaging.
- Basic worksheet drawing relationship generation.
- Simple anchor placement.

### Missing or incomplete

- Reading existing worksheet drawings into an object model.
- Preserving existing drawing relationships during round-trip.
- One-cell anchor parsing and editing.
- Two-cell anchor parsing and editing.
- Absolute anchor support.
- Correct EMU conversions.
- Image width/height and aspect-ratio control.
- Position offsets inside cells.
- Cropping.
- Rotation and flip.
- Transparency handling.
- Alt text, title and descriptive metadata.
- Hyperlinks attached to images.
- Broad media format preservation.
- Multiple drawings per worksheet.
- Z-order.
- Safe handling of duplicated media resources.

### Required tests

- Read an image created by Excel.
- Read images created by OpenPyXL and LibreOffice.
- Round-trip one-cell, two-cell and absolute anchors.
- Round-trip multiple images on one worksheet.
- Resize, move, remove and replace an existing image.
- Preserve unknown drawing properties.
- Verify visible image count with an independent reader.

---

## 3.3 Charts

**Priority: P0/P1/P2**

### Existing support

Basic generation for a subset such as:

- bar;
- line;
- pie;
- scatter;
- doughnut;
- radar;
- area;
- bubble.

### Critical gaps

- Existing charts are not safely preserved during round-trip.
- No complete chart parser.
- Limited series and axis model.
- Limited styling and layout fidelity.

### Missing chart features

- Reading and preserving existing charts.
- Chart title rich text.
- Axis titles.
- Major/minor units.
- Number formats.
- Secondary axes.
- Combined charts.
- 3D bar, line, area and pie.
- Stock charts.
- Surface charts.
- Projected pie and bar-of-pie.
- Data labels.
- Error bars.
- Trendlines.
- Up/down bars.
- High-low lines.
- Drop lines.
- Data tables.
- Manual layout.
- Legend positioning and overlays.
- Series markers and advanced line/fill styling.
- Chart themes.
- External data references.
- Chartsheets.
- Extension-list preservation.

### Completion criteria

- Read and preserve Excel-created charts without changing visible output.
- Modify title, series range and selected styling on an existing chart.
- Preserve unsupported chart properties.
- Independent readers report unchanged chart counts.
- Excel opens and saves fixtures without repairs.

---

## 3.4 Pivot tables and pivot caches

**Priority: P0/P1/P2**

### Existing support

- Basic generated pivot cache definition.
- Basic pivot cache records.
- Basic row and data fields.
- Basic pivot table view creation.

### Critical gaps

- Existing pivots are not fully parsed.
- Existing pivot relationships can disappear after round-trip.
- Existing cache definitions and cache records are not safely preserved.
- Excel compatibility remains fragile.

### Missing capabilities

- Parse `pivotTableDefinition`.
- Parse `pivotCacheDefinition`.
- Parse `pivotCacheRecords`.
- Preserve external and worksheet data sources.
- Preserve cache IDs and workbook pivot cache references.
- Read/edit row, column, page and data fields.
- Multiple data fields.
- Aggregations: sum, count, average, min, max, product, variance and standard deviation.
- Number formats.
- Sorting and filtering.
- Grouping by date, number and manual group.
- Calculated fields and calculated items.
- Show-values-as behavior.
- Subtotals and grand totals.
- Compact, outline and tabular layouts.
- Repeat item labels.
- Blank/error handling.
- Pivot styles.
- Refresh-on-load and save-data semantics.
- Multiple pivot tables sharing one cache.
- Pivot charts.
- Slicers and timelines.

### Required tests

- Preserve an Excel-created pivot through load/save.
- Preserve two pivots sharing one cache.
- Preserve multiple caches.
- Modify a pivot caption without rebuilding its entire cache.
- Add/remove data fields.
- Excel COM refresh test.
- Detect recovery logs automatically.

---

## 3.5 VBA and macro projects

**Priority: P1/P2/P3 — source-only macro authoring substantially expanded in P0Z-H**

### Existing support

- Attach/remove and byte-preserve an existing `vbaProject.bin`.
- Generate a real CFB/OVBA `vbaProject.bin` from VBA source.
- Author/read/update/remove Standard modules.
- Author/read Class modules and preserve `MODULEREADONLY` / `MODULEPRIVATE`.
- Author `ThisWorkbook` and worksheet Document modules, including event procedures.
- Stable worksheet `codeName` identity through load/save, reorder, delete and copy.
- Copying a worksheet copies its authored document-event source under a fresh code name.
- Project Name, Description, Help File, Help Context ID and conditional-compilation constants.
- Raw `vbaProject.bin` byte access/export.
- Signature-part detection and source-editability state.
- Signed projects reject source/metadata rewrites that would invalidate the signature.
- Externally supplied projects reject destructive source/metadata rewrites, preserving unknown references/forms/protection/signature state.
- Macro-enabled `.xlsm` package generation and reload.

### Remaining gaps

- Repeated validation on Microsoft Excel Desktop and an Excel-version compatibility matrix.
- UserForms / designer modules, designer storages and FRX binary resources.
- ActiveX controls and form controls.
- Custom reference/type-library authoring and reference repair across machines.
- VBA project password/locking authoring or removal.
- Digital-signature generation/re-signing and certificate-chain management.
- VBA stomping detection and compiled p-code/source-consistency analysis.
- Macro-enabled template `.xltm` authoring/host validation.
- Full Unicode/code-page authoring beyond the current source-only MBCS-oriented project builder.
- Large real-world VBA/UserForm corpora and automated Excel recovery-log checks.

### Completion criteria

- Macro appears in `Alt+F8` and executes in supported Excel Desktop versions.
- Workbook/worksheet event handlers execute after XL++ generation.
- External macro projects round-trip byte-for-byte when untouched.
- Signed macro projects remain byte-preserved when untouched and are never silently invalidated by a source rewrite.
- UserForm fixtures preserve and, eventually, author both code and designer/FRX resources.
- Custom project references open without repair across the supported Windows/Office matrix.

---

## 3.6 File encryption and protection

**Priority: P2 compatibility depth — P0X password-encryption core substantially complete**

### Existing support

- Legacy worksheet protection and workbook structure protection.
- Password-to-open OOXML encryption through load/save options.
- Agile AES-256-CBC/SHA-512 read/write, password verification, segmented package encryption and HMAC integrity.
- Standard AES-128/192/256 + SHA-1 **read/write**, 50,000-round derivation and password verification.
- `EncryptionInfo` inspection, `EncryptedPackage`, DataSpaces and OLE/CFB miniFAT/FAT/DIFAT support.
- OS CSPRNG, constant-time verification and malformed-container bounds/cycle hardening.
- Path/stream APIs, password rotation/removal, >9 MiB DIFAT regression.
- Independent LibreOffice interoperability: XL++ Agile output, LibreOffice Standard input, and XL++ Standard AES-128 output.
- Python/C/C# surfaces include password workflow and Standard/Agile mode selection; native C ABI is regression-tested.

### Remaining capabilities

- Alternate Agile cipher/hash profiles.
- Certificate/private-key Agile key encryptors.
- Microsoft Excel Desktop encrypted-fixture/recovery-dialog validation.
- Modern worksheet/workbook protection hashes and macro/template/signature compatibility corpus.

---

## 3.7 Structural row and column editing

**Priority: P2 compatibility depth — P0X transactional core implemented**

### Existing support

Workbook-level insert/delete row/column transactions now share a unified A1 reference translator and update supported dependencies across cells/formulas, shared/array/dynamic metadata, defined names, merged ranges, dimensions, freeze panes, AutoFilter/sort state, conditional formatting, data validation, tables, print ranges/titles, hyperlinks, chart references/caches, drawing anchors and Pivot source/location. Invalidated references become `#REF!` and can trigger full rollback. Successful transactions apply in place so existing worksheet/binding handles stay stable.

### Remaining correctness depth

- Worksheet rename/move/copy topology translation.
- Page-break/comment/unsupported DrawingML semantics beyond preservation.
- Deeper external-link and calc-chain topology updates.
- Microsoft Excel Desktop comparison matrix for boundary insertion/deletion semantics.

### Completion-depth criteria

- Expand Excel-created integration fixtures containing every supported object class.
- Compare boundary and mixed-reference behavior against Excel Desktop.
- Preserve unsupported objects and prove zero repair dialogs/recovery logs.

---

## 3.8 Formula system

**Priority: P2 catalog/interop depth — P0X core-family matrix complete**

### Existing support

- Formula text/cached values, shared/array/dynamic-array metadata and calculation metadata.
- Parser/operators, scalar/error literals and two-dimensional array constants.
- A1/range/cross-sheet references, defined names and structured table references.
- Dynamic-array/spill families (`SEQUENCE`, `FILTER`, `SORT`, `UNIQUE`, stacking/selection/shape helpers) with `#SPILL!` collision handling.
- Reference-returning `INDIRECT`/`OFFSET`, `LET`, lookup/criteria/statistical/text/date/math/financial families.
- Recursive dependencies, public dependency graph, circular detection and opt-in iterative fixed-point calculation.
- External-workbook references through an explicit C++ resolver callback; the core never silently opens another workbook.
- Cached-result updates, diagnostic no-mutation calculation and calculate-before-save.
- Structural reference translation across workbook transactions.

The scope-defined P0X Formula matrix is **26/26 capability families**. This is not a claim to implement every function in the Excel catalog.

### Remaining depth

- Broader specialized financial/database/engineering/cube catalog and `LAMBDA` ecosystem.
- More complete Excel 365 implicit-intersection/coercion edge cases.
- Microsoft Excel Desktop comparison corpus.
- External resolver parity in non-C++ bindings.

---

## 3.9 Styles, themes and formatting fidelity

**Priority: P2**

### Existing support

- Fonts.
- Pattern fills.
- Borders.
- Alignment.
- Number formats.
- Cell protection.
- Named styles.
- Differential styles.

### Missing or partial

- Full theme parsing.
- Theme color resolution.
- Indexed color palette.
- Tint and shade calculations.
- Gradient fills.
- Diagonal borders.
- Vertical/horizontal internal borders.
- Start/end borders.
- Rich border color types.
- Phonetic properties.
- East Asian font settings.
- Charset/family/scheme fidelity.
- Style inheritance.
- Complete `cellStyleXfs`, `cellXfs`, `dxfs` and named style fidelity.
- Quote prefix.
- Pivot-specific style information.
- Complete alignment flags.
- Conditional formatting extension styles.

### Required tests

- Excel-created themed workbook.
- Theme color plus tint.
- Gradient fill.
- Complex borders.
- Named style inheritance.
- Round-trip visual comparison using Excel rendering screenshots or PDF export.

---

## 3.10 AutoFilter and sorting

**Priority: P2**

### Existing support

- Basic equality/comparison filters.
- Custom filter operators.
- Basic sorting.

### Missing

- Date group items.
- Dynamic filters.
- Color filters.
- Icon filters.
- Top10 filters.
- Blank/nonblank special cases.
- Multiple sort conditions.
- Case-sensitive sort.
- Sort by cell color.
- Sort by font color.
- Sort by icon.
- Custom lists.
- Locale/collation behavior.
- Table-specific filter preservation.
- Extension-list preservation.

---

## 3.11 Conditional formatting

**Priority: P2**

### Existing support

- Formula rules.
- Cell-is rules.
- Color scales.
- Data bars.
- Icon sets.
- Differential styles.

### Missing or incomplete

- Top/bottom rules.
- Above/below average.
- Duplicate/unique values.
- Text contains/begins/ends.
- Time-period rules.
- Blanks/errors rules.
- Advanced data bar extensions.
- Negative data bar colors.
- Axis position/color.
- Gradient/solid data bars.
- Custom icon sets.
- Stop-if-true semantics across overlapping ranges.
- Full priority normalization.
- Extension-list preservation.
- Formula localization and relative-reference fidelity.

---

## 3.12 Data validation

**Priority: P2**

### Existing support

- Common validation types.
- Common operators.
- Prompt and error messages.
- Basic list validation.

### Missing or incomplete

- Extension-based validation used by newer Excel versions.
- Cross-sheet list sources.
- Defined-name list sources.
- Dynamic array list sources.
- Full formula preservation.
- IME mode.
- Complex multi-range validation.
- Interoperability corpus from Excel, LibreOffice and OpenPyXL.
- Structural edit updates.

---

## 3.13 Workbook and worksheet state

**Priority: P1/P2**

### Missing or incomplete

- Active worksheet.
- Selected sheets.
- Multiple selected tabs.
- First visible tab.
- Sheet tab ratio.
- Sheet visibility:
  - visible;
  - hidden;
  - veryHidden.
- Workbook window position and size.
- Minimized/maximized state.
- Multiple workbook views.
- Custom views.
- Sheet views collection.
- Show formulas.
- Show zeros.
- View type.
- Page-break preview.
- Page layout view.
- Chartsheet state.

---

## 3.14 Comments and collaboration

**Priority: P2/P3**

### Existing support

- Legacy comments/notes.
- Authors.
- Rich text in note content.
- VML note shapes.

### Missing

- Threaded comments.
- People metadata.
- Replies.
- Resolved state.
- Mentions.
- Modern comments conversion.
- Preservation when legacy and threaded comments coexist.
- Comment shape editing.
- Comment visibility and position fidelity.
- Cross-application compatibility tests.

---

## 3.15 Shapes, text boxes and general drawings

**Priority: P3**

### Missing

- AutoShapes.
- Text boxes.
- Lines and arrows.
- Connectors.
- Groups.
- Shape rotation.
- Shape fill and outline.
- Rich text inside shapes.
- Shape hyperlinks.
- Shape locking.
- Z-order.
- WordArt.
- SmartArt preservation.
- Embedded diagrams.
- General DrawingML object graph.

At minimum, untouched shapes and SmartArt must be preserved before editing support is attempted.

---

## 3.16 Sparklines

**Priority: P2/P3**

### Missing

- Line sparklines.
- Column sparklines.
- Win/loss sparklines.
- Sparkline groups.
- Axis settings.
- Marker colors.
- Hidden/empty cell behavior.
- Date axis.
- Grouped editing.
- Extension-list preservation.

---

## 3.17 Slicers and timelines

**Priority: P3**

### Missing

- Slicer caches.
- Slicer objects.
- Pivot slicers.
- Table slicers.
- Timeline caches.
- Timeline objects.
- Connections between slicers and multiple pivots.
- Slicer styles.
- Preservation of unsupported slicer extensions.

---

## 3.18 External links, connections and query objects

**Priority: P2/P3**

### Missing or very limited

- External workbook links.
- External name definitions.
- Connection definitions.
- Query tables.
- Web queries.
- Text import connections.
- ODBC/OLE DB connections.
- Power Query/Mashup preservation.
- Refresh properties.
- Credentials metadata.
- Background refresh settings.
- Data model relationships.
- Linked data types.

Initial target should be preservation, not full editing.

---

## 3.19 OLE, ActiveX and form controls

**Priority: P3**

### Missing

- Embedded OLE packages.
- Linked OLE objects.
- ActiveX controls.
- Form controls.
- Control properties.
- Control relationships.
- VML control drawings.
- Control macros.
- Binary control streams.
- Safe preservation of unsupported controls.

---

## 3.20 Page layout and printing

**Priority: P2**

### Existing support

- Page setup.
- Margins.
- Print options.
- Headers and footers.
- Print area.
- Print titles.

### Missing or incomplete

- Horizontal page breaks.
- Vertical page breaks.
- Manual page-break collections.
- Print quality.
- First page number edge cases.
- Advanced paper sizes.
- Printer settings binary part.
- Header/footer images.
- Different first/odd/even header/footer fidelity.
- Page setup properties from multiple applications.
- Chartsheet printing.

---

## 3.21 Templates and file formats

**Priority: P1/P2**

### Existing support

- `.xlsx`.
- Experimental `.xlsm`.

### Missing first-class workflows

- `.xltx`.
- `.xltm`.
- Template content types.
- Template-aware save-as behavior.
- Macro-enabled templates.
- Preserve template status when editing.
- Detect format from package rather than only extension.
- CSV/TSV convenience layer, if included in product scope.
- Binary `.xlsb` is not recommended for the near-term roadmap.

---

## 3.22 Bindings and public packaging

**Priority: P2**

### Missing

- Production Python binding.
- Production C#/.NET binding.
- Stable C ABI where needed.
- Versioned ABI policy.
- Package managers:
  - vcpkg;
  - Conan;
  - NuGet for .NET binding;
  - PyPI wheels for Python binding.
- Installable CMake package.
- Exported CMake targets.
- API documentation.
- Migration guide.
- Semantic compatibility policy.
- Reproducible release workflow.
- Signed release artifacts.
- Public license and notices review.
- Security policy.
- Contribution guide.
- Issue and feature templates.

---

# 4. Development roadmap

## Phase 0 — establish a correctness baseline

**Priority:** Immediate  
**Goal:** stop adding broad features until preservation failures are measurable and reproducible.

### Work items

- Create a canonical external fixture corpus.
- Include files created by:
  - Microsoft Excel;
  - LibreOffice;
  - OpenPyXL;
  - ClosedXML or EPPlus where licensing permits fixtures.
- Add fixture categories:
  - image;
  - chart;
  - pivot;
  - macros;
  - comments;
  - themes;
  - external links;
  - unknown extension parts.
- Build a relationship graph inspector.
- Build an orphan-part detector.
- Build a content-type consistency validator.
- Record object counts before and after round-trip.
- Store repaired-record logs from Excel COM tests.

### Deliverables

- `tests/fixtures/excel/`
- `tests/fixtures/libreoffice/`
- `tests/fixtures/openpyxl/`
- `tests/roundtrip/`
- `tools/xlpp-package-validator/`
- `docs/COMPATIBILITY_MATRIX.md`

### Exit criteria

- Every current object-loss bug has a deterministic failing test.
- Tests fail if an object part exists but is unreachable.
- Test reports separate generation success from preservation success.

---

## Phase 1 — OPC package and relationship graph

**Priority:** P0  
**Goal:** guarantee that untouched package objects remain connected.

### Work items

- Introduce `PackagePart`.
- Introduce `PackageRelationship`.
- Introduce `RelationshipGraph`.
- Resolve internal/relative targets.
- Preserve external targets.
- Preserve unknown parts.
- Preserve unknown relationships.
- Preserve content types.
- Allocate IDs without unnecessary renumbering.
- Detect cycles safely.
- Detect orphaned internal parts.
- Detect dangling relationships.
- Support graph-level copy-on-write.
- Define ownership rules between workbook model and raw package parts.

### Suggested API/internal structure

```text
Package
├── ContentTypes
├── Parts
│   ├── URI
│   ├── ContentType
│   ├── RawBytes
│   ├── ParsedModel (optional)
│   └── DirtyState
└── Relationships
    ├── SourceURI
    ├── Id
    ├── Type
    ├── Target
    └── TargetMode
```

### Exit criteria

- Untouched chart/image/pivot parts and relationships remain byte-identical where possible.
- No P0 fixture loses visible objects.
- Package validator reports zero orphan/dangling errors.
- Excel opens every preservation fixture without repair.

---

## Phase 2 — drawing and image preservation

**Priority:** P0/P1  
**Goal:** read, preserve and edit worksheet images safely.

### Work items

- Parse worksheet drawing relationships.
- Parse one-cell anchors.
- Parse two-cell anchors.
- Parse absolute anchors.
- Parse embedded image relationships.
- Introduce stable drawing object IDs.
- Add move/resize/remove/replace APIs.
- Add EMU conversion utilities.
- Preserve unsupported drawing properties.
- Support duplicate media deduplication optionally.
- Preserve media file names where possible.

### Exit criteria

- Images survive unrelated cell edits.
- Image count and anchor locations remain stable.
- Modified images are visible in Excel.
- Multiple-image and multi-drawing fixtures pass.

---

## Phase 3 — chart preservation and editing foundation

**Priority:** P1  
**Goal:** safely preserve all charts and edit a controlled subset.

### Work items

- Parse chart relationships.
- Parse chart-space metadata.
- Preserve unsupported chart XML.
- Build minimal common chart object model:
  - title;
  - legend;
  - plot area;
  - axes;
  - series;
  - data references.
- Add chart dirty-region serialization.
- Preserve extension lists.
- Implement chartsheet preservation.
- Add limited edit APIs for:
  - title;
  - series formula;
  - category formula;
  - placement;
  - size.

### Exit criteria

- Excel-created charts survive round-trip unchanged.
- Supported edits do not erase unsupported chart properties.
- Independent readers retain chart counts.
- Visual regression fixtures show no unexpected differences.

---

## Phase 4 — pivot preservation and reader

**Priority:** P1  
**Goal:** safely preserve existing pivot tables before expanding creation features.

### Work items

- Parse workbook pivot caches.
- Parse cache definitions and records.
- Parse pivot table definitions.
- Link pivot tables to caches.
- Preserve shared cache relationships.
- Preserve source references.
- Add read-only public pivot inspection APIs.
- Add controlled caption/style/refresh edits.
- Preserve unsupported grouping and calculation records.
- Add cache consistency validator.

### Exit criteria

- One and multiple pivot fixtures survive round-trip.
- Shared caches remain shared.
- Excel refresh works without recovery.
- Basic pivot properties can be read reliably.

---

## Phase 5 — dependable structural editing

**Status:** P0X core completed; P2 compatibility-depth validation remains.  
**Goal:** keep row/column insertion/deletion safe across the supported object graph and broaden Excel Desktop parity.

### Work items

- **Completed:** unified A1/range reference translator with absolute/mixed references.
- **Completed:** workbook transaction rewrites formulas, names, tables, validation/CF, merges, hyperlinks, chart/image anchors, Pivot sources and print ranges.
- **Completed:** transactional preflight/rollback and stable live handles.
- **Next:** worksheet rename/move/copy topology, deeper unsupported-object semantics, Excel Desktop boundary-behavior comparison.

### Exit criteria

- Structural edit integration workbook passes all object checks.
- Results match Excel for a defined test matrix.
- No silent object deletion.

---

## Phase 6 — workbook state, templates and encryption

**Priority:** P1  
**Goal:** cover core document lifecycle expected from an editor library.

### Work items

- Active sheet and selected sheet APIs.
- Hidden and veryHidden sheet state.
- Workbook view/window model.
- `.xltx` and `.xltm`.
- File-open encryption: P0X Agile AES-256/SHA-512 and Standard AES-128/192/256 + SHA-1 read/write complete; alternate/certificate profiles remain.
- Password add/change/remove.
- Modern protection hashes.
- Format detection from package metadata.
- Preserve macro/template/encryption combinations.

### Exit criteria

- Encrypted fixtures open with correct passwords.
- Incorrect passwords fail predictably.
- Templates retain template content types.
- Hidden/veryHidden states survive round-trip.

---

## Phase 7 — common Excel parity

**Priority:** P2  
**Goal:** close the largest feature gaps encountered in normal business workbooks.

### Workstreams

#### 7A — filters and sorting

- date-group;
- dynamic;
- color;
- icon;
- Top10;
- multi-condition sorts;
- color/icon sorts.

#### 7B — conditional formatting

- top/bottom;
- average;
- unique/duplicate;
- text;
- time periods;
- blank/error;
- advanced data bars.

#### 7C — styles and themes

- theme colors;
- indexed colors;
- tint;
- gradients;
- advanced borders;
- style inheritance.

#### 7D — page layout

- page breaks;
- printer settings preservation;
- header/footer images;
- chartsheet print settings.

#### 7E — collaboration

- threaded comments;
- people;
- replies;
- resolved state.

### Exit criteria

- Compatibility matrix documents each supported subtype.
- Excel, LibreOffice and OpenPyXL fixture results are recorded.
- Unsupported extensions are preserved.

---

## Phase 8 — chart and pivot expansion

**Priority:** P2  
**Goal:** extend editing breadth after preservation is stable.

**P0Z-F update:** chart-family generation is complete across the current Excel Insert-chart families and the non-OLAP Pivot authoring surface is substantially expanded. Remaining bullets below now primarily represent Pivot ecosystem/host-validation work rather than missing base chart families.

### Chart work

- [x] 3D charts.
- [x] Stock charts.
- [x] Surface charts.
- [x] Combined charts.
- [x] Secondary axes.
- [x] Data labels.
- [x] Trendlines.
- [x] Error bars.
- [x] Data tables.
- [x] Advanced layouts, including ChartEx-backed current Excel families.
- [ ] Chartsheets.

### Pivot work

- [x] column/page fields;
- [x] multiple data fields;
- [x] worksheet-source cache records and lifecycle metadata;
- [x] grouping;
- [ ] calculated fields/items;
- [x] legacy and Office 2010+ `Show Values As`;
- [x] subtotals;
- [x] styles and display behavior;
- [x] filters, hidden items and multi-axis item tuples;
- [ ] pivot charts;
- [ ] OLAP/Data Model caches, slicers and timelines.

### Exit criteria

- Each added chart/pivot subtype has:
  - generation test;
  - independent read test;
  - round-trip test;
  - Excel Desktop test.

---

## Phase 9 — advanced Excel ecosystem

**Priority:** P3  
**Goal:** support preservation and selected editing for advanced enterprise content.

### Work items

- Sparklines.
- Slicers.
- Timelines.
- Shapes.
- Text boxes.
- SmartArt preservation.
- OLE.
- ActiveX.
- Form controls.
- External links.
- Connections.
- Query tables.
- Power Query preservation.
- Data model preservation.
- VBA UserForms and signatures.

### Strategy

Preservation must precede editing. Unknown advanced parts should remain connected and byte-identical when the user edits unrelated cells.

---

## Phase 10 — bindings, packaging and public release hardening

**Priority:** P2/P3  
**Goal:** make the library consumable and maintainable as a public product.

### Work items

- Stable public API review.
- ABI policy.
- CMake install/export.
- vcpkg package.
- Conan package.
- Python binding.
- Python wheels.
- C# binding.
- NuGet packages.
- API documentation.
- Cookbook.
- Migration guides.
- Compatibility matrix.
- Security policy.
- Fuzzing.
- OSS-Fuzz or equivalent.
- Static analysis.
- Sanitizer CI.
- Windows/Linux/macOS CI.
- Excel COM integration CI on Windows.
- Release signing.
- SBOM.
- License/notice audit.

### Exit criteria

- Clean build from a fresh environment.
- Installable package verified by external consumer projects.
- Supported-platform matrix documented.
- Release artifacts reproducible.
- No critical preservation regressions in the compatibility corpus.

---

# 5. Test strategy

Every feature must be validated at several layers.

## Layer 1 — unit tests

Validate:

- model behavior;
- parser branches;
- serializer branches;
- input validation;
- mutation and clear/remove APIs;
- malformed XML handling.

## Layer 2 — independent fixtures

Fixtures must not be created only by XL++.

Use files produced by:

- Microsoft Excel;
- LibreOffice;
- OpenPyXL;
- other independent libraries when appropriate.

## Layer 3 — round-trip preservation tests

For every fixture:

1. inspect original package;
2. load through XL++;
3. modify one unrelated cell;
4. save;
5. compare object counts;
6. compare relationship graph;
7. compare content types;
8. inspect orphaned parts;
9. reopen using an independent reader.

## Layer 4 — Microsoft Excel Desktop integration

Use Windows CI with Excel COM automation to:

- open workbook;
- capture repair/recovery prompts or logs;
- save workbook;
- verify visible objects;
- verify pivot refresh;
- verify macros appear;
- execute test macros;
- export sheets/charts to PDF for visual regression where useful.

## Layer 5 — fuzzing and corruption testing

Targets:

- ZIP reader;
- XML parser;
- relationship resolver;
- shared strings;
- styles;
- formulas;
- drawings;
- pivot caches;
- CFB/OVBA;
- encrypted packages.

## Layer 6 — performance testing

Track:

- load time;
- save time;
- memory use;
- streaming throughput;
- shared-string performance;
- style deduplication;
- compression time;
- large workbook round-trip time.

Test sizes should include:

- 10 thousand cells;
- 1 million cells;
- multi-sheet workbooks;
- large shared-string tables;
- large media files;
- large VBA projects.

---

# 6. Definition of done

A feature is not considered complete merely because XL++ can generate XML for it.

A feature is complete only when all applicable conditions are met:

- Public API is documented.
- Writer test passes.
- Independent-reader test passes.
- Independent fixture can be read.
- Round-trip preserves the feature.
- Unrelated edits do not remove the feature.
- Remove/clear API is tested.
- Invalid input is tested.
- Strict and Transitional OOXML behavior is defined.
- Excel Desktop opens the output without recovery.
- LibreOffice behavior is recorded.
- OpenPyXL or another independent parser recognizes the result where applicable.
- Unknown extension data is preserved.
- Performance impact is measured.
- Security implications are reviewed.

---

# 7. Priority backlog

## P0 — correctness and data-loss prevention

- Relationship graph.
- Orphan/dangling relationship validator.
- Image preservation.
- Chart preservation.
- Pivot preservation.
- Excel recovery-log CI.
- Unknown part preservation.
- Round-trip object-count tests.

## P1 — dependable workbook editing

- Structural reference update engine.
- Full sheet/workbook state.
- Image editing.
- Basic chart editing.
- Pivot reader.
- Complete remaining encryption profiles/Excel Desktop validation.
- `.xltx/.xltm`.
- Macro host validation.
- External `vbaProject.bin` preservation.

## P2 — common feature parity

- Advanced filters.
- Advanced conditional formatting.
- Themes and advanced styles.
- Page breaks.
- Threaded comments.
- More chart types.
- Deeper pivot features.
- External-link preservation.
- Python and C# bindings.
- Package manager support.

## P3 — advanced ecosystem

- Slicers.
- Timelines.
- Sparklines.
- Shapes/text boxes.
- SmartArt.
- OLE.
- ActiveX.
- Form controls.
- Power Query.
- Data model.
- VBA UserForms and signatures.

---

# 8. Suggested milestone sequence

| Milestone | Main result |
|---|---|
| M1 — Preservation Core | No chart/image/pivot loss during unrelated edits |
| M2 — Safe Editor | Structural edits and workbook states are dependable |
| M3 — Secure Lifecycle | Encryption and templates are supported |
| M4 — Common Parity | Advanced formatting/filter/comment gaps are reduced |
| M5 — Visual Objects | Charts, pivots and drawings reach practical editing depth |
| M6 — Enterprise Preservation | Advanced objects and connections survive round-trip |
| M7 — Public SDK | Bindings, packages, docs and release process are production-ready |

Version numbers should be assigned according to the project's existing version policy. Milestone completion should be tied to acceptance criteria rather than only API count.

---

# 9. Recommended next development batch

The next batch should focus only on P0 preservation.

## Scope

1. Add package relationship graph classes.
2. Load all worksheet/workbook relationships.
3. Preserve drawing relationships.
4. Preserve pivot relationships.
5. Preserve unknown relationships.
6. Add orphan-part validation.
7. Add three independent regression fixtures:
   - image;
   - chart;
   - pivot.
8. Assert object counts before and after load/save.
9. Add a report listing:
   - added parts;
   - removed parts;
   - changed parts;
   - orphaned parts;
   - dangling relationships.

## Explicitly defer

- New chart types.
- New pivot layout types.
- New VBA features.
- Slicers.
- Shapes.
- Bindings.

These should not be prioritized until load/save no longer destroys existing workbook objects.

---

# 10. Final target

XL++ can be considered a dependable general-purpose Excel library when it can:

- generate new workbooks;
- stream very large workbooks;
- read common Excel files;
- edit cells and structures;
- preserve all untouched package features;
- safely handle images, charts, pivots and macros;
- support encrypted documents and templates;
- pass independent OOXML validation;
- pass Microsoft Excel Desktop integration tests;
- expose a stable, documented and installable public API.

Until the preservation roadmap is complete, the library should be described as:

> A high-performance C++ XLSX generation and streaming library with growing workbook editing support.

It should not yet be described as fully feature-complete or as a safe universal replacement for OpenPyXL, ClosedXML or EPPlus.
