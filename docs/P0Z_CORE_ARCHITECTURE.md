> **Superseded by P0Z-B/v1.5.0:** this document describes the first architecture-refactor step. See `P0ZB_CORE_DECOMPOSITION.md` for the current tree.

# P0Z — Core Architecture Refactor / v1.4.0

P0Z is an architecture-only milestone built on the P0Y/v1.3.0 hardened feature baseline. It intentionally keeps the public C++ API and the P0X 90.7/100 editing-core feature scope stable while reorganizing implementation ownership so the core can scale without a single Workbook translation unit becoming the dependency center of the project.

## What changed

### Domain source tree

P0Z moves implementation code into explicit domains:

- workbook/worksheet model -> `Model/`;
- structural/topology editing -> `Dependencies/`;
- ZIP, OPC and XML -> `Package/`;
- atomic load/save facade -> `IO/`;
- OOXML codecs -> `OOXML/` subdomains;
- VBA -> `VBA/`;
- workbook validation -> `Validation/`;
- chart-cache services -> `Charts/`;
- retained-part policy -> `Preservation/`;
- platform and threading primitives -> `Platform/` and `Core/`.

The old relative `../..` implementation include style was replaced with source-root-relative internal includes.

### Workbook facade split

`Workbook::load()` / `Workbook::save()` path and stream entry points now live in `IO/WorkbookIO.cpp`. Atomic same-directory replacement/temp-file helpers live in `IO/FileTransaction.*`. The implementation continues to delegate OOXML package construction/parsing to internal codec code, preserving P0Y strong-load and atomic-save guarantees.

### OOXML decomposition

The former Workbook codec no longer owns every generated object serializer. P0Z extracts:

- `OOXML/Styles/StyleCodec.*`;
- `OOXML/Common/RichTextCodec.*`;
- `OOXML/Common/PackageRelationships.*`;
- `OOXML/Worksheet/WorksheetFeatureCodec.*`;
- `OOXML/Worksheet/WorksheetWriter.*`;
- `OOXML/Worksheet/WorksheetBatchWriter.*`;
- `OOXML/Pivot/PivotCodec.*`;
- `OOXML/Tables/TableWriter.*`;
- `OOXML/Drawings/DrawingWriter.*`;
- `OOXML/Comments/CommentCodec.*`.

`OOXML/Workbook/WorkbookCodec.cpp` remains an integration codec for the preservation-aware Chart/Drawing paths that still share significant internal state. It is smaller than the P0Y `Workbook.cpp`, but further decomposition remains intentional future debt rather than being forced through cross-module globals.

### Internal CMake modules

CMake now compiles separate object libraries for Model, Formula, Dependencies, Package, IO, Preservation, OOXML, Charts, Streaming, Encryption, VBA and Validation, then combines those objects into the existing `xlpp_static` target. This improves incremental-build ownership and makes architectural coupling visible without changing consumer linkage.

### Test decomposition

The ~9.7k-line `tests/XLPP.UnitTests/main.cpp` was removed. Test registration/execution is separated from model/formula/regression test translation units under `tests/unit/`. All existing suites/checks remain registered.

### Visual Studio project

`src/XLPP/XLPP.vcxproj` is regenerated against the new physical source tree and retains Platform Toolset v145. The internal source root is an include directory so the same root-relative internal headers used by CMake are valid under MSBuild.

## Architectural invariants

1. `Package/` operates on bytes, streams, XML, OPC parts and relationships; it must not depend on Workbook semantic objects.
2. Public headers remain under `include/XLPP/`; `src/XLPP/` headers are private implementation details.
3. OOXML adapters translate between semantic models and package XML rather than putting XML serialization in model classes.
4. Preservation behavior is retained during refactor; selective imported-object editing must not become full-object regeneration merely to simplify source layout.
5. The public deliverable remains one XL++ library.
6. Refactor commits must keep the native regression baseline green before the next extraction is attempted.

## Scope

P0Z does not claim additional Excel feature parity. Formula, Encryption, Pivot, structural editing and compatibility coverage are inherited from P0Y/P0X. The milestone is intended to reduce architectural/build debt and make later Excel-compatibility work safer.

## Remaining architecture debt

> **P0Z-G/v1.10.0 supersession note:** the `WorkbookCodec.cpp` and regression-monolith bullets below are historical and were closed by P0Z-B/P0Z-C. P0Z-G also adds dependency guards for Formula/Dependencies/Validation -> serialization layers. Python/C# binary verification remains a CI/toolchain gate.


- preservation-aware imported Chart/Drawing parsing/patching is still concentrated in `WorkbookCodec.cpp`;
- `tests/unit/Regression/RegressionTests.cpp` is still large and should be split into Package, Charts, Pivot, Streaming and compatibility translation units;
- the object-library layout currently exposes build boundaries, but stricter target-level dependency enforcement can be added later;
- Python/C# binding build verification is separate from this native-core refactor.
