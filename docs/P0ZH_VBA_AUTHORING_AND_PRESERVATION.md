# P0Z-H — VBA Authoring + Preservation / v1.11.0

P0Z-H expands XL++ from basic Standard-module generation into a source-oriented VBA project subsystem while keeping opaque Excel-created projects non-destructive.

## Scope delivered

### Module model

`VbaModule` now distinguishes Standard, Document and Class modules and carries the OVBA module flags `readOnly` and `privateModule`.

XL++ can author:

- Standard modules;
- Class modules;
- `ThisWorkbook` document source;
- worksheet document source, including event procedures.

The reader uses the compressed `dir` stream for procedural-vs-object shape and the textual `PROJECT` stream to refine object modules into Document versus Class semantics. Module stream names are read independently from display names.

### Stable worksheet code names

Each `Worksheet` owns a stable `vbaCodeName`. `sheetPr/@codeName` is loaded and written, and new/copied sheets receive unique host code names. This prevents worksheet event code from being retargeted simply because sheets were reordered or a preceding sheet was deleted.

`copyWorksheet()` also copies authored document-event source under the clone's fresh code name.

### Project properties

`VbaProjectProperties` exposes:

- project name;
- description;
- Help file;
- Help context ID;
- conditional compilation constants.

The builder writes these to the VBA `PROJECT` and compressed `dir` project-information records. The reader restores them from existing projects, and later module/topology rebuilds preserve them rather than resetting metadata to defaults.

The source-only builder is currently MBCS-oriented; broad Unicode/code-page authoring is still a compatibility item.

### Binary and signature safety

XL++ exposes the raw `vbaProject.bin` bytes and can export them to a file. It also reports whether VBA signature parts are present and whether the current project is safe for XL++ source editing.

Two ownership modes are deliberate:

1. **XL++ generated source project** — source and metadata can be edited and rebuilt.
2. **Externally supplied/Excel-owned project** — bytes are preserved, source/metadata rewrite is rejected because rebuilding could discard references, designer streams, password/lock state, p-code or signatures.

A project with VBA signature parts is never advertised as source-editable. Source/module/metadata mutations are rejected instead of silently invalidating the signature.

XL++ does not generate or re-sign digital signatures in this milestone.

## Binding surface

The additions are bridged through the additive C ABI and source-level Python/C# bindings:

- Standard/Class/Document module authoring;
- module type/read-only/private metadata;
- worksheet VBA code name;
- project Name/Description/Help/Help Context/constants;
- raw project bytes/export;
- signature and source-editability state.

No existing C ABI structure was enlarged for this milestone.

## Verification

The VBA regression now exercises generated `.xlsm` save/load, class/document semantics, workbook and worksheet events, module flags, stable host topology, project metadata and conditional constants, raw binary export, external-project safety and signed-project mutation rejection.

The release gate also includes C ABI smoke coverage for the new VBA surface plus binding parity/manifest checks.

## Explicitly remaining

P0Z-H does **not** claim complete VBA/VBE parity. Remaining major work includes:

- UserForms, designer storages and FRX resources;
- ActiveX/Form controls;
- arbitrary VBA/type-library reference authoring;
- project password/locking authoring;
- digital signing/re-signing;
- p-code/stomping analysis;
- full Unicode/code-page authoring;
- `.xltm` and broad Excel Desktop/Trust Center execution validation.

These are kept separate because they require designer, COM/type-library, cryptographic or host-runtime subsystems rather than just OVBA source serialization.
