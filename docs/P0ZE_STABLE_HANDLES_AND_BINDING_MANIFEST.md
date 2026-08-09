# P0Z-E — Stable Handles + Generated Binding Manifest / v1.8.0

P0Z-E keeps the P0X/P0Y **90.7/100 scope-defined editing-core feature baseline** unchanged and improves native/binding lifetime guarantees, incremental compilation, and binding governance.

## Stable child identity

Bindings frequently retain handles to model children while the parent collection continues to grow. Vector relocation made those handles unsafe for several public collections. P0Z-E introduces `XLPP/Core/StableVector.h`, an owning vector-like sequence whose element addresses remain stable across capacity growth and erasure of other elements.

The stable storage is used for handle-exposed collections including workbook named styles/defined names/custom properties; worksheet tables/images/charts/pivots; table columns; chart series; conditional-formatting entries/rules; data validations; rich-text runs.

The guarantee is intentionally precise: a handle remains valid while its element still exists. Erasing the element itself, replacing/copy-assigning the owning model, or destroying the owner invalidates that element's handle.

Native regression intentionally retains child references and grows owning collections by 128–256 elements. C ABI smoke repeats the same pattern through raw handles. Python parity tests repeat it through pybind live objects.

## Formula build decomposition

Formula calculation is now split into independent translation/build units:

- calculation engine state;
- parser/evaluation context;
- function dispatch/support;
- Math/Statistical/Financial functions;
- Logical/Text/Date functions;
- Criteria/Lookup functions;
- Dynamic-array/Reference functions;
- dependency graph;
- reference translator.

The native core aggregates **35 private object modules** into the unchanged `xlpp_static` / `XLPP::xlpp` target. Adding a formula function no longer requires recompiling the full parser/calculation engine or unrelated function families.

## Generated binding manifest

`tools/generate_binding_manifest.py` derives the public `Workbook` and `Worksheet` method surface from native headers and records binding candidates/coverage in `bindings/PARITY_MANIFEST.json`.

`BindingManifestTests` regenerates the manifest during CTest and fails on any drift. A new native high-level method therefore becomes an explicit binding-review event instead of silently landing only in C++.

The existing `BindingParityTests` remains for invariants that cannot be inferred from names alone: package version synchronization, scoped-defined-name semantics, and complete C# P/Invoke declaration/native-header consistency.

## Binding status

- C ABI stable-handle smoke: verified.
- Python CPython 3.13 extension: binary-built from the same native library and **136/136 tests pass**.
- C# wrapper contains matching stable-handle regression in its executable binding test; the release host still has no .NET SDK/compiler, so local C# binary verification is not claimed. The .NET CI binding job remains the binary gate.

## Compatibility

No Excel feature-score increase is claimed by v1.8.0. This milestone changes storage/architecture and binding governance while preserving the public semantic model and XLSX behavior.
