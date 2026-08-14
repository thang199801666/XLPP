# Binding parity policy — v1.12.0

XLPP maintains four public access layers: native C++, C ABI, Python, and C#. High-level Workbook/Worksheet capabilities are release-blocking parity items; low-level package/preservation internals may remain native-only when exposing them would create unsafe or unstable abstractions.

| Capability family | C++ | C ABI | Python | C# |
|---|:---:|:---:|:---:|:---:|
| Path load/save + options | ✅ | ✅ | ✅ | ✅ |
| Memory/bytes load-save | ✅ | ✅ | ✅ | ✅ |
| Load cancel/progress | ✅ | ✅ | ✅ | ✅ |
| Formula calculation options | ✅ | ✅ | ✅ | ✅ |
| External workbook resolver | ✅ | ✅ | ✅ | ✅ |
| Structural transactions | ✅ | ✅ | ✅ | ✅ |
| Dependency-aware sheet rename | ✅ | ✅ | ✅ | ✅ |
| Formula dependency graph | ✅ | ✅ | ✅ | ✅ |
| Chart-cache synchronization/tracking | ✅ | ✅ | ✅ | ✅ |
| Workbook validation/diagnostics | ✅ | ✅ | ✅ | ✅ |
| VBA Standard/Class/Document modules, stable code names, project properties/raw export/signature state | ✅ | ✅ | ✅ | ✅ |
| Scoped defined names | ✅ | ✅ | ✅ | ✅ |
| Agile/Standard password encryption workflow | ✅ | ✅ | ✅ | ✅ |
| Stable surviving child handles across collection growth | ✅ | ✅ | ✅ | ✅ |

Verification rules:

- C# declares every exported C ABI function. The current bridge covers all 869 exports, including rich text, formula metadata/reference translation, drawing anchors, streaming reader options, nested Chart line/fill/rich-text/cache/layout/3D/trendline/error-bar editing, encryption inspection, and enterprise feature inspection/editing.

- C ABI must compile as C and pass `xlpp_capi_smoke`.
- Python must binary-build and run `bindings/python/tests` in binding CI/release verification.
- C# must build and run its managed executable/test project in the .NET binding CI job.
- `BindingManifestTests` regenerates the high-level native `Workbook`/`Worksheet` surface from public headers and compares it with `bindings/PARITY_MANIFEST.json`.
- `BindingParityTests` enforces non-name-inferable bridge invariants, C# P/Invoke consistency and synchronized package versions.
- New high-level native APIs must review/regenerate the manifest and update bindings or explicitly document why the surface is native-only.

## Child-handle lifetime

Handle-exposed collections now use stable native element storage. A child handle remains valid when its owning collection grows or when a different element is erased. A handle is still invalid after its own element is erased, after its owner is destroyed/replaced, or after an operation explicitly documented to replace the containing model. Bindings must not dereference handles after those invalidating operations.
