# P1W-A Template / Chartsheet Auxiliary Package Ownership Hardening

P1W-A continues the `.xltx` / `.xltm` and Chartsheet production track. The milestone focuses on package ownership that sits beside the normal Chartsheet chart/drawing closure: printer settings and header/footer VML drawings.

## First-class Chartsheet printer settings payload

`Chartsheet` can now retain opaque printer-settings bytes associated with top-level `pageSetup@r:id`:

```cpp
xlpp::Chartsheet& cs = wb.chartsheet("Dashboard");
const std::string devModeBytes{"...binary..."};
cs.setPrinterSettingsData(devModeBytes);

if (cs.hasPrinterSettings()) {
    const auto& bytes = *cs.printerSettingsData();
}

cs.clearPrinterSettings();
```

The payload is intentionally opaque. XL++ owns and preserves the binary part and relationship but does not attempt to decode platform-specific DEVMODE structures.

On save, generated/replaced settings are emitted under `xl/printerSettings/printerSettingsN.bin` with the SpreadsheetML printer-settings content type. Imported `pageSetup` relationships are resolved and validated during load. Replacing or clearing printer settings retires only the old printer-settings closure, not unrelated Chartsheet relationships.

## Header/footer VML ownership preservation

P1W fixes a prior preservation defect in metadata-only Chartsheet serialization: the actual owner node is `legacyDrawingHF`, not `drawingHF`.

Imported Chartsheets that own header/footer pictures through a VML drawing now retain:

- the `legacyDrawingHF` owner node;
- the Chartsheet `vmlDrawing` relationship;
- the VML part itself;
- unrelated auxiliary relationships when the chart/drawing subtree is regenerated.

Chart regeneration now retires only the old chart drawing closure. It no longer suppresses every child relationship of the imported Chartsheet.

## Package graph validation

`RelationshipGraph` / `xlpp-package-validator` now understand Chartsheet auxiliary ownership:

- `legacyDrawingHF@r:id` must resolve to a VML drawing relationship;
- `pageSetup@r:id` must resolve to a printer-settings relationship;
- orphan/dangling auxiliary parts are surfaced as package validation errors;
- package inventory reports header/footer drawings and printer-settings parts separately.

## C ABI

P1W adds opaque binary printer-settings access to the C API:

```c
xlpp_chartsheet_set_printer_settings(cs, data, size);
uint64_t needed = xlpp_chartsheet_printer_settings_size(cs);
xlpp_chartsheet_copy_printer_settings(cs, out, capacity);
xlpp_chartsheet_clear_printer_settings(cs);
```

The API is length-based and supports embedded NUL bytes.

## Verification

Final P1W-A verification on 2026-08-14:

- Main regression: **198/198 suites PASS, 3,632/3,632 checks PASS**.
- P1P / P1Q / P1R / P1S / P1T / P1U / P1V / P1W dedicated suites: PASS.
- `xlpp_capi_smoke`: PASS, including binary printer-settings save/load/copy.
- Standalone public-header target: PASS.
- Installed `find_package(XLPP CONFIG)` consumer: PASS.
- Source `add_subdirectory()` consumer: PASS.
- Focused Clang 17 ASan + UBSan + leak detection on `XLPP_P1W_ChartsheetPackageTests`: PASS.
- Independent openpyxl 3.1.5 host check accepts the generated template/Chartsheet artifacts and sees the auxiliary package ownership intact.
- `xlpp-package-validator` reports zero graph/content-type/owner-reference errors for generated printer-settings and imported header/footer-drawing artifacts.

## Architectural note

P1W deliberately avoids a risky last-minute package-writer refactor. The strengthened ownership merge logic currently leaves `Workbook.cpp` at roughly **579 KiB / 9,547 lines**. The next milestone should extract relationship merge/closure and Chartsheet package-writer code into dedicated translation units before adding another large Chartsheet feature family.

## ABI note

`Chartsheet` gains new state for printer-settings ownership. Native C++ consumers should rebuild together with the P1W core. The C API additions are source-additive when the wrapper is rebuilt.
