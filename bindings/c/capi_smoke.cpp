#include "xlpp_capi.h"
#include <cstring>
#include <cstdio>
#include <cmath>

// CI-friendly checks: return 1 (nonzero) on the first failure instead of
// aborting, so the failure is visible in the workflow log.
#define CHECK(cond, msg) do { if (!(cond)) { fprintf(stderr, "FAIL: %s\n", msg); return 1; } } while (0)
#define CHECK_STR(a, b, msg) do { if (std::strcmp((a), (b)) != 0) { fprintf(stderr, "FAIL: %s (got '%s')\n", msg, (a)); return 1; } } while (0)

int main() {
    // Error-handling API.
    xlpp_clear_error();
    CHECK_STR(xlpp_last_error(), "", "clear_error");

    // Null / empty-name rejection.
    auto workbook = xlpp_workbook_create();
    CHECK(workbook != nullptr, "workbook_create");
    CHECK(xlpp_workbook_add_sheet(workbook, nullptr) == nullptr, "null_name_rejected");

    // Worksheet creation.
    xlpp_clear_error();
    auto sheet = xlpp_workbook_add_sheet(workbook, "Smoke");
    CHECK(sheet != nullptr, "add_sheet");
    CHECK_STR(xlpp_sheet_name(sheet), "Smoke", "sheet_name");

    // Cell string value.
    auto cell = xlpp_sheet_cell(sheet, "A1");
    CHECK(cell != nullptr, "cell");
    CHECK_STR(xlpp_cell_address(cell), "A1", "cell_address");
    xlpp_cell_set_string(cell, "hello");
    CHECK_STR(xlpp_cell_get_string(cell), "hello", "set_string");
    CHECK(xlpp_cell_value_type(cell) == XLPP_VALUE_STRING, "type_string");

    // Cell number value.
    auto numCell = xlpp_sheet_cell(sheet, "D1");
    xlpp_cell_set_number(numCell, 42);
    CHECK(std::fabs(xlpp_cell_get_number(numCell) - 42.0) < 1e-9, "set_number");
    CHECK(xlpp_cell_value_type(numCell) == XLPP_VALUE_NUMBER, "type_number");

    // Cell bool value.
    auto boolCell = xlpp_sheet_cell(sheet, "E1");
    xlpp_cell_set_bool(boolCell, 1);
    CHECK(xlpp_cell_get_bool(boolCell) == 1, "set_bool");

    // Cell error value.
    auto errCell = xlpp_sheet_cell(sheet, "B1");
    xlpp_cell_set_error(errCell, XLPP_ERROR_DIV0);
    CHECK(xlpp_cell_is_error(errCell) == 1, "is_error");
    CHECK(xlpp_cell_error_code(errCell) == XLPP_ERROR_DIV0, "error_code");

    // Formula.
    auto fCell = xlpp_sheet_cell(sheet, "C1");
    xlpp_cell_set_formula(fCell, "=SUM(A1:A10)");
    CHECK(xlpp_cell_has_formula(fCell) == 1, "has_formula");
    CHECK_STR(xlpp_cell_get_formula(fCell), "=SUM(A1:A10)", "formula_value");

    // Styles.
    auto font = xlpp_cell_font(cell);
    xlpp_font_set_bold(font, 1);
    CHECK(xlpp_font_get_bold(font) == 1, "font_bold");
    xlpp_font_set_size(font, 14);
    CHECK(std::fabs(xlpp_font_get_size(font) - 14.0) < 1e-9, "font_size");
    auto fill = xlpp_cell_fill(cell);
    xlpp_fill_set_pattern(fill, "solid");
    xlpp_fill_set_fg_color(fill, "FFFF0000");
    auto border = xlpp_cell_border(cell);
    xlpp_borderside_set_style(xlpp_border_left(border), "thin");
    auto align = xlpp_cell_alignment(cell);
    xlpp_alignment_set_horizontal(align, "center");

    // Worksheet features.
    xlpp_sheet_merge_cells(sheet, "A1:B1");
    CHECK(xlpp_sheet_is_merged(sheet, "A1") == 1, "is_merged");
    xlpp_sheet_freeze_panes(sheet, "A2");

    // Workbook-safe sheet rename/remove. These C entry points must never
    // allow a C++ exception to cross the C ABI boundary.
    auto tempSheet = xlpp_workbook_add_sheet(workbook, "Temp");
    CHECK(tempSheet != nullptr, "add_temp_sheet");
    CHECK(xlpp_workbook_rename_sheet(workbook, "temp", "Renamed") == 1, "rename_sheet_case_insensitive_lookup");
    CHECK(xlpp_workbook_sheet_by_name(workbook, "RENAMED") != nullptr, "rename_sheet_case_insensitive_result");
    CHECK(xlpp_workbook_remove_sheet(workbook, "renamed") == 1, "remove_sheet_case_insensitive_lookup");
    CHECK(xlpp_workbook_sheet_count(workbook) == 1, "remove_sheet_count");
    xlpp_clear_error();
    CHECK(xlpp_workbook_remove_sheet(workbook, "Smoke") == 0, "remove_last_sheet_rejected");
    CHECK(std::strlen(xlpp_last_error()) != 0, "remove_last_sheet_sets_error");

    // P1T mixed workbook tabs + chartsheet/combo-chart C ABI. Legacy sheet
    // count remains worksheet-only while the tab model includes chart sheets.
    auto chartSheet = xlpp_workbook_add_chartsheet(workbook, "Dashboard", 0);
    CHECK(chartSheet != nullptr, "add_chartsheet");
    CHECK(xlpp_workbook_sheet_count(workbook) == 1, "legacy_sheet_count_excludes_chartsheet");
    CHECK(xlpp_workbook_chartsheet_count(workbook) == 1, "chartsheet_count");
    CHECK(xlpp_workbook_tab_count(workbook) == 2, "mixed_tab_count");
    CHECK(xlpp_workbook_tab_kind(workbook, 1) == 1, "mixed_tab_kind_chartsheet");
    const unsigned char printerSettings[] = {0x10, 0x00, 0x20, 0x30, 0x40};
    CHECK(xlpp_chartsheet_set_printer_settings(chartSheet, printerSettings, sizeof(printerSettings)) == 1,
          "chartsheet_set_printer_settings");
    CHECK(xlpp_chartsheet_printer_settings_size(chartSheet) == sizeof(printerSettings),
          "chartsheet_printer_settings_size");
    auto chart = xlpp_chartsheet_chart(chartSheet);
    CHECK(chart != nullptr, "chartsheet_chart");
    xlpp_chart_set_title(chart, "C API Combo");
    CHECK(xlpp_chart_add_plot(chart, 0, 3, 0) == 0, "chart_add_primary_plot");
    auto primarySeries = xlpp_chart_add_series_to_plot(chart, 0, "Primary");
    CHECK(primarySeries != nullptr, "chart_add_primary_series");
    xlpp_chartseries_set_categories_reference(primarySeries, "'Smoke'!$D$1:$D$1");
    xlpp_chartseries_set_values_reference(primarySeries, "'Smoke'!$D$1:$D$1");
    CHECK(xlpp_chart_add_plot(chart, 1, 0, 1) == 1, "chart_add_secondary_plot");
    auto secondarySeries = xlpp_chart_add_series_to_plot(chart, 1, "Secondary");
    CHECK(secondarySeries != nullptr, "chart_add_secondary_series");
    xlpp_chartseries_set_categories_reference(secondarySeries, "'Smoke'!$D$1:$D$1");
    xlpp_chartseries_set_values_reference(secondarySeries, "'Smoke'!$D$1:$D$1");
    CHECK(xlpp_chart_plot_count(chart) == 2, "chart_plot_count");
    CHECK(xlpp_chart_plot_uses_secondary_axes(chart, 1) == 1, "chart_secondary_plot");
    CHECK(xlpp_workbook_move_tab(workbook, 1, 0) == 1, "move_chartsheet_tab");
    CHECK(xlpp_workbook_set_active_tab(workbook, 0) == 1, "set_active_chartsheet_tab");
    CHECK(xlpp_workbook_active_tab(workbook) == 0, "active_chartsheet_tab");
    CHECK(xlpp_workbook_set_tab_visibility(workbook, 1, 1) == 1, "set_hidden_worksheet_tab");
    CHECK(xlpp_workbook_tab_visibility(workbook, 1) == 1, "hidden_worksheet_tab");
    CHECK(xlpp_workbook_set_tab_visibility(workbook, 1, 0) == 1, "restore_visible_worksheet_tab");
    xlpp_workbook_set_template(workbook, 1);
    CHECK(xlpp_workbook_is_template(workbook) == 1, "set_template_identity");
    xlpp_workbook_set_template(workbook, 0);
    CHECK(xlpp_workbook_is_template(workbook) == 0, "clear_template_identity");
    char tabName[64]{};
    CHECK_STR(xlpp_workbook_tab_name(workbook, 0, tabName, sizeof(tabName)), "Dashboard", "moved_tab_name");

    // Save / load round-trip.
    const char* path = "smoke-roundtrip.xlsx";
    CHECK(xlpp_workbook_save(workbook, path) == 1, "save");
    xlpp_workbook_destroy(workbook);

    auto wb2 = xlpp_workbook_create();
    CHECK(wb2 != nullptr, "workbook_create2");
    CHECK(xlpp_workbook_load(wb2, path) == 1, "load");
    auto s2 = xlpp_workbook_sheet_by_name(wb2, "Smoke");
    CHECK(s2 != nullptr, "sheet_after_load");
    auto c2 = xlpp_sheet_cell(s2, "A1");
    CHECK(c2 != nullptr, "cell_after_load");
    CHECK_STR(xlpp_cell_get_string(c2), "hello", "roundtrip_string");
    auto e2 = xlpp_sheet_cell(s2, "B1");
    CHECK(xlpp_cell_error_code(e2) == XLPP_ERROR_DIV0, "roundtrip_error");
    auto f2 = xlpp_sheet_cell(s2, "C1");
    CHECK_STR(xlpp_cell_get_formula(f2), "=SUM(A1:A10)", "roundtrip_formula");
    CHECK(xlpp_workbook_chartsheet_count(wb2) == 1, "roundtrip_chartsheet_count");
    CHECK(xlpp_workbook_tab_count(wb2) == 2 && xlpp_workbook_tab_kind(wb2, 0) == 1, "roundtrip_mixed_tab_order");
    auto loadedChartSheet = xlpp_workbook_chartsheet_by_name(wb2, "Dashboard");
    CHECK(loadedChartSheet != nullptr, "roundtrip_chartsheet_lookup");
    CHECK(xlpp_chartsheet_printer_settings_size(loadedChartSheet) == sizeof(printerSettings),
          "roundtrip_printer_settings_size");
    unsigned char printerSettingsCopy[sizeof(printerSettings)]{};
    CHECK(xlpp_chartsheet_copy_printer_settings(loadedChartSheet, printerSettingsCopy, sizeof(printerSettingsCopy)) == sizeof(printerSettings),
          "roundtrip_copy_printer_settings");
    CHECK(std::memcmp(printerSettingsCopy, printerSettings, sizeof(printerSettings)) == 0,
          "roundtrip_printer_settings_bytes");
    auto loadedChart = xlpp_chartsheet_chart(loadedChartSheet);
    CHECK(loadedChart != nullptr && xlpp_chart_plot_count(loadedChart) == 2, "roundtrip_chartsheet_combo_chart");
    const char* encryptedPath = "xlpp_capi_encrypted_smoke.xlsx";
    CHECK(xlpp_workbook_save_password(wb2, encryptedPath, "CapiPass!", 1000) == 1, "save_password");
    CHECK(xlpp_workbook_is_password_encrypted_file(encryptedPath) == 1, "detect_password_encryption");
    auto wb3 = xlpp_workbook_create();
    CHECK(wb3 != nullptr, "create_encrypted_reload");
    CHECK(xlpp_workbook_load_password(wb3, encryptedPath, "CapiPass!") == 1, "load_password");
    CHECK(xlpp_workbook_load_password(wb3, encryptedPath, "wrong") == 0, "wrong_password_rejected");
    int encFormat = -1; unsigned encBits = 0; int encHash = -1; uint64_t encSpins = 0; int encIntegrity = 0;
    CHECK(xlpp_workbook_encryption_profile(encryptedPath, &encFormat, &encBits, &encHash, &encSpins, &encIntegrity) == 1, "encryption_profile_agile");
    CHECK(encFormat == XLPP_ENCRYPTION_FORMAT_AGILE && encBits == 256 && encHash == XLPP_ENCRYPTION_HASH_SHA512 && encSpins == 1000 && encIntegrity == 1, "agile_profile_values");
    uint64_t totalKeys = 0, passwordKeys = 0, certificateKeys = 0;
    CHECK(xlpp_workbook_encryption_key_encryptor_counts(encryptedPath, &totalKeys, &passwordKeys, &certificateKeys) == 1, "encryption_key_encryptor_counts");
    CHECK(totalKeys == 1 && passwordKeys == 1 && certificateKeys == 0, "agile_key_encryptor_count_values");
    CHECK(xlpp_workbook_load_password_ex(wb3, encryptedPath, "CapiPass!", 2000, 1024 * 1024, 1, 1, 1024 * 1024) == 1, "load_password_ex_agile_policy");

    const char* standardPath = "xlpp_capi_standard_smoke.xlsx";
    CHECK(xlpp_workbook_save_password_ex(wb2, standardPath, "Compat!", XLPP_ENCRYPTION_STANDARD, 128, XLPP_ENCRYPTION_HASH_SHA512, 7) == 1, "save_password_ex_standard");
    encFormat = -1; encBits = 0; encHash = -1; encSpins = 0; encIntegrity = 1;
    CHECK(xlpp_workbook_encryption_profile(standardPath, &encFormat, &encBits, &encHash, &encSpins, &encIntegrity) == 1, "encryption_profile_standard");
    CHECK(encFormat == XLPP_ENCRYPTION_FORMAT_STANDARD && encBits == 128 && encHash == XLPP_ENCRYPTION_HASH_SHA1 && encSpins == 50000 && encIntegrity == 0, "standard_profile_values");
    CHECK(xlpp_workbook_load_password(wb3, standardPath, "Compat!") == 1, "load_standard_password");
    CHECK(xlpp_workbook_load_password_ex(wb3, standardPath, "Compat!", 1000000, 1024 * 1024, 0, 0, 1024 * 1024) == 0, "load_password_ex_blocks_standard_by_policy");
    std::remove(standardPath);
    xlpp_workbook_destroy(wb3);
    std::remove(encryptedPath);
    xlpp_workbook_destroy(wb2);
    std::remove(path);

    fprintf(stderr, "XLPP C API smoke test OK\n");
    return 0;
}
