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
    xlpp_workbook_destroy(wb2);
    std::remove(path);

    fprintf(stderr, "XLPP C API smoke test OK\n");
    return 0;
}
