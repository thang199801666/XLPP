// XLPP C API Implementation
#include "xlpp_capi.h"
#include <XLPP/XLPP.h>
#include <string>
#include <vector>
#include <cstring>
#include <exception>

namespace {
thread_local std::string g_lastError;

void clearError() noexcept { g_lastError.clear(); }
void setError(const char* message) noexcept { g_lastError = message ? message : "XLPP C API error"; }
void setError(const std::exception& error) noexcept { g_lastError = error.what(); }
}

// Simple handle wrapping: cast between opaque pointer and C++ pointer
#define WB(h)   reinterpret_cast<xlpp::Workbook*>(h)
#define WS(h)   reinterpret_cast<xlpp::Worksheet*>(h)
#define CELL(h) reinterpret_cast<xlpp::Cell*>(h)
#define FONT(h) reinterpret_cast<xlpp::Font*>(h)
#define FILL(h) reinterpret_cast<xlpp::Fill*>(h)
#define BDR(h)  reinterpret_cast<xlpp::Border*>(h)
#define BS(h)   reinterpret_cast<xlpp::BorderSide*>(h)
#define ALN(h)  reinterpret_cast<xlpp::Alignment*>(h)
#define STY(h)  reinterpret_cast<xlpp::Style*>(h)
#define PROP(h) reinterpret_cast<xlpp::DocumentProperties*>(h)

// ============================================================
// Workbook
// ============================================================
extern "C" {

XLPP_API const char* xlpp_version(void) {
    return "1.1.1";
}

XLPP_API const char* xlpp_last_error(void) {
    return g_lastError.c_str();
}

XLPP_API void xlpp_clear_error(void) {
    clearError();
}

XLPP_API xlpp_workbook xlpp_workbook_create(void) {
    try {
        clearError();
        return reinterpret_cast<xlpp_workbook>(new xlpp::Workbook());
    } catch (const std::exception& error) {
        setError(error);
    } catch (...) {
        setError("Failed to create workbook");
    }
    return nullptr;
}

XLPP_API void xlpp_workbook_destroy(xlpp_workbook wb) {
    delete WB(wb);
}

XLPP_API xlpp_worksheet xlpp_workbook_add_sheet(xlpp_workbook wb, const char* name) {
    // addWorksheet throws on empty/duplicate names; never let a C++ exception
    // cross the C ABI.
    try {
        clearError();
        if (!wb || !name) { setError("Workbook and sheet name are required"); return nullptr; }
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->addWorksheet(name));
    } catch (const std::exception& error) { setError(error); return nullptr; }
    catch (...) { setError("Failed to add worksheet"); return nullptr; }
}

XLPP_API int xlpp_workbook_sheet_count(xlpp_workbook wb) {
    if (!wb) { setError("Workbook handle is null"); return 0; }
    return static_cast<int>(WB(wb)->sheetCount());
}

XLPP_API xlpp_worksheet xlpp_workbook_get_sheet(xlpp_workbook wb, int index) {
    try {
        clearError();
        if (!wb || index < 0) { setError("Workbook handle or sheet index is invalid"); return nullptr; }
        return reinterpret_cast<xlpp_worksheet>(&(*WB(wb))[static_cast<std::size_t>(index)]);
    } catch (const std::exception& error) { setError(error); return nullptr; }
    catch (...) { setError("Failed to get worksheet"); return nullptr; }
}

XLPP_API xlpp_worksheet xlpp_workbook_sheet_by_name(xlpp_workbook wb, const char* name) {
    if (!wb || !name) { setError("Workbook and sheet name are required"); return nullptr; }
    auto* ws = WB(wb)->worksheet(name);
    return reinterpret_cast<xlpp_worksheet>(ws);
}

XLPP_API int xlpp_workbook_remove_sheet(xlpp_workbook wb, const char* name) {
    return WB(wb)->removeWorksheet(name) ? 1 : 0;
}

XLPP_API int xlpp_workbook_load(xlpp_workbook wb, const char* path) {
    try {
        clearError();
        if (!wb || !path) { setError("Workbook and path are required"); return 0; }
        WB(wb)->load(std::filesystem::path(path));
        return 1;
    } catch (const std::exception& error) { setError(error); return 0; }
    catch (...) { setError("Failed to load workbook"); return 0; }
}

XLPP_API int xlpp_workbook_save(xlpp_workbook wb, const char* path) {
    try {
        clearError();
        if (!wb || !path) { setError("Workbook and path are required"); return 0; }
        WB(wb)->save(std::filesystem::path(path));
        return 1;
    } catch (const std::exception& error) { setError(error); return 0; }
    catch (...) { setError("Failed to save workbook"); return 0; }
}

XLPP_API xlpp_properties xlpp_workbook_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_properties>(&WB(wb)->properties());
}

// ============================================================
// Properties
// ============================================================
XLPP_API void xlpp_properties_set_title(xlpp_properties p, const char* v)    { PROP(p)->setTitle(v); }
XLPP_API void xlpp_properties_set_creator(xlpp_properties p, const char* v)   { PROP(p)->setCreator(v); }
XLPP_API void xlpp_properties_set_subject(xlpp_properties p, const char* v)   { PROP(p)->setSubject(v); }
XLPP_API const char* xlpp_properties_get_title(xlpp_properties p)             { return PROP(p)->title().c_str(); }
XLPP_API const char* xlpp_properties_get_creator(xlpp_properties p)           { return PROP(p)->creator().c_str(); }

// ============================================================
// Worksheet
// ============================================================
XLPP_API const char* xlpp_sheet_name(xlpp_worksheet ws) { return WS(ws)->name().c_str(); }
XLPP_API void xlpp_sheet_rename(xlpp_worksheet ws, const char* name) { WS(ws)->rename(name); }

XLPP_API xlpp_cell xlpp_sheet_cell(xlpp_worksheet ws, const char* address) {
    try {
        return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(address));
    } catch (...) { return nullptr; }
}

XLPP_API xlpp_cell xlpp_sheet_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col) {
    try {
        return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(static_cast<std::size_t>(row), static_cast<std::size_t>(col)));
    } catch (...) { return nullptr; }
}

XLPP_API int xlpp_sheet_has_cell(xlpp_worksheet ws, const char* address) {
    return WS(ws)->tryCell(address) != nullptr ? 1 : 0;
}

XLPP_API void xlpp_sheet_append_row(xlpp_worksheet ws, const char** values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        if (values[i] && values[i][0])
            cv.push_back(std::string(values[i]));
        else
            cv.push_back(std::monostate{});
    }
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_append_doubles(xlpp_worksheet ws, const double* values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) cv.push_back(values[i]);
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_merge_cells(xlpp_worksheet ws, const char* range)     { try { WS(ws)->mergeCells(range); } catch (...) {} }
XLPP_API void xlpp_sheet_unmerge_cells(xlpp_worksheet ws, const char* range)   { try { WS(ws)->unmergeCells(range); } catch (...) {} }
XLPP_API void xlpp_sheet_freeze_panes(xlpp_worksheet ws, const char* cell)     { try { WS(ws)->freezePanes(cell); } catch (...) {} }

XLPP_API uint64_t xlpp_sheet_max_row(xlpp_worksheet ws) { return WS(ws)->maxRow(); }
XLPP_API uint64_t xlpp_sheet_max_col(xlpp_worksheet ws) { return WS(ws)->maxColumn(); }

XLPP_API void xlpp_sheet_dimensions(xlpp_worksheet ws, char* out, int outSize) {
    auto d = WS(ws)->dimensions();
    std::memcpy(out, d.c_str(), (std::min)(d.size(), static_cast<std::size_t>(outSize - 1)));
    out[(std::min)(d.size(), static_cast<std::size_t>(outSize - 1))] = '\0';
}

XLPP_API void xlpp_sheet_insert_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->insertRows(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_delete_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->deleteRows(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_insert_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->insertColumns(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}
XLPP_API void xlpp_sheet_delete_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount) {
    try { WS(ws)->deleteColumns(static_cast<std::size_t>(index), static_cast<std::size_t>(amount)); } catch (...) {}
}

XLPP_API xlpp_table xlpp_sheet_add_table(xlpp_worksheet ws, const char* name, const char* reference) {
    try {
        if (!ws || !name || !reference) { setError("Worksheet, table name, and reference are required"); return nullptr; }
        return reinterpret_cast<xlpp_table>(&WS(ws)->addTable(name, reference));
    } catch (const std::exception& error) { setError(error); return nullptr; }
    catch (...) { setError("Failed to add table"); return nullptr; }
}

XLPP_API xlpp_table xlpp_sheet_table(xlpp_worksheet ws, const char* name) {
    if (!ws || !name) return nullptr;
    return reinterpret_cast<xlpp_table>(WS(ws)->table(name));
}

// ============================================================
// Cell
// ============================================================
XLPP_API const char* xlpp_cell_address(xlpp_cell c)  { return CELL(c)->address().c_str(); }
XLPP_API uint64_t xlpp_cell_row(xlpp_cell c)          { return CELL(c)->row(); }
XLPP_API uint64_t xlpp_cell_column(xlpp_cell c)       { return CELL(c)->column(); }

XLPP_API int xlpp_cell_value_type(xlpp_cell c) {
    const auto& v = CELL(c)->value();
    if (std::holds_alternative<std::monostate>(v)) return XLPP_VALUE_EMPTY;
    if (std::holds_alternative<bool>(v))           return XLPP_VALUE_BOOL;
    if (std::holds_alternative<double>(v))           return XLPP_VALUE_NUMBER;
    if (std::holds_alternative<std::string>(v))      return XLPP_VALUE_STRING;
    if (std::holds_alternative<xlpp::CellError>(v))  return XLPP_VALUE_ERROR;
    if (std::holds_alternative<xlpp::DateTime>(v))   return XLPP_VALUE_DATE;
    return XLPP_VALUE_EMPTY;
}

XLPP_API int xlpp_cell_get_bool(xlpp_cell c) {
    try {
        return std::get<bool>(CELL(c)->value()) ? 1 : 0;
    } catch (...) { return 0; }
}
XLPP_API double xlpp_cell_get_number(xlpp_cell c) {
    if (auto* v = std::get_if<double>(&CELL(c)->value())) return *v;
    if (auto* v = std::get_if<xlpp::DateTime>(&CELL(c)->value()))
        return xlpp::toExcelSerial(*v, false);
    return 0.0;
}
XLPP_API const char* xlpp_cell_get_string(xlpp_cell c) {
    if (auto* v = std::get_if<std::string>(&CELL(c)->value())) return v->c_str();
    return "";
}
XLPP_API int xlpp_cell_is_empty(xlpp_cell c) { return CELL(c)->empty() ? 1 : 0; }

XLPP_API void xlpp_cell_set_string(xlpp_cell c, const char* v)   { CELL(c)->setStringValue(v ? v : ""); }
XLPP_API void xlpp_cell_set_number(xlpp_cell c, double v)        { CELL(c)->setNumericValue(v); }
XLPP_API void xlpp_cell_set_bool(xlpp_cell c, int v)             { CELL(c)->setBoolValue(v != 0); }
XLPP_API void xlpp_cell_set_empty(xlpp_cell c)                   { CELL(c)->setValue(std::monostate{}); }
XLPP_API void xlpp_cell_clear(xlpp_cell c)                       { CELL(c)->clear(); }

XLPP_API const char* xlpp_cell_get_formula(xlpp_cell c)          { return CELL(c)->formula().c_str(); }
XLPP_API void xlpp_cell_set_formula(xlpp_cell c, const char* f)   { CELL(c)->setFormula(f ? f : ""); }
XLPP_API int xlpp_cell_has_formula(xlpp_cell c)                  { return CELL(c)->hasFormula() ? 1 : 0; }

XLPP_API xlpp_style xlpp_cell_style(xlpp_cell c)         { return reinterpret_cast<xlpp_style>(&CELL(c)->style()); }
XLPP_API xlpp_font xlpp_cell_font(xlpp_cell c)           { return reinterpret_cast<xlpp_font>(&CELL(c)->font()); }
XLPP_API xlpp_fill xlpp_cell_fill(xlpp_cell c)           { return reinterpret_cast<xlpp_fill>(&CELL(c)->fill()); }
XLPP_API xlpp_border xlpp_cell_border(xlpp_cell c)       { return reinterpret_cast<xlpp_border>(&CELL(c)->border()); }
XLPP_API xlpp_alignment xlpp_cell_alignment(xlpp_cell c)  { return reinterpret_cast<xlpp_alignment>(&CELL(c)->alignment()); }

XLPP_API void xlpp_cell_set_hyperlink(xlpp_cell c, const char* url) {
    CELL(c)->setHyperlink(xlpp::Hyperlink(url ? url : ""));
}
XLPP_API void xlpp_cell_set_comment(xlpp_cell c, const char* text, const char* author) {
    CELL(c)->setComment(xlpp::Comment(text ? text : "", author ? author : ""));
}

XLPP_API const char* xlpp_table_name(xlpp_table table) { return table ? reinterpret_cast<xlpp::Table*>(table)->name().c_str() : ""; }
XLPP_API const char* xlpp_table_reference(xlpp_table table) { return table ? reinterpret_cast<xlpp::Table*>(table)->reference().c_str() : ""; }
XLPP_API void xlpp_table_set_reference(xlpp_table table, const char* value) { if (table && value) reinterpret_cast<xlpp::Table*>(table)->setReference(value); }
XLPP_API int xlpp_table_show_header_row(xlpp_table table) { return table && reinterpret_cast<xlpp::Table*>(table)->showHeaderRow() ? 1 : 0; }
XLPP_API void xlpp_table_set_show_header_row(xlpp_table table, int value) { if (table) reinterpret_cast<xlpp::Table*>(table)->setShowHeaderRow(value != 0); }
XLPP_API int xlpp_table_show_totals_row(xlpp_table table) { return table && reinterpret_cast<xlpp::Table*>(table)->showTotalsRow() ? 1 : 0; }
XLPP_API void xlpp_table_set_show_totals_row(xlpp_table table, int value) { if (table) reinterpret_cast<xlpp::Table*>(table)->setShowTotalsRow(value != 0); }
XLPP_API int xlpp_table_column_count(xlpp_table table) { return table ? static_cast<int>(reinterpret_cast<xlpp::Table*>(table)->columns().size()) : 0; }
XLPP_API void xlpp_table_add_column(xlpp_table table, const char* name) { if (table && name) reinterpret_cast<xlpp::Table*>(table)->addColumn(name); }
XLPP_API const char* xlpp_table_display_name(xlpp_table table) { return table ? reinterpret_cast<xlpp::Table*>(table)->displayName().c_str() : ""; }
XLPP_API void xlpp_table_set_display_name(xlpp_table table, const char* value) { if (table && value) reinterpret_cast<xlpp::Table*>(table)->setDisplayName(value); }
XLPP_API const char* xlpp_table_style_name(xlpp_table table) { return table ? reinterpret_cast<xlpp::Table*>(table)->styleInfo().name().c_str() : ""; }
XLPP_API void xlpp_table_set_style_name(xlpp_table table, const char* value) { if (table && value) reinterpret_cast<xlpp::Table*>(table)->styleInfo().setName(value); }
XLPP_API int xlpp_table_show_row_stripes(xlpp_table table) { return table && reinterpret_cast<xlpp::Table*>(table)->styleInfo().showRowStripes() ? 1 : 0; }
XLPP_API void xlpp_table_set_show_row_stripes(xlpp_table table, int value) { if (table) reinterpret_cast<xlpp::Table*>(table)->styleInfo().setShowRowStripes(value != 0); }

XLPP_API xlpp_data_validation xlpp_sheet_add_list_validation(xlpp_worksheet ws, const char* reference, const char* formula) {
    try {
        if (!ws || !reference || !formula) { setError("Worksheet, validation reference, and formula are required"); return nullptr; }
        auto& value = WS(ws)->dataValidations().add(xlpp::DataValidationType::List, reference);
        value.setFormula1(formula);
        return reinterpret_cast<xlpp_data_validation>(&value);
    } catch (const std::exception& error) { setError(error); return nullptr; }
    catch (...) { setError("Failed to add list validation"); return nullptr; }
}

XLPP_API void xlpp_validation_set_allow_blank(xlpp_data_validation validation, int value) { if (validation) reinterpret_cast<xlpp::DataValidation*>(validation)->setAllowBlank(value != 0); }
XLPP_API void xlpp_validation_set_prompt(xlpp_data_validation validation, const char* title, const char* text) { if (validation) { auto* v = reinterpret_cast<xlpp::DataValidation*>(validation); v->setPromptTitle(title ? title : ""); v->setPrompt(text ? text : ""); } }
XLPP_API void xlpp_validation_set_error(xlpp_data_validation validation, const char* title, const char* text) { if (validation) { auto* v = reinterpret_cast<xlpp::DataValidation*>(validation); v->setErrorTitle(title ? title : ""); v->setError(text ? text : ""); } }

XLPP_API xlpp_chart xlpp_chart_create(int type) {
    if (type < 0 || type > 7) { setError("Chart type is invalid"); return nullptr; }
    return reinterpret_cast<xlpp_chart>(new xlpp::Chart(static_cast<xlpp::Chart::Type>(type)));
}
XLPP_API void xlpp_chart_destroy(xlpp_chart chart) { delete reinterpret_cast<xlpp::Chart*>(chart); }
XLPP_API void xlpp_chart_set_title(xlpp_chart chart, const char* title) { if (!chart) { setError("Chart handle is null"); return; } reinterpret_cast<xlpp::Chart*>(chart)->setTitle(title ? title : ""); }
XLPP_API void xlpp_chart_set_x_axis_title(xlpp_chart chart, const char* title) { if (!chart) { setError("Chart handle is null"); return; } reinterpret_cast<xlpp::Chart*>(chart)->setXAxisTitle(title ? title : ""); }
XLPP_API void xlpp_chart_set_y_axis_title(xlpp_chart chart, const char* title) { if (!chart) { setError("Chart handle is null"); return; } reinterpret_cast<xlpp::Chart*>(chart)->setYAxisTitle(title ? title : ""); }
XLPP_API void xlpp_chart_set_style(xlpp_chart chart, const char* style) { if (!chart) { setError("Chart handle is null"); return; } reinterpret_cast<xlpp::Chart*>(chart)->setStyle(style ? style : ""); }
XLPP_API void xlpp_chart_set_grouping(xlpp_chart chart, int grouping) { if (!chart) { setError("Chart handle is null"); return; } if (grouping < 0 || grouping > 3) { setError("Chart grouping is invalid"); return; } reinterpret_cast<xlpp::Chart*>(chart)->setGrouping(static_cast<xlpp::Chart::Grouping>(grouping)); }
XLPP_API void xlpp_chart_set_size(xlpp_chart chart, int width, int height) { if (!chart) { setError("Chart handle is null"); return; } if (width <= 0 || height <= 0) { setError("Chart size must be positive"); return; } auto* c = reinterpret_cast<xlpp::Chart*>(chart); c->setWidth(width); c->setHeight(height); }
XLPP_API void xlpp_chart_set_legend(xlpp_chart chart, int show, const char* position) { if (!chart) { setError("Chart handle is null"); return; } auto* c = reinterpret_cast<xlpp::Chart*>(chart); c->setShowLegend(show != 0); if (position) c->setLegendPosition(position); }
XLPP_API xlpp_chart_series xlpp_chart_add_series(xlpp_chart chart, const char* title) { if (!chart) { setError("Chart handle is null"); return nullptr; } return reinterpret_cast<xlpp_chart_series>(&reinterpret_cast<xlpp::Chart*>(chart)->addSeries(xlpp::ChartSeries(title ? title : ""))); }
XLPP_API void xlpp_chart_series_set_values_reference(xlpp_chart_series series, const char* reference) { if (!series) { setError("Chart series handle is null"); return; } reinterpret_cast<xlpp::ChartSeries*>(series)->setValuesReference(reference ? reference : ""); }
XLPP_API void xlpp_chart_series_set_categories_reference(xlpp_chart_series series, const char* reference) { if (!series) { setError("Chart series handle is null"); return; } reinterpret_cast<xlpp::ChartSeries*>(series)->setCategoriesReference(reference ? reference : ""); }
XLPP_API int xlpp_sheet_add_chart(xlpp_worksheet ws, xlpp_chart chart) {
    if (!chart) return 0;
    auto* value = reinterpret_cast<xlpp::Chart*>(chart);
    if (!ws) { delete value; return 0; }
    try {
        WS(ws)->addChart(std::move(*value));
        delete value;
        return 1;
    } catch (const std::exception& error) {
        setError(error);
        delete value;
        return 0;
    } catch (...) {
        setError("Failed to add chart");
        delete value;
        return 0;
    }
}

// ============================================================
// Font
// ============================================================
XLPP_API void xlpp_font_set_name(xlpp_font f, const char* v)       { FONT(f)->setName(v); }
XLPP_API void xlpp_font_set_size(xlpp_font f, double v)            { FONT(f)->setSize(v); }
XLPP_API void xlpp_font_set_bold(xlpp_font f, int v)               { FONT(f)->setBold(v != 0); }
XLPP_API void xlpp_font_set_italic(xlpp_font f, int v)             { FONT(f)->setItalic(v != 0); }
XLPP_API void xlpp_font_set_underline(xlpp_font f, int v)          { FONT(f)->setUnderline(v != 0); }
XLPP_API void xlpp_font_set_color(xlpp_font f, const char* argb)   { FONT(f)->color().setArgb(argb); }
XLPP_API const char* xlpp_font_get_name(xlpp_font f)               { return FONT(f)->name().c_str(); }
XLPP_API double xlpp_font_get_size(xlpp_font f)                    { return FONT(f)->size(); }
XLPP_API int xlpp_font_get_bold(xlpp_font f)                       { return FONT(f)->bold() ? 1 : 0; }

// ============================================================
// Fill
// ============================================================
XLPP_API void xlpp_fill_set_pattern(xlpp_fill f, const char* v)     { FILL(f)->setPatternType(v); }
XLPP_API void xlpp_fill_set_fg_color(xlpp_fill f, const char* argb) { FILL(f)->foregroundColor().setArgb(argb); }
XLPP_API void xlpp_fill_set_bg_color(xlpp_fill f, const char* argb) { FILL(f)->backgroundColor().setArgb(argb); }

// ============================================================
// Border
// ============================================================
XLPP_API xlpp_borderside xlpp_border_left(xlpp_border b)    { return reinterpret_cast<xlpp_borderside>(&BDR(b)->left()); }
XLPP_API xlpp_borderside xlpp_border_right(xlpp_border b)   { return reinterpret_cast<xlpp_borderside>(&BDR(b)->right()); }
XLPP_API xlpp_borderside xlpp_border_top(xlpp_border b)     { return reinterpret_cast<xlpp_borderside>(&BDR(b)->top()); }
XLPP_API xlpp_borderside xlpp_border_bottom(xlpp_border b)  { return reinterpret_cast<xlpp_borderside>(&BDR(b)->bottom()); }
XLPP_API void xlpp_borderside_set_style(xlpp_borderside s, const char* v) { BS(s)->setStyle(v); }
XLPP_API void xlpp_borderside_set_color(xlpp_borderside s, const char* argb) { BS(s)->color().setArgb(argb); }

// ============================================================
// Alignment
// ============================================================
XLPP_API void xlpp_alignment_set_horizontal(xlpp_alignment a, const char* v) { ALN(a)->setHorizontal(v); }
XLPP_API void xlpp_alignment_set_vertical(xlpp_alignment a, const char* v)   { ALN(a)->setVertical(v); }
XLPP_API void xlpp_alignment_set_wrap_text(xlpp_alignment a, int v)           { ALN(a)->setWrapText(v != 0); }

// ============================================================
// Style
// ============================================================
XLPP_API xlpp_font xlpp_style_font(xlpp_style s)            { return reinterpret_cast<xlpp_font>(&STY(s)->font()); }
XLPP_API xlpp_fill xlpp_style_fill(xlpp_style s)            { return reinterpret_cast<xlpp_fill>(&STY(s)->fill()); }
XLPP_API xlpp_border xlpp_style_border(xlpp_style s)        { return reinterpret_cast<xlpp_border>(&STY(s)->border()); }
XLPP_API xlpp_alignment xlpp_style_alignment(xlpp_style s)   { return reinterpret_cast<xlpp_alignment>(&STY(s)->alignment()); }
XLPP_API void xlpp_style_set_number_format(xlpp_style s, const char* v) { STY(s)->setNumberFormat(v); }

XLPP_API void xlpp_free_string(const char* str) {
    // Internal strings are owned by the C++ objects — no-op
    (void)str;
}

} // extern "C"
