// XLPP C API Implementation
#include "xlpp_capi.h"
#include <XLPP/XLPP.h>
#include <string>
#include <vector>
#include <cstring>

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
    return "1.0.0";
}

XLPP_API xlpp_workbook xlpp_workbook_create(void) {
    return reinterpret_cast<xlpp_workbook>(new xlpp::Workbook());
}

XLPP_API void xlpp_workbook_destroy(xlpp_workbook wb) {
    delete WB(wb);
}

XLPP_API xlpp_worksheet xlpp_workbook_add_sheet(xlpp_workbook wb, const char* name) {
    // addWorksheet throws on empty/duplicate names; never let a C++ exception
    // cross the C ABI.
    try {
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->addWorksheet(name));
    } catch (...) { return nullptr; }
}

XLPP_API int xlpp_workbook_sheet_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->sheetCount());
}

XLPP_API xlpp_worksheet xlpp_workbook_get_sheet(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_worksheet>(&(*WB(wb))[static_cast<std::size_t>(index)]);
    } catch (...) { return nullptr; }
}

XLPP_API xlpp_worksheet xlpp_workbook_sheet_by_name(xlpp_workbook wb, const char* name) {
    auto* ws = WB(wb)->worksheet(name);
    return reinterpret_cast<xlpp_worksheet>(ws);
}

XLPP_API int xlpp_workbook_remove_sheet(xlpp_workbook wb, const char* name) {
    return WB(wb)->removeWorksheet(name) ? 1 : 0;
}

XLPP_API int xlpp_workbook_load(xlpp_workbook wb, const char* path) {
    try {
        WB(wb)->load(std::filesystem::path(path));
        return 1;
    } catch (...) { return 0; }
}

XLPP_API int xlpp_workbook_save(xlpp_workbook wb, const char* path) {
    try {
        WB(wb)->save(std::filesystem::path(path));
        return 1;
    } catch (...) { return 0; }
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
