// XLPP C API — Opaque handle-based wrapper for C interop (P/Invoke, FFI)
#pragma once

#ifdef __cplusplus
extern "C" {
#endif

#ifdef _WIN32
  #ifdef XLPP_CAPI_EXPORTS
    #define XLPP_API __declspec(dllexport)
  #else
    #define XLPP_API __declspec(dllimport)
  #endif
#else
  #define XLPP_API
#endif

#ifdef __cplusplus
#include <cstdint>
using std::uint64_t;
#else
#include <stdint.h>
#endif

// Opaque handle types
typedef struct xlpp_workbook_t*    xlpp_workbook;
typedef struct xlpp_worksheet_t*   xlpp_worksheet;
typedef struct xlpp_cell_t*        xlpp_cell;
typedef struct xlpp_font_t*        xlpp_font;
typedef struct xlpp_fill_t*        xlpp_fill;
typedef struct xlpp_border_t*      xlpp_border;
typedef struct xlpp_borderside_t*  xlpp_borderside;
typedef struct xlpp_alignment_t*   xlpp_alignment;
typedef struct xlpp_style_t*       xlpp_style;
typedef struct xlpp_properties_t*  xlpp_properties;
typedef struct xlpp_hyperlink_t*   xlpp_hyperlink;
typedef struct xlpp_comment_t*     xlpp_comment;
typedef struct xlpp_table_t*      xlpp_table;
typedef struct xlpp_data_validation_t* xlpp_data_validation;
typedef struct xlpp_conditional_rule_t* xlpp_conditional_rule;
typedef struct xlpp_chart_t*        xlpp_chart;
typedef struct xlpp_chart_series_t* xlpp_chart_series;

// Cell value types
#define XLPP_VALUE_EMPTY  0
#define XLPP_VALUE_BOOL   1
#define XLPP_VALUE_NUMBER 2
#define XLPP_VALUE_STRING 3
#define XLPP_VALUE_ERROR  4
#define XLPP_VALUE_DATE   5

// ============================================================
// Workbook
// ============================================================
XLPP_API xlpp_workbook  xlpp_workbook_create(void);
XLPP_API void           xlpp_workbook_destroy(xlpp_workbook wb);

XLPP_API xlpp_worksheet xlpp_workbook_add_sheet(xlpp_workbook wb, const char* name);
XLPP_API int            xlpp_workbook_sheet_count(xlpp_workbook wb);
XLPP_API xlpp_worksheet xlpp_workbook_get_sheet(xlpp_workbook wb, int index);
XLPP_API xlpp_worksheet xlpp_workbook_sheet_by_name(xlpp_workbook wb, const char* name);
XLPP_API int            xlpp_workbook_remove_sheet(xlpp_workbook wb, const char* name);

XLPP_API int  xlpp_workbook_load(xlpp_workbook wb, const char* path);
XLPP_API int  xlpp_workbook_save(xlpp_workbook wb, const char* path);
XLPP_API void xlpp_workbook_clear(xlpp_workbook wb);
XLPP_API int  xlpp_workbook_date1904(xlpp_workbook wb);
XLPP_API void xlpp_workbook_set_date1904(xlpp_workbook wb, int enabled);

XLPP_API xlpp_properties xlpp_workbook_properties(xlpp_workbook wb);

// ============================================================
// Properties
// ============================================================
XLPP_API void xlpp_properties_set_title(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_creator(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_subject(xlpp_properties p, const char* v);
XLPP_API const char* xlpp_properties_get_title(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_creator(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_subject(xlpp_properties p);
XLPP_API void xlpp_properties_set_description(xlpp_properties p, const char* v);
XLPP_API const char* xlpp_properties_get_description(xlpp_properties p);

// ============================================================
// Worksheet
// ============================================================
XLPP_API const char*  xlpp_sheet_name(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_rename(xlpp_worksheet ws, const char* name);

XLPP_API xlpp_cell    xlpp_sheet_cell(xlpp_worksheet ws, const char* address);
XLPP_API xlpp_cell    xlpp_sheet_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col);
XLPP_API int          xlpp_sheet_has_cell(xlpp_worksheet ws, const char* address);
XLPP_API void         xlpp_sheet_append_row_numbers(xlpp_worksheet ws, const double* values, int count);

XLPP_API void         xlpp_sheet_append_row(xlpp_worksheet ws, const char** values, int count);
XLPP_API void         xlpp_sheet_append_doubles(xlpp_worksheet ws, const double* values, int count);
XLPP_API void         xlpp_sheet_append_mixed(xlpp_worksheet ws, const char** values, const int* types, int count);

XLPP_API void         xlpp_sheet_merge_cells(xlpp_worksheet ws, const char* range);
XLPP_API void         xlpp_sheet_unmerge_cells(xlpp_worksheet ws, const char* range);
XLPP_API void         xlpp_sheet_freeze_panes(xlpp_worksheet ws, const char* cell);
XLPP_API void         xlpp_sheet_clear_freeze_panes(xlpp_worksheet ws);
XLPP_API int          xlpp_sheet_is_merged(xlpp_worksheet ws, const char* address);
XLPP_API int          xlpp_sheet_merged_range_count(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_set_print_area(xlpp_worksheet ws, const char* value);
XLPP_API const char*  xlpp_sheet_print_area(xlpp_worksheet ws);

XLPP_API uint64_t     xlpp_sheet_max_row(xlpp_worksheet ws);
XLPP_API uint64_t     xlpp_sheet_max_col(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_dimensions(xlpp_worksheet ws, char* out, int outSize);

XLPP_API void         xlpp_sheet_insert_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_delete_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_insert_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_delete_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API int           xlpp_sheet_frozen_pane(xlpp_worksheet ws, char* out, int outSize);
XLPP_API int           xlpp_sheet_merged_range(xlpp_worksheet ws, int index, char* out, int outSize);
XLPP_API xlpp_table    xlpp_sheet_add_table(xlpp_worksheet ws, const char* name, const char* reference);
XLPP_API xlpp_table    xlpp_sheet_table(xlpp_worksheet ws, const char* name);

// ============================================================
// Cell
// ============================================================
XLPP_API const char*  xlpp_cell_address(xlpp_cell c);
XLPP_API uint64_t     xlpp_cell_row(xlpp_cell c);
XLPP_API uint64_t     xlpp_cell_column(xlpp_cell c);

XLPP_API int          xlpp_cell_value_type(xlpp_cell c);
XLPP_API int          xlpp_cell_get_bool(xlpp_cell c);
XLPP_API double       xlpp_cell_get_number(xlpp_cell c);
XLPP_API const char*  xlpp_cell_get_string(xlpp_cell c);
XLPP_API int          xlpp_cell_is_empty(xlpp_cell c);

XLPP_API void         xlpp_cell_set_string(xlpp_cell c, const char* v);
XLPP_API void         xlpp_cell_set_number(xlpp_cell c, double v);
XLPP_API void         xlpp_cell_set_bool(xlpp_cell c, int v);
XLPP_API void         xlpp_cell_set_empty(xlpp_cell c);
XLPP_API void         xlpp_cell_clear(xlpp_cell c);

XLPP_API const char*  xlpp_cell_get_formula(xlpp_cell c);
XLPP_API void         xlpp_cell_set_formula(xlpp_cell c, const char* f);
XLPP_API int          xlpp_cell_has_formula(xlpp_cell c);
XLPP_API void         xlpp_cell_clear_formula(xlpp_cell c);
XLPP_API int          xlpp_cell_has_hyperlink(xlpp_cell c);
XLPP_API void         xlpp_cell_clear_hyperlink(xlpp_cell c);
XLPP_API int          xlpp_cell_has_comment(xlpp_cell c);
XLPP_API void         xlpp_cell_clear_comment(xlpp_cell c);
XLPP_API const char*  xlpp_cell_get_number_format(xlpp_cell c);
XLPP_API void         xlpp_cell_set_number_format(xlpp_cell c, const char* value);

XLPP_API xlpp_style   xlpp_cell_style(xlpp_cell c);
XLPP_API xlpp_font    xlpp_cell_font(xlpp_cell c);
XLPP_API xlpp_fill    xlpp_cell_fill(xlpp_cell c);
XLPP_API xlpp_border  xlpp_cell_border(xlpp_cell c);
XLPP_API xlpp_alignment xlpp_cell_alignment(xlpp_cell c);

XLPP_API void         xlpp_cell_set_hyperlink(xlpp_cell c, const char* url);
XLPP_API void         xlpp_cell_set_comment(xlpp_cell c, const char* text, const char* author);

// ============================================================
// Font
// ============================================================
XLPP_API void         xlpp_font_set_name(xlpp_font f, const char* v);
XLPP_API void         xlpp_font_set_size(xlpp_font f, double v);
XLPP_API void         xlpp_font_set_bold(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_italic(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_underline(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_color(xlpp_font f, const char* argb);
XLPP_API void         xlpp_font_set_strike(xlpp_font f, int v);
XLPP_API const char*  xlpp_font_get_name(xlpp_font f);
XLPP_API double       xlpp_font_get_size(xlpp_font f);
XLPP_API int          xlpp_font_get_bold(xlpp_font f);

// ============================================================
// Fill
// ============================================================
XLPP_API void         xlpp_fill_set_pattern(xlpp_fill f, const char* v);
XLPP_API void         xlpp_fill_set_fg_color(xlpp_fill f, const char* argb);
XLPP_API void         xlpp_fill_set_bg_color(xlpp_fill f, const char* argb);

// ============================================================
// Border
// ============================================================
XLPP_API xlpp_borderside xlpp_border_left(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_right(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_top(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_bottom(xlpp_border b);
XLPP_API void         xlpp_borderside_set_style(xlpp_borderside s, const char* v);
XLPP_API void         xlpp_borderside_set_color(xlpp_borderside s, const char* argb);

// ============================================================
// Alignment
// ============================================================
XLPP_API void         xlpp_alignment_set_horizontal(xlpp_alignment a, const char* v);
XLPP_API void         xlpp_alignment_set_vertical(xlpp_alignment a, const char* v);
XLPP_API void         xlpp_alignment_set_wrap_text(xlpp_alignment a, int v);

// ============================================================
// Style
// ============================================================
XLPP_API xlpp_font    xlpp_style_font(xlpp_style s);
XLPP_API xlpp_fill    xlpp_style_fill(xlpp_style s);
XLPP_API xlpp_border  xlpp_style_border(xlpp_style s);
XLPP_API xlpp_alignment xlpp_style_alignment(xlpp_style s);
XLPP_API void         xlpp_style_set_number_format(xlpp_style s, const char* v);
XLPP_API int           xlpp_style_get_locked(xlpp_style s);
XLPP_API void          xlpp_style_set_locked(xlpp_style s, int v);
XLPP_API int           xlpp_style_get_hidden(xlpp_style s);
XLPP_API void          xlpp_style_set_hidden(xlpp_style s, int v);

// Table
XLPP_API const char*   xlpp_table_name(xlpp_table table);
XLPP_API const char*   xlpp_table_reference(xlpp_table table);
XLPP_API void          xlpp_table_set_reference(xlpp_table table, const char* value);
XLPP_API int            xlpp_table_show_header_row(xlpp_table table);
XLPP_API void           xlpp_table_set_show_header_row(xlpp_table table, int value);
XLPP_API int            xlpp_table_show_totals_row(xlpp_table table);
XLPP_API void           xlpp_table_set_show_totals_row(xlpp_table table, int value);
XLPP_API int            xlpp_table_column_count(xlpp_table table);
XLPP_API void           xlpp_table_add_column(xlpp_table table, const char* name);
XLPP_API const char*    xlpp_table_display_name(xlpp_table table);
XLPP_API void           xlpp_table_set_display_name(xlpp_table table, const char* value);
XLPP_API const char*    xlpp_table_style_name(xlpp_table table);
XLPP_API void           xlpp_table_set_style_name(xlpp_table table, const char* value);
XLPP_API int            xlpp_table_show_row_stripes(xlpp_table table);
XLPP_API void           xlpp_table_set_show_row_stripes(xlpp_table table, int value);

// Data validation
XLPP_API xlpp_data_validation xlpp_sheet_add_list_validation(xlpp_worksheet ws, const char* reference, const char* formula);
XLPP_API void xlpp_validation_set_allow_blank(xlpp_data_validation validation, int value);
XLPP_API void xlpp_validation_set_prompt(xlpp_data_validation validation, const char* title, const char* text);
XLPP_API void xlpp_validation_set_error(xlpp_data_validation validation, const char* title, const char* text);

// Conditional formatting
XLPP_API xlpp_conditional_rule xlpp_sheet_add_formula_rule(xlpp_worksheet ws, const char* reference, const char* formula);
XLPP_API xlpp_conditional_rule xlpp_sheet_add_data_bar_rule(xlpp_worksheet ws, const char* reference, const char* color);
XLPP_API void xlpp_conditional_rule_set_priority(xlpp_conditional_rule rule, uint64_t priority);
XLPP_API void xlpp_conditional_rule_set_stop_if_true(xlpp_conditional_rule rule, int value);
XLPP_API void xlpp_conditional_rule_set_font_color(xlpp_conditional_rule rule, const char* argb);
XLPP_API void xlpp_conditional_rule_set_fill_color(xlpp_conditional_rule rule, const char* argb);
XLPP_API xlpp_chart xlpp_chart_create(int type);
XLPP_API void xlpp_chart_destroy(xlpp_chart chart);
XLPP_API void xlpp_chart_set_title(xlpp_chart chart, const char* title);
XLPP_API void xlpp_chart_set_x_axis_title(xlpp_chart chart, const char* title);
XLPP_API void xlpp_chart_set_y_axis_title(xlpp_chart chart, const char* title);
XLPP_API void xlpp_chart_set_style(xlpp_chart chart, const char* style);
XLPP_API void xlpp_chart_set_grouping(xlpp_chart chart, int grouping);
XLPP_API void xlpp_chart_set_size(xlpp_chart chart, int width, int height);
XLPP_API void xlpp_chart_set_legend(xlpp_chart chart, int show, const char* position);
XLPP_API xlpp_chart_series xlpp_chart_add_series(xlpp_chart chart, const char* title);
XLPP_API void xlpp_chart_series_set_values_reference(xlpp_chart_series series, const char* reference);
XLPP_API void xlpp_chart_series_set_categories_reference(xlpp_chart_series series, const char* reference);
XLPP_API int xlpp_sheet_add_chart(xlpp_worksheet ws, xlpp_chart chart);

// ============================================================
// Utility
// ============================================================
XLPP_API const char*  xlpp_version(void);
XLPP_API const char*  xlpp_last_error(void);
XLPP_API void         xlpp_clear_error(void);
XLPP_API void         xlpp_free_string(const char* str);

#ifdef __cplusplus
}
#endif
