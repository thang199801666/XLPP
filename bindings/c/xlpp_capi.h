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
typedef struct xlpp_richtext_t*    xlpp_richtext;
typedef struct xlpp_richtextrun_t* xlpp_richtextrun;
typedef struct xlpp_namedstyle_t*  xlpp_namedstyle;
typedef struct xlpp_definedname_t* xlpp_definedname;
typedef struct xlpp_cellrange_t*   xlpp_cellrange;
typedef struct xlpp_rowdim_t*      xlpp_rowdim;
typedef struct xlpp_coldim_t*      xlpp_coldim;
typedef struct xlpp_autofilter_t*  xlpp_autofilter;
typedef struct xlpp_filtercol_t*   xlpp_filtercol;
typedef struct xlpp_sortstate_t*   xlpp_sortstate;
typedef struct xlpp_pagesetup_t*   xlpp_pagesetup;
typedef struct xlpp_pagemargins_t* xlpp_pagemargins;
typedef struct xlpp_printopts_t*   xlpp_printopts;
typedef struct xlpp_headerfooter_t* xlpp_headerfooter;
typedef struct xlpp_wssprotection_t* xlpp_wssprotection;
typedef struct xlpp_wbprotection_t* xlpp_wbprotection;
typedef struct xlpp_sheetview_t*   xlpp_sheetview;
typedef struct xlpp_image_t*       xlpp_image;
typedef struct xlpp_table_t*       xlpp_table;
typedef struct xlpp_tablecolumn_t* xlpp_tablecolumn;
typedef struct xlpp_tablestyle_t*  xlpp_tablestyle;
typedef struct xlpp_chart_t*       xlpp_chart;
typedef struct xlpp_chartsheet_t*  xlpp_chartsheet;
typedef struct xlpp_chartseries_t* xlpp_chartseries;
typedef struct xlpp_pivottable_t*  xlpp_pivottable;
typedef struct xlpp_pivotfield_t*  xlpp_pivotfield;
typedef struct xlpp_pivotcache_t*  xlpp_pivotcache;
typedef struct xlpp_calcprops_t*   xlpp_calcprops;
typedef struct xlpp_customprops_t* xlpp_customprops;
typedef struct xlpp_customprop_t*  xlpp_customprop;
typedef struct xlpp_cfcollection_t* xlpp_cfcollection;
typedef struct xlpp_cfentry_t*     xlpp_cfentry;
typedef struct xlpp_cfrule_t*      xlpp_cfrule;
typedef struct xlpp_dvcollection_t* xlpp_dvcollection;
typedef struct xlpp_datavalidation_t* xlpp_datavalidation;
typedef struct xlpp_stream_writer_t* xlpp_stream_writer;
typedef struct xlpp_stream_reader_t* xlpp_stream_reader;

// Cell value types
#define XLPP_VALUE_EMPTY  0
#define XLPP_VALUE_BOOL   1
#define XLPP_VALUE_NUMBER 2
#define XLPP_VALUE_STRING 3
#define XLPP_VALUE_ERROR  4
#define XLPP_VALUE_DATE   5

// CellError values
#define XLPP_ERROR_NULL             0
#define XLPP_ERROR_DIV0             1
#define XLPP_ERROR_VALUE            2
#define XLPP_ERROR_REF              3
#define XLPP_ERROR_NAME             4
#define XLPP_ERROR_NUM              5
#define XLPP_ERROR_NA               6
#define XLPP_ERROR_GETTING_DATA     7

// CompressionLevel
#define XLPP_COMPRESS_STORE    0
#define XLPP_COMPRESS_FASTEST  1
#define XLPP_COMPRESS_DEFAULT  2
#define XLPP_COMPRESS_BEST     3

// CompressionStrategy
#define XLPP_STRATEGY_DEFAULT      0
#define XLPP_STRATEGY_FILTERED     1
#define XLPP_STRATEGY_HUFFMAN_ONLY 2
#define XLPP_STRATEGY_RLE          3
#define XLPP_STRATEGY_FIXED        4


// PackageEncryptionMode
#define XLPP_ENCRYPTION_AGILE    0
#define XLPP_ENCRYPTION_STANDARD 1

// PackageEncryptionHash
#define XLPP_ENCRYPTION_HASH_SHA1   0
#define XLPP_ENCRYPTION_HASH_SHA256 1
#define XLPP_ENCRYPTION_HASH_SHA384 2
#define XLPP_ENCRYPTION_HASH_SHA512 3

// PackageEncryptionFormat inspection values
#define XLPP_ENCRYPTION_FORMAT_NONE        0
#define XLPP_ENCRYPTION_FORMAT_AGILE       1
#define XLPP_ENCRYPTION_FORMAT_STANDARD    2
#define XLPP_ENCRYPTION_FORMAT_UNSUPPORTED 3

// SharedStringMode
#define XLPP_SSM_DISABLED   0
#define XLPP_SSM_HASH       1
#define XLPP_SSM_BOUNDED_LRU 2

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
XLPP_API int            xlpp_workbook_rename_sheet(xlpp_workbook wb, const char* old_name, const char* new_name);
XLPP_API xlpp_worksheet xlpp_workbook_copy_sheet(xlpp_workbook wb, xlpp_worksheet src, const char* new_name);
XLPP_API int            xlpp_workbook_sheet_index(xlpp_workbook wb, xlpp_worksheet ws);
XLPP_API const char*    xlpp_workbook_sheet_name(xlpp_workbook wb, int index, char* out, int outSize);
XLPP_API int            xlpp_workbook_sheet_names_count(xlpp_workbook wb);

// P1T mixed worksheet/chartsheet workbook tab model. Legacy sheet_* APIs above
// remain worksheet-only for source compatibility. kind: 0=worksheet, 1=chartsheet.
XLPP_API int             xlpp_workbook_tab_count(xlpp_workbook wb);
XLPP_API const char*     xlpp_workbook_tab_name(xlpp_workbook wb, int index, char* out, int outSize);
XLPP_API int             xlpp_workbook_tab_kind(xlpp_workbook wb, int index);
// visibility: 0=visible, 1=hidden, 2=veryHidden.
XLPP_API int             xlpp_workbook_tab_visibility(xlpp_workbook wb, int index);
XLPP_API int             xlpp_workbook_set_tab_visibility(xlpp_workbook wb, int index, int visibility);
XLPP_API int             xlpp_workbook_active_tab(xlpp_workbook wb);
XLPP_API int             xlpp_workbook_set_active_tab(xlpp_workbook wb, int index);
XLPP_API int             xlpp_workbook_move_tab(xlpp_workbook wb, int from_index, int to_index);
XLPP_API xlpp_chartsheet xlpp_workbook_add_chartsheet(xlpp_workbook wb, const char* name, int chart_type);
XLPP_API int             xlpp_workbook_chartsheet_count(xlpp_workbook wb);
XLPP_API xlpp_chartsheet xlpp_workbook_chartsheet_at(xlpp_workbook wb, int index);
XLPP_API xlpp_chartsheet xlpp_workbook_chartsheet_by_name(xlpp_workbook wb, const char* name);
XLPP_API int             xlpp_workbook_rename_chartsheet(xlpp_workbook wb, const char* old_name, const char* new_name);
XLPP_API int             xlpp_workbook_remove_chartsheet(xlpp_workbook wb, const char* name);
XLPP_API const char*     xlpp_chartsheet_name(xlpp_chartsheet cs);
XLPP_API xlpp_chart      xlpp_chartsheet_chart(xlpp_chartsheet cs);
// Opaque DevMode printer-settings payload owned by the Chartsheet pageSetup.
XLPP_API int             xlpp_chartsheet_set_printer_settings(xlpp_chartsheet cs, const unsigned char* data, uint64_t size);
XLPP_API uint64_t        xlpp_chartsheet_printer_settings_size(xlpp_chartsheet cs);
XLPP_API uint64_t        xlpp_chartsheet_copy_printer_settings(xlpp_chartsheet cs, unsigned char* out, uint64_t capacity);
XLPP_API void            xlpp_chartsheet_clear_printer_settings(xlpp_chartsheet cs);

XLPP_API int  xlpp_workbook_load(xlpp_workbook wb, const char* path);
XLPP_API int  xlpp_workbook_save(xlpp_workbook wb, const char* path);
XLPP_API void xlpp_workbook_set_template(xlpp_workbook wb, int enabled);
XLPP_API int  xlpp_workbook_is_template(xlpp_workbook wb);
XLPP_API int  xlpp_workbook_load_password(xlpp_workbook wb, const char* path, const char* password_utf8);
XLPP_API int  xlpp_workbook_load_password_ex(xlpp_workbook wb, const char* path, const char* password_utf8, uint64_t max_spin_count, uint64_t max_decrypted_package_bytes, int allow_standard_encryption, int require_agile_data_integrity, uint64_t max_encryption_info_bytes);
XLPP_API int  xlpp_workbook_save_password(xlpp_workbook wb, const char* path, const char* password_utf8, uint64_t spin_count);
XLPP_API int  xlpp_workbook_save_password_ex(xlpp_workbook wb, const char* path, const char* password_utf8, int mode, unsigned key_bits, int hash_algorithm, uint64_t spin_count);
XLPP_API int  xlpp_workbook_is_password_encrypted_file(const char* path);
XLPP_API int  xlpp_workbook_encryption_profile(const char* path, int* format, unsigned* key_bits, int* hash_algorithm, uint64_t* spin_count, int* has_data_integrity);
XLPP_API int  xlpp_workbook_encryption_key_encryptor_counts(const char* path, uint64_t* total_key_encryptors, uint64_t* password_key_encryptors, uint64_t* certificate_key_encryptors);
XLPP_API void xlpp_workbook_set_date1904(xlpp_workbook wb, int v);
XLPP_API int  xlpp_workbook_date1904(xlpp_workbook wb);
XLPP_API void xlpp_workbook_clear(xlpp_workbook wb);
XLPP_API int  xlpp_workbook_strict_namespaces(xlpp_workbook wb);

XLPP_API xlpp_properties xlpp_workbook_properties(xlpp_workbook wb);
XLPP_API xlpp_wbprotection xlpp_workbook_protection(xlpp_workbook wb);
XLPP_API xlpp_calcprops  xlpp_workbook_calc_properties(xlpp_workbook wb);
XLPP_API xlpp_customprops xlpp_workbook_custom_properties(xlpp_workbook wb);

XLPP_API xlpp_namedstyle xlpp_workbook_add_named_style(xlpp_workbook wb, const char* name, int* ok);
XLPP_API xlpp_namedstyle xlpp_workbook_named_style(xlpp_workbook wb, const char* name);
XLPP_API int             xlpp_workbook_named_styles_count(xlpp_workbook wb);
XLPP_API xlpp_namedstyle xlpp_workbook_named_style_at(xlpp_workbook wb, int index);
XLPP_API void            xlpp_workbook_apply_named_style(xlpp_workbook wb, xlpp_cell c, const char* name);

XLPP_API xlpp_definedname xlpp_workbook_add_defined_name(xlpp_workbook wb, const char* name, const char* value, int* ok);
XLPP_API xlpp_definedname xlpp_workbook_defined_name(xlpp_workbook wb, const char* name);
XLPP_API int              xlpp_workbook_defined_names_count(xlpp_workbook wb);
XLPP_API xlpp_definedname xlpp_workbook_defined_name_at(xlpp_workbook wb, int index);

// ============================================================
// Properties
// ============================================================
XLPP_API void xlpp_properties_set_title(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_creator(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_subject(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_description(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_keywords(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_category(xlpp_properties p, const char* v);
XLPP_API void xlpp_properties_set_last_modified_by(xlpp_properties p, const char* v);
XLPP_API const char* xlpp_properties_get_title(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_creator(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_subject(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_description(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_keywords(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_category(xlpp_properties p);
XLPP_API const char* xlpp_properties_get_last_modified_by(xlpp_properties p);

// ============================================================
// Workbook protection
// ============================================================
XLPP_API void xlpp_wbprotection_set_lock_structure(xlpp_wbprotection p, int v);
XLPP_API int  xlpp_wbprotection_lock_structure(xlpp_wbprotection p);
XLPP_API void xlpp_wbprotection_set_lock_windows(xlpp_wbprotection p, int v);
XLPP_API int  xlpp_wbprotection_lock_windows(xlpp_wbprotection p);
XLPP_API void xlpp_wbprotection_set_lock_revision(xlpp_wbprotection p, int v);
XLPP_API int  xlpp_wbprotection_lock_revision(xlpp_wbprotection p);
XLPP_API void xlpp_wbprotection_set_password_hash(xlpp_wbprotection p, const char* v);
XLPP_API const char* xlpp_wbprotection_password_hash(xlpp_wbprotection p);

// ============================================================
// CalcProperties
// ============================================================
XLPP_API void xlpp_calcprops_set_calc_id(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_calc_id(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_calc_mode(xlpp_calcprops p, const char* v);
XLPP_API const char* xlpp_calcprops_calc_mode(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_calc_on_save(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_calc_on_save(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_full_calc_on_load(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_full_calc_on_load(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_full_precision(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_full_precision(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_iterate(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_iterate(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_iterate_count(xlpp_calcprops p, int v);
XLPP_API int  xlpp_calcprops_iterate_count(xlpp_calcprops p);
XLPP_API void xlpp_calcprops_set_iterate_delta(xlpp_calcprops p, double v);
XLPP_API double xlpp_calcprops_iterate_delta(xlpp_calcprops p);

// ============================================================
// Custom properties
// ============================================================
XLPP_API xlpp_customprop xlpp_customprops_add(xlpp_customprops c, const char* name, const char* value, int type);
XLPP_API int  xlpp_customprops_count(xlpp_customprops c);
XLPP_API xlpp_customprop xlpp_customprops_at(xlpp_customprops c, int index);
XLPP_API const char* xlpp_customprop_name(xlpp_customprop p);
XLPP_API const char* xlpp_customprop_value(xlpp_customprop p);
XLPP_API const char* xlpp_customprop_type(xlpp_customprop p);

// ============================================================
// Worksheet
// ============================================================
XLPP_API const char*  xlpp_sheet_name(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_rename(xlpp_worksheet ws, const char* name);

XLPP_API xlpp_cell    xlpp_sheet_cell(xlpp_worksheet ws, const char* address);
XLPP_API xlpp_cell    xlpp_sheet_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col);
XLPP_API int          xlpp_sheet_has_cell(xlpp_worksheet ws, const char* address);
XLPP_API int          xlpp_sheet_has_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col);

XLPP_API void         xlpp_sheet_append_row(xlpp_worksheet ws, const char** values, int count);
XLPP_API void         xlpp_sheet_append_doubles(xlpp_worksheet ws, const double* values, int count);
XLPP_API void         xlpp_sheet_append_values(xlpp_worksheet ws, const double* nums, const int* types, int count);

XLPP_API xlpp_cellrange xlpp_sheet_range(xlpp_worksheet ws, const char* address);
XLPP_API xlpp_cellrange xlpp_sheet_range_rc(xlpp_worksheet ws, uint64_t minRow, uint64_t minCol, uint64_t maxRow, uint64_t maxCol);

XLPP_API void         xlpp_sheet_merge_cells(xlpp_worksheet ws, const char* range);
XLPP_API void         xlpp_sheet_unmerge_cells(xlpp_worksheet ws, const char* range);
XLPP_API int          xlpp_sheet_is_merged(xlpp_worksheet ws, const char* cell);
XLPP_API int          xlpp_sheet_merged_count(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_merged_at(xlpp_worksheet ws, int index, char* out, int outSize);

XLPP_API void         xlpp_sheet_freeze_panes(xlpp_worksheet ws, const char* cell);
XLPP_API void         xlpp_sheet_clear_freeze_panes(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_frozen_pane(xlpp_worksheet ws, char* out, int outSize);

XLPP_API xlpp_rowdim  xlpp_sheet_row_dimension(xlpp_worksheet ws, uint64_t row);
XLPP_API xlpp_coldim  xlpp_sheet_col_dimension(xlpp_worksheet ws, uint64_t col);

XLPP_API uint64_t     xlpp_sheet_max_row(xlpp_worksheet ws);
XLPP_API uint64_t     xlpp_sheet_max_col(xlpp_worksheet ws);
XLPP_API void         xlpp_sheet_dimensions(xlpp_worksheet ws, char* out, int outSize);
XLPP_API int          xlpp_sheet_empty(xlpp_worksheet ws);
XLPP_API uint64_t     xlpp_sheet_row_count(xlpp_worksheet ws);
XLPP_API uint64_t     xlpp_sheet_col_count(xlpp_worksheet ws);

XLPP_API void         xlpp_sheet_insert_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_delete_rows(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_insert_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount);
XLPP_API void         xlpp_sheet_delete_cols(xlpp_worksheet ws, uint64_t index, uint64_t amount);

XLPP_API void         xlpp_sheet_set_print_area(xlpp_worksheet ws, const char* v);
XLPP_API void         xlpp_sheet_print_area(xlpp_worksheet ws, char* out, int outSize);
XLPP_API void         xlpp_sheet_set_print_titles_rows(xlpp_worksheet ws, const char* v);
XLPP_API void         xlpp_sheet_print_titles_rows(xlpp_worksheet ws, char* out, int outSize);
XLPP_API void         xlpp_sheet_set_print_titles_cols(xlpp_worksheet ws, const char* v);
XLPP_API void         xlpp_sheet_print_titles_cols(xlpp_worksheet ws, char* out, int outSize);

XLPP_API xlpp_autofilter xlpp_sheet_auto_filter(xlpp_worksheet ws);
XLPP_API xlpp_cfcollection xlpp_sheet_conditional_formatting(xlpp_worksheet ws);
XLPP_API xlpp_dvcollection xlpp_sheet_data_validations(xlpp_worksheet ws);
XLPP_API xlpp_pagesetup xlpp_sheet_page_setup(xlpp_worksheet ws);
XLPP_API xlpp_pagemargins xlpp_sheet_page_margins(xlpp_worksheet ws);
XLPP_API xlpp_printopts xlpp_sheet_print_options(xlpp_worksheet ws);
XLPP_API xlpp_headerfooter xlpp_sheet_header_footer(xlpp_worksheet ws);
XLPP_API xlpp_wssprotection xlpp_sheet_protection(xlpp_worksheet ws);
XLPP_API xlpp_sheetview xlpp_sheet_view(xlpp_worksheet ws);

XLPP_API xlpp_table xlpp_sheet_add_table(xlpp_worksheet ws, const char* name, const char* reference, int* ok);
XLPP_API xlpp_table xlpp_sheet_table(xlpp_worksheet ws, const char* name);
XLPP_API int         xlpp_sheet_table_count(xlpp_worksheet ws);
XLPP_API xlpp_table  xlpp_sheet_table_at(xlpp_worksheet ws, int index);

XLPP_API xlpp_image xlpp_sheet_add_image(xlpp_worksheet ws, const char* path, const char* anchor, int* ok);
XLPP_API int        xlpp_sheet_image_count(xlpp_worksheet ws);
XLPP_API xlpp_image xlpp_sheet_image_at(xlpp_worksheet ws, int index);

XLPP_API void xlpp_sheet_add_chart(xlpp_worksheet ws, int type);
XLPP_API int  xlpp_sheet_chart_count(xlpp_worksheet ws);
XLPP_API xlpp_chart xlpp_sheet_chart_at(xlpp_worksheet ws, int index);

XLPP_API void xlpp_sheet_add_pivot(xlpp_worksheet ws, const char* name, const char* location);
XLPP_API int  xlpp_sheet_pivot_count(xlpp_worksheet ws);
XLPP_API xlpp_pivottable xlpp_sheet_pivot_at(xlpp_worksheet ws, int index);

// ============================================================
// Row / Column dimensions
// ============================================================
XLPP_API void xlpp_rowdim_set_height(xlpp_rowdim d, double v);
XLPP_API int  xlpp_rowdim_has_height(xlpp_rowdim d);
XLPP_API double xlpp_rowdim_height(xlpp_rowdim d);
XLPP_API void xlpp_rowdim_set_hidden(xlpp_rowdim d, int v);
XLPP_API int  xlpp_rowdim_hidden(xlpp_rowdim d);
XLPP_API void xlpp_rowdim_set_outline_level(xlpp_rowdim d, int v);
XLPP_API int  xlpp_rowdim_outline_level(xlpp_rowdim d);
XLPP_API void xlpp_rowdim_set_collapsed(xlpp_rowdim d, int v);
XLPP_API int  xlpp_rowdim_collapsed(xlpp_rowdim d);

XLPP_API void xlpp_coldim_set_width(xlpp_coldim d, double v);
XLPP_API int  xlpp_coldim_has_width(xlpp_coldim d);
XLPP_API double xlpp_coldim_width(xlpp_coldim d);
XLPP_API void xlpp_coldim_set_hidden(xlpp_coldim d, int v);
XLPP_API int  xlpp_coldim_hidden(xlpp_coldim d);
XLPP_API void xlpp_coldim_set_best_fit(xlpp_coldim d, int v);
XLPP_API int  xlpp_coldim_best_fit(xlpp_coldim d);
XLPP_API void xlpp_coldim_set_outline_level(xlpp_coldim d, int v);
XLPP_API int  xlpp_coldim_outline_level(xlpp_coldim d);
XLPP_API void xlpp_coldim_set_collapsed(xlpp_coldim d, int v);
XLPP_API int  xlpp_coldim_collapsed(xlpp_coldim d);

// ============================================================
// CellRange
// ============================================================
XLPP_API uint64_t xlpp_range_min_row(xlpp_cellrange r);
XLPP_API uint64_t xlpp_range_min_col(xlpp_cellrange r);
XLPP_API uint64_t xlpp_range_max_row(xlpp_cellrange r);
XLPP_API uint64_t xlpp_range_max_col(xlpp_cellrange r);
XLPP_API uint64_t xlpp_range_row_count(xlpp_cellrange r);
XLPP_API uint64_t xlpp_range_col_count(xlpp_cellrange r);
XLPP_API void     xlpp_range_address(xlpp_cellrange r, char* out, int outSize);
XLPP_API xlpp_cell xlpp_range_cell(xlpp_cellrange r, uint64_t relRow, uint64_t relCol);
XLPP_API void     xlpp_range_set_value(xlpp_cellrange r, double v);
XLPP_API void     xlpp_range_set_string(xlpp_cellrange r, const char* v);
XLPP_API void     xlpp_range_clear(xlpp_cellrange r);
XLPP_API void     xlpp_range_values(xlpp_cellrange r, double* out, int* outCount);

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
XLPP_API int          xlpp_cell_has_value(xlpp_cell c);
XLPP_API int          xlpp_cell_is_numeric(xlpp_cell c);
XLPP_API int          xlpp_cell_is_string(xlpp_cell c);
XLPP_API int          xlpp_cell_is_bool(xlpp_cell c);
XLPP_API int          xlpp_cell_is_date(xlpp_cell c);
XLPP_API int          xlpp_cell_is_error(xlpp_cell c);
XLPP_API int          xlpp_cell_error_code(xlpp_cell c);
XLPP_API int          xlpp_cell_date(xlpp_cell c, int* year, int* month, int* day, int* hour, int* minute, double* second);

XLPP_API void         xlpp_cell_set_string(xlpp_cell c, const char* v);
XLPP_API void         xlpp_cell_set_number(xlpp_cell c, double v);
XLPP_API void         xlpp_cell_set_bool(xlpp_cell c, int v);
XLPP_API void         xlpp_cell_set_empty(xlpp_cell c);
XLPP_API void         xlpp_cell_set_error(xlpp_cell c, int errorCode);
XLPP_API void         xlpp_cell_set_date(xlpp_cell c, int year, int month, int day, int hour, int minute, double second, int hasTime);
XLPP_API void         xlpp_cell_clear(xlpp_cell c);

XLPP_API const char*  xlpp_cell_get_formula(xlpp_cell c);
XLPP_API void         xlpp_cell_set_formula(xlpp_cell c, const char* f);
XLPP_API void         xlpp_cell_set_shared_formula(xlpp_cell c, const char* f, unsigned sharedIndex, const char* reference);
XLPP_API void         xlpp_cell_set_array_formula(xlpp_cell c, const char* f, const char* reference);
XLPP_API void         xlpp_cell_set_dynamic_array_formula(xlpp_cell c, const char* f, const char* reference);
XLPP_API int          xlpp_cell_has_formula(xlpp_cell c);
XLPP_API void         xlpp_cell_clear_formula(xlpp_cell c);

XLPP_API xlpp_style   xlpp_cell_style(xlpp_cell c);
XLPP_API xlpp_font    xlpp_cell_font(xlpp_cell c);
XLPP_API xlpp_fill    xlpp_cell_fill(xlpp_cell c);
XLPP_API xlpp_border  xlpp_cell_border(xlpp_cell c);
XLPP_API xlpp_alignment xlpp_cell_alignment(xlpp_cell c);
XLPP_API void         xlpp_cell_set_number_format(xlpp_cell c, const char* v);
XLPP_API void         xlpp_cell_number_format(xlpp_cell c, char* out, int outSize);
XLPP_API void         xlpp_cell_set_named_style(xlpp_cell c, const char* name);
XLPP_API void         xlpp_cell_named_style(xlpp_cell c, char* out, int outSize);

XLPP_API int          xlpp_cell_has_hyperlink(xlpp_cell c);
XLPP_API xlpp_hyperlink xlpp_cell_hyperlink(xlpp_cell c);
XLPP_API void         xlpp_cell_set_hyperlink(xlpp_cell c, const char* url);
XLPP_API void         xlpp_cell_set_hyperlink_full(xlpp_cell c, const char* url, const char* display, const char* tooltip, int external);
XLPP_API void         xlpp_cell_clear_hyperlink(xlpp_cell c);

XLPP_API int          xlpp_cell_has_comment(xlpp_cell c);
XLPP_API xlpp_comment xlpp_cell_comment(xlpp_cell c);
XLPP_API void         xlpp_cell_set_comment(xlpp_cell c, const char* text, const char* author);
XLPP_API void         xlpp_cell_clear_comment(xlpp_cell c);

// ============================================================
// Hyperlink / Comment
// ============================================================
XLPP_API void xlpp_hyperlink_set_target(xlpp_hyperlink h, const char* v);
XLPP_API void xlpp_hyperlink_set_display(xlpp_hyperlink h, const char* v);
XLPP_API void xlpp_hyperlink_set_tooltip(xlpp_hyperlink h, const char* v);
XLPP_API void xlpp_hyperlink_set_external(xlpp_hyperlink h, int v);
XLPP_API const char* xlpp_hyperlink_target(xlpp_hyperlink h);
XLPP_API const char* xlpp_hyperlink_display(xlpp_hyperlink h);
XLPP_API const char* xlpp_hyperlink_tooltip(xlpp_hyperlink h);
XLPP_API int  xlpp_hyperlink_external(xlpp_hyperlink h);

XLPP_API void xlpp_comment_set_text(xlpp_comment c, const char* v);
XLPP_API void xlpp_comment_set_author(xlpp_comment c, const char* v);
XLPP_API const char* xlpp_comment_text(xlpp_comment c);
XLPP_API const char* xlpp_comment_author(xlpp_comment c);

// ============================================================
// RichText
// ============================================================
XLPP_API xlpp_richtext xlpp_richtext_create(void);
XLPP_API void xlpp_richtext_destroy(xlpp_richtext rt);
XLPP_API int  xlpp_richtext_run_count(xlpp_richtext rt);
XLPP_API xlpp_richtextrun xlpp_richtext_add_run(xlpp_richtext rt, const char* text);
XLPP_API xlpp_richtextrun xlpp_richtext_run_at(xlpp_richtext rt, int index);
XLPP_API int  xlpp_richtext_empty(xlpp_richtext rt);
XLPP_API void xlpp_richtext_plain_text(xlpp_richtext rt, char* out, int outSize);

XLPP_API void xlpp_richtextrun_set_text(xlpp_richtextrun r, const char* v);
XLPP_API void xlpp_richtextrun_set_bold(xlpp_richtextrun r, int v);
XLPP_API void xlpp_richtextrun_set_italic(xlpp_richtextrun r, int v);
XLPP_API void xlpp_richtextrun_set_underline(xlpp_richtextrun r, int v);
XLPP_API void xlpp_richtextrun_set_strike(xlpp_richtextrun r, int v);
XLPP_API void xlpp_richtextrun_set_color(xlpp_richtextrun r, const char* v);
XLPP_API void xlpp_richtextrun_set_size(xlpp_richtextrun r, double v);
XLPP_API void xlpp_richtextrun_set_font(xlpp_richtextrun r, const char* v);
XLPP_API const char* xlpp_richtextrun_text(xlpp_richtextrun r);

// ============================================================
// Named style / Defined name
// ============================================================
XLPP_API void xlpp_namedstyle_set_name(xlpp_namedstyle s, const char* v);
XLPP_API const char* xlpp_namedstyle_name(xlpp_namedstyle s);
XLPP_API xlpp_style xlpp_namedstyle_style(xlpp_namedstyle s);

XLPP_API void xlpp_definedname_set_value(xlpp_definedname d, const char* v);
XLPP_API const char* xlpp_definedname_name(xlpp_definedname d);
XLPP_API const char* xlpp_definedname_value(xlpp_definedname d);
XLPP_API void xlpp_definedname_set_local_sheet_id(xlpp_definedname d, uint64_t v);
XLPP_API void xlpp_definedname_clear_local_sheet_id(xlpp_definedname d);
XLPP_API int  xlpp_definedname_has_local_sheet_id(xlpp_definedname d);
XLPP_API void xlpp_definedname_set_hidden(xlpp_definedname d, int v);
XLPP_API int  xlpp_definedname_hidden(xlpp_definedname d);
XLPP_API void xlpp_definedname_set_comment(xlpp_definedname d, const char* v);
XLPP_API const char* xlpp_definedname_comment(xlpp_definedname d);

// ============================================================
// AutoFilter / Sort
// ============================================================
XLPP_API void xlpp_autofilter_set_reference(xlpp_autofilter f, const char* v);
XLPP_API void xlpp_autofilter_reference(xlpp_autofilter f, char* out, int outSize);
XLPP_API int  xlpp_autofilter_enabled(xlpp_autofilter f);
XLPP_API void xlpp_autofilter_clear(xlpp_autofilter f);
XLPP_API xlpp_filtercol xlpp_autofilter_column(xlpp_autofilter f, uint64_t columnId);
XLPP_API xlpp_sortstate xlpp_autofilter_sort_state(xlpp_autofilter f);

XLPP_API uint64_t xlpp_filtercol_column_id(xlpp_filtercol c);
XLPP_API void xlpp_filtercol_add_value(xlpp_filtercol c, const char* v);
XLPP_API void xlpp_filtercol_clear_values(xlpp_filtercol c);
XLPP_API int  xlpp_filtercol_value_count(xlpp_filtercol c);
XLPP_API void xlpp_filtercol_value_at(xlpp_filtercol c, int index, char* out, int outSize);
XLPP_API void xlpp_filtercol_set_and_mode(xlpp_filtercol c, int v);
XLPP_API int  xlpp_filtercol_and_mode(xlpp_filtercol c);
XLPP_API void xlpp_filtercol_set_include_blank(xlpp_filtercol c, int v);
XLPP_API int  xlpp_filtercol_include_blank(xlpp_filtercol c);

XLPP_API void xlpp_sortstate_set_reference(xlpp_sortstate s, const char* v);
XLPP_API void xlpp_sortstate_reference(xlpp_sortstate s, char* out, int outSize);
XLPP_API void xlpp_sortstate_set_case_sensitive(xlpp_sortstate s, int v);
XLPP_API int  xlpp_sortstate_case_sensitive(xlpp_sortstate s);
XLPP_API void xlpp_sortstate_add_condition(xlpp_sortstate s, const char* reference, int descending);
XLPP_API void xlpp_sortstate_clear(xlpp_sortstate s);

// ============================================================
// Page setup / margins / print options / header footer
// ============================================================
XLPP_API void xlpp_pagesetup_set_orientation(xlpp_pagesetup p, int v);
XLPP_API int  xlpp_pagesetup_orientation(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_paper_size(xlpp_pagesetup p, int v);
XLPP_API int  xlpp_pagesetup_paper_size(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_scale(xlpp_pagesetup p, unsigned v);
XLPP_API unsigned xlpp_pagesetup_scale(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_fit_to_width(xlpp_pagesetup p, unsigned v);
XLPP_API unsigned xlpp_pagesetup_fit_to_width(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_fit_to_height(xlpp_pagesetup p, unsigned v);
XLPP_API unsigned xlpp_pagesetup_fit_to_height(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_fit_to_page(xlpp_pagesetup p, int v);
XLPP_API int  xlpp_pagesetup_fit_to_page(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_black_and_white(xlpp_pagesetup p, int v);
XLPP_API int  xlpp_pagesetup_black_and_white(xlpp_pagesetup p);
XLPP_API void xlpp_pagesetup_set_draft(xlpp_pagesetup p, int v);
XLPP_API int  xlpp_pagesetup_draft(xlpp_pagesetup p);

XLPP_API void xlpp_pagemargins_set_left(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_left(xlpp_pagemargins m);
XLPP_API void xlpp_pagemargins_set_right(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_right(xlpp_pagemargins m);
XLPP_API void xlpp_pagemargins_set_top(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_top(xlpp_pagemargins m);
XLPP_API void xlpp_pagemargins_set_bottom(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_bottom(xlpp_pagemargins m);
XLPP_API void xlpp_pagemargins_set_header(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_header(xlpp_pagemargins m);
XLPP_API void xlpp_pagemargins_set_footer(xlpp_pagemargins m, double v);
XLPP_API double xlpp_pagemargins_footer(xlpp_pagemargins m);

XLPP_API void xlpp_printopts_set_horizontal_centered(xlpp_printopts p, int v);
XLPP_API int  xlpp_printopts_horizontal_centered(xlpp_printopts p);
XLPP_API void xlpp_printopts_set_vertical_centered(xlpp_printopts p, int v);
XLPP_API int  xlpp_printopts_vertical_centered(xlpp_printopts p);
XLPP_API void xlpp_printopts_set_headings(xlpp_printopts p, int v);
XLPP_API int  xlpp_printopts_headings(xlpp_printopts p);
XLPP_API void xlpp_printopts_set_grid_lines(xlpp_printopts p, int v);
XLPP_API int  xlpp_printopts_grid_lines(xlpp_printopts p);

XLPP_API void xlpp_headerfooter_set_odd_header(xlpp_headerfooter h, const char* v);
XLPP_API void xlpp_headerfooter_set_odd_footer(xlpp_headerfooter h, const char* v);
XLPP_API void xlpp_headerfooter_set_even_header(xlpp_headerfooter h, const char* v);
XLPP_API void xlpp_headerfooter_set_even_footer(xlpp_headerfooter h, const char* v);
XLPP_API void xlpp_headerfooter_set_different_odd_even(xlpp_headerfooter h, int v);
XLPP_API int  xlpp_headerfooter_different_odd_even(xlpp_headerfooter h);
XLPP_API void xlpp_headerfooter_set_different_first(xlpp_headerfooter h, int v);
XLPP_API int  xlpp_headerfooter_different_first(xlpp_headerfooter h);

// ============================================================
// Worksheet protection
// ============================================================
XLPP_API void xlpp_wssprotection_set_enabled(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_enabled(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_password_hash(xlpp_wssprotection p, const char* v);
XLPP_API const char* xlpp_wssprotection_password_hash(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_select_locked(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_select_locked(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_select_unlocked(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_select_unlocked(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_format_cells(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_format_cells(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_format_columns(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_format_columns(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_format_rows(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_format_rows(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_insert_rows(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_insert_rows(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_insert_columns(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_insert_columns(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_delete_rows(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_delete_rows(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_delete_columns(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_delete_columns(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_sort(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_sort(xlpp_wssprotection p);
XLPP_API void xlpp_wssprotection_set_auto_filter(xlpp_wssprotection p, int v);
XLPP_API int  xlpp_wssprotection_auto_filter(xlpp_wssprotection p);

// ============================================================
// SheetView
// ============================================================
XLPP_API void xlpp_sheetview_set_workbook_view_id(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_workbook_view_id(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_tab_color(xlpp_sheetview s, const char* v);
XLPP_API void xlpp_sheetview_clear_tab_color(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_tab_color(xlpp_sheetview s, char* out, int outSize);
XLPP_API void xlpp_sheetview_set_zoom_scale(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_zoom_scale(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_zoom_scale_normal(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_zoom_scale_normal(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_show_grid_lines(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_show_grid_lines(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_tab_selected(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_tab_selected(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_right_to_left(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_right_to_left(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_show_outline_symbols(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_show_outline_symbols(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_pane(xlpp_sheetview s, const char* v);
XLPP_API void xlpp_sheetview_pane(xlpp_sheetview s, char* out, int outSize);
XLPP_API void xlpp_sheetview_set_top_left_cell(xlpp_sheetview s, const char* v);
XLPP_API void xlpp_sheetview_top_left_cell(xlpp_sheetview s, char* out, int outSize);
XLPP_API void xlpp_sheetview_set_x_split(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_x_split(xlpp_sheetview s);
XLPP_API void xlpp_sheetview_set_y_split(xlpp_sheetview s, int v);
XLPP_API int  xlpp_sheetview_y_split(xlpp_sheetview s);

// ============================================================
// Image
// ============================================================
XLPP_API void xlpp_image_set_anchor(xlpp_image img, const char* v);
XLPP_API void xlpp_image_anchor(xlpp_image img, char* out, int outSize);
XLPP_API void xlpp_image_extension(xlpp_image img, char* out, int outSize);
XLPP_API void xlpp_image_set_width(xlpp_image img, double v);
XLPP_API double xlpp_image_width(xlpp_image img);
XLPP_API void xlpp_image_set_height(xlpp_image img, double v);
XLPP_API double xlpp_image_height(xlpp_image img);
XLPP_API void xlpp_image_set_name(xlpp_image img, const char* v);
XLPP_API void xlpp_image_name(xlpp_image img, char* out, int outSize);

// ============================================================
// Table
// ============================================================
XLPP_API void xlpp_table_set_display_name(xlpp_table t, const char* v);
XLPP_API void xlpp_table_name(xlpp_table t, char* out, int outSize);
XLPP_API void xlpp_table_display_name(xlpp_table t, char* out, int outSize);
XLPP_API void xlpp_table_set_reference(xlpp_table t, const char* v);
XLPP_API void xlpp_table_reference(xlpp_table t, char* out, int outSize);
XLPP_API void xlpp_table_set_show_header_row(xlpp_table t, int v);
XLPP_API int  xlpp_table_show_header_row(xlpp_table t);
XLPP_API void xlpp_table_set_show_totals_row(xlpp_table t, int v);
XLPP_API int  xlpp_table_show_totals_row(xlpp_table t);
XLPP_API int  xlpp_table_column_count(xlpp_table t);
XLPP_API xlpp_tablecolumn xlpp_table_column_at(xlpp_table t, int index);
XLPP_API xlpp_tablecolumn xlpp_table_add_column(xlpp_table t, const char* name);
XLPP_API xlpp_tablestyle xlpp_table_style_info(xlpp_table t);

XLPP_API void xlpp_tablecolumn_set_name(xlpp_tablecolumn c, const char* v);
XLPP_API void xlpp_tablecolumn_name(xlpp_tablecolumn c, char* out, int outSize);
XLPP_API uint64_t xlpp_tablecolumn_id(xlpp_tablecolumn c);

XLPP_API void xlpp_tablestyle_set_name(xlpp_tablestyle s, const char* v);
XLPP_API void xlpp_tablestyle_name(xlpp_tablestyle s, char* out, int outSize);
XLPP_API void xlpp_tablestyle_set_show_first(xlpp_tablestyle s, int v);
XLPP_API int  xlpp_tablestyle_show_first(xlpp_tablestyle s);
XLPP_API void xlpp_tablestyle_set_show_last(xlpp_tablestyle s, int v);
XLPP_API int  xlpp_tablestyle_show_last(xlpp_tablestyle s);
XLPP_API void xlpp_tablestyle_set_show_row_stripes(xlpp_tablestyle s, int v);
XLPP_API int  xlpp_tablestyle_show_row_stripes(xlpp_tablestyle s);
XLPP_API void xlpp_tablestyle_set_show_column_stripes(xlpp_tablestyle s, int v);
XLPP_API int  xlpp_tablestyle_show_column_stripes(xlpp_tablestyle s);

// ============================================================
// Chart
// ============================================================
XLPP_API void xlpp_chart_set_grouping(xlpp_chart c, int v);
XLPP_API void xlpp_chart_set_scatter_style(xlpp_chart c, const char* v);
XLPP_API void xlpp_chart_scatter_style(xlpp_chart c, char* out, int outSize);
XLPP_API int  xlpp_chart_add_plot(xlpp_chart c, int type, int grouping, int secondary_axes);
XLPP_API int  xlpp_chart_plot_count(xlpp_chart c);
XLPP_API int  xlpp_chart_plot_type(xlpp_chart c, int plot_index);
XLPP_API int  xlpp_chart_plot_uses_secondary_axes(xlpp_chart c, int plot_index);
XLPP_API xlpp_chartseries xlpp_chart_add_series_to_plot(xlpp_chart c, int plot_index, const char* title);
XLPP_API int  xlpp_chart_grouping(xlpp_chart c);
XLPP_API void xlpp_chart_set_title(xlpp_chart c, const char* v);
XLPP_API void xlpp_chart_title(xlpp_chart c, char* out, int outSize);
XLPP_API void xlpp_chart_set_x_axis_title(xlpp_chart c, const char* v);
XLPP_API void xlpp_chart_set_y_axis_title(xlpp_chart c, const char* v);
XLPP_API void xlpp_chart_set_style(xlpp_chart c, const char* v);
XLPP_API void xlpp_chart_set_width(xlpp_chart c, int v);
XLPP_API int  xlpp_chart_width(xlpp_chart c);
XLPP_API void xlpp_chart_set_height(xlpp_chart c, int v);
XLPP_API int  xlpp_chart_height(xlpp_chart c);
XLPP_API void xlpp_chart_set_show_legend(xlpp_chart c, int v);
XLPP_API int  xlpp_chart_show_legend(xlpp_chart c);
XLPP_API void xlpp_chart_set_legend_position(xlpp_chart c, const char* v);
XLPP_API int  xlpp_chart_series_count(xlpp_chart c);
XLPP_API xlpp_chartseries xlpp_chart_series_at(xlpp_chart c, int index);
XLPP_API xlpp_chartseries xlpp_chart_add_series(xlpp_chart c, const char* title);

XLPP_API void xlpp_chartseries_set_title(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_values_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_categories_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_bubble_size_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_bubble_size_reference(xlpp_chartseries s, char* out, int outSize);

// ============================================================
// Pivot table
// ============================================================
XLPP_API void xlpp_pivottable_set_name(xlpp_pivottable p, const char* v);
XLPP_API void xlpp_pivottable_name(xlpp_pivottable p, char* out, int outSize);
XLPP_API void xlpp_pivottable_set_location(xlpp_pivottable p, const char* v);
XLPP_API void xlpp_pivottable_location(xlpp_pivottable p, char* out, int outSize);
XLPP_API xlpp_pivotcache xlpp_pivottable_cache(xlpp_pivottable p);
XLPP_API void xlpp_pivottable_add_row_field(xlpp_pivottable p, const char* name);
XLPP_API void xlpp_pivottable_add_column_field(xlpp_pivottable p, const char* name);
XLPP_API void xlpp_pivottable_add_page_field(xlpp_pivottable p, const char* name);
XLPP_API void xlpp_pivottable_add_data_field(xlpp_pivottable p);

XLPP_API void xlpp_pivotcache_set_cache_id(xlpp_pivotcache c, int v);
XLPP_API int  xlpp_pivotcache_cache_id(xlpp_pivotcache c);
XLPP_API void xlpp_pivotcache_set_source_data(xlpp_pivotcache c, const char* v);
XLPP_API void xlpp_pivotcache_source_data(xlpp_pivotcache c, char* out, int outSize);

// ============================================================
// Conditional formatting
// ============================================================
XLPP_API int  xlpp_cfcollection_entry_count(xlpp_cfcollection c);
XLPP_API xlpp_cfentry xlpp_cfcollection_add_entry(xlpp_cfcollection c, const char* reference);
XLPP_API xlpp_cfentry xlpp_cfcollection_entry_at(xlpp_cfcollection c, int index);
XLPP_API void xlpp_cfcollection_clear(xlpp_cfcollection c);
XLPP_API int  xlpp_cfcollection_empty(xlpp_cfcollection c);

XLPP_API void xlpp_cfentry_reference(xlpp_cfentry e, char* out, int outSize);
XLPP_API void xlpp_cfentry_set_reference(xlpp_cfentry e, const char* v);
XLPP_API int  xlpp_cfentry_rule_count(xlpp_cfentry e);
XLPP_API xlpp_cfrule xlpp_cfentry_add_rule(xlpp_cfentry e, int type);
XLPP_API xlpp_cfrule xlpp_cfentry_rule_at(xlpp_cfentry e, int index);

XLPP_API void xlpp_cfrule_add_formula(xlpp_cfrule r, const char* f);
XLPP_API int  xlpp_cfrule_formula_count(xlpp_cfrule r);
XLPP_API void xlpp_cfrule_formula_at(xlpp_cfrule r, int index, char* out, int outSize);
XLPP_API void xlpp_cfrule_set_operator(xlpp_cfrule r, int op);
XLPP_API int  xlpp_cfrule_operator(xlpp_cfrule r);
XLPP_API int  xlpp_cfrule_type(xlpp_cfrule r);
XLPP_API void xlpp_cfrule_set_priority(xlpp_cfrule r, uint64_t v);
XLPP_API uint64_t xlpp_cfrule_priority(xlpp_cfrule r);
XLPP_API void xlpp_cfrule_set_stop_if_true(xlpp_cfrule r, int v);
XLPP_API int  xlpp_cfrule_stop_if_true(xlpp_cfrule r);
XLPP_API void xlpp_cfrule_set_differential_style(xlpp_cfrule r, xlpp_style s);

// ============================================================
// Data validation
// ============================================================
XLPP_API int  xlpp_dvcollection_count(xlpp_dvcollection c);
XLPP_API xlpp_datavalidation xlpp_dvcollection_add(xlpp_dvcollection c, int type, const char* reference);
XLPP_API xlpp_datavalidation xlpp_dvcollection_at(xlpp_dvcollection c, int index);
XLPP_API void xlpp_dvcollection_clear(xlpp_dvcollection c);

XLPP_API int  xlpp_datavalidation_type(xlpp_datavalidation d);
XLPP_API void xlpp_datavalidation_set_type(xlpp_datavalidation d, int v);
XLPP_API int  xlpp_datavalidation_operator(xlpp_datavalidation d);
XLPP_API void xlpp_datavalidation_set_operator(xlpp_datavalidation d, int v);
XLPP_API int  xlpp_datavalidation_error_style(xlpp_datavalidation d);
XLPP_API void xlpp_datavalidation_set_error_style(xlpp_datavalidation d, int v);
XLPP_API void xlpp_datavalidation_set_formula1(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_formula1(xlpp_datavalidation d, char* out, int outSize);
XLPP_API void xlpp_datavalidation_set_formula2(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_formula2(xlpp_datavalidation d, char* out, int outSize);
XLPP_API void xlpp_datavalidation_set_reference(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_reference(xlpp_datavalidation d, char* out, int outSize);
XLPP_API void xlpp_datavalidation_set_allow_blank(xlpp_datavalidation d, int v);
XLPP_API int  xlpp_datavalidation_allow_blank(xlpp_datavalidation d);
XLPP_API void xlpp_datavalidation_set_show_drop_down(xlpp_datavalidation d, int v);
XLPP_API int  xlpp_datavalidation_show_drop_down(xlpp_datavalidation d);
XLPP_API void xlpp_datavalidation_set_show_input_message(xlpp_datavalidation d, int v);
XLPP_API void xlpp_datavalidation_set_show_error_message(xlpp_datavalidation d, int v);
XLPP_API void xlpp_datavalidation_set_prompt_title(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_set_prompt(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_set_error_title(xlpp_datavalidation d, const char* v);
XLPP_API void xlpp_datavalidation_set_error(xlpp_datavalidation d, const char* v);

// ============================================================
// Font
// ============================================================
XLPP_API void         xlpp_font_set_name(xlpp_font f, const char* v);
XLPP_API void         xlpp_font_set_size(xlpp_font f, double v);
XLPP_API void         xlpp_font_set_bold(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_italic(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_underline(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_strike(xlpp_font f, int v);
XLPP_API void         xlpp_font_set_color(xlpp_font f, const char* argb);
XLPP_API const char*  xlpp_font_get_name(xlpp_font f);
XLPP_API double       xlpp_font_get_size(xlpp_font f);
XLPP_API int          xlpp_font_get_bold(xlpp_font f);
XLPP_API int          xlpp_font_get_italic(xlpp_font f);
XLPP_API int          xlpp_font_get_underline(xlpp_font f);
XLPP_API int          xlpp_font_get_strike(xlpp_font f);
XLPP_API void         xlpp_font_get_color(xlpp_font f, char* out, int outSize);

// ============================================================
// Fill
// ============================================================
XLPP_API void         xlpp_fill_set_pattern(xlpp_fill f, const char* v);
XLPP_API void         xlpp_fill_set_fg_color(xlpp_fill f, const char* argb);
XLPP_API void         xlpp_fill_set_bg_color(xlpp_fill f, const char* argb);
XLPP_API void         xlpp_fill_get_pattern(xlpp_fill f, char* out, int outSize);
XLPP_API void         xlpp_fill_get_fg_color(xlpp_fill f, char* out, int outSize);
XLPP_API void         xlpp_fill_get_bg_color(xlpp_fill f, char* out, int outSize);

// ============================================================
// Border
// ============================================================
XLPP_API xlpp_borderside xlpp_border_left(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_right(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_top(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_bottom(xlpp_border b);
XLPP_API xlpp_borderside xlpp_border_diagonal(xlpp_border b);
XLPP_API void         xlpp_borderside_set_style(xlpp_borderside s, const char* v);
XLPP_API void         xlpp_borderside_set_color(xlpp_borderside s, const char* argb);
XLPP_API void         xlpp_borderside_get_style(xlpp_borderside s, char* out, int outSize);
XLPP_API void         xlpp_borderside_get_color(xlpp_borderside s, char* out, int outSize);

// ============================================================
// Alignment
// ============================================================
XLPP_API void         xlpp_alignment_set_horizontal(xlpp_alignment a, const char* v);
XLPP_API void         xlpp_alignment_set_vertical(xlpp_alignment a, const char* v);
XLPP_API void         xlpp_alignment_set_wrap_text(xlpp_alignment a, int v);
XLPP_API void         xlpp_alignment_set_shrink_to_fit(xlpp_alignment a, int v);
XLPP_API void         xlpp_alignment_set_text_rotation(xlpp_alignment a, int v);
XLPP_API void         xlpp_alignment_set_indent(xlpp_alignment a, int v);
XLPP_API void         xlpp_alignment_get_horizontal(xlpp_alignment a, char* out, int outSize);
XLPP_API void         xlpp_alignment_get_vertical(xlpp_alignment a, char* out, int outSize);
XLPP_API int          xlpp_alignment_get_wrap_text(xlpp_alignment a);
XLPP_API int          xlpp_alignment_get_shrink_to_fit(xlpp_alignment a);
XLPP_API int          xlpp_alignment_get_text_rotation(xlpp_alignment a);
XLPP_API int          xlpp_alignment_get_indent(xlpp_alignment a);

// ============================================================
// Style
// ============================================================
XLPP_API xlpp_font    xlpp_style_font(xlpp_style s);
XLPP_API xlpp_fill    xlpp_style_fill(xlpp_style s);
XLPP_API xlpp_border  xlpp_style_border(xlpp_style s);
XLPP_API xlpp_alignment xlpp_style_alignment(xlpp_style s);
XLPP_API void         xlpp_style_set_number_format(xlpp_style s, const char* v);
XLPP_API void         xlpp_style_number_format(xlpp_style s, char* out, int outSize);
XLPP_API void         xlpp_style_set_num_fmt_id(xlpp_style s, int v);
XLPP_API int          xlpp_style_num_fmt_id(xlpp_style s);
XLPP_API void         xlpp_style_set_locked(xlpp_style s, int v);
XLPP_API int          xlpp_style_locked(xlpp_style s);
XLPP_API void         xlpp_style_set_hidden(xlpp_style s, int v);
XLPP_API int          xlpp_style_hidden(xlpp_style s);
XLPP_API int          xlpp_style_is_default(xlpp_style s);

// ============================================================
// Streaming writer
// ============================================================
XLPP_API xlpp_stream_writer xlpp_stream_create(const char* path);
XLPP_API void               xlpp_stream_destroy(xlpp_stream_writer w);
XLPP_API uint64_t           xlpp_stream_add_sheet(xlpp_stream_writer w, const char* name);
XLPP_API void               xlpp_stream_append_row(xlpp_stream_writer w, uint64_t sheetIndex, const char** values, int count);
XLPP_API void               xlpp_stream_append_doubles(xlpp_stream_writer w, uint64_t sheetIndex, const double* values, int count);
XLPP_API uint64_t           xlpp_stream_row_count(xlpp_stream_writer w, uint64_t sheetIndex);
XLPP_API uint64_t           xlpp_stream_sheet_count(xlpp_stream_writer w);
XLPP_API void               xlpp_stream_set_date1904(xlpp_stream_writer w, int v);
XLPP_API void               xlpp_stream_set_compression_level(xlpp_stream_writer w, int level);
XLPP_API void               xlpp_stream_set_parallel_workers(xlpp_stream_writer w, uint64_t workers);
XLPP_API void               xlpp_stream_close(xlpp_stream_writer w);

// ============================================================
// Streaming reader
// ============================================================
XLPP_API xlpp_stream_reader xlpp_stream_reader_open(const char* path);
XLPP_API void               xlpp_stream_reader_destroy(xlpp_stream_reader r);
XLPP_API int                xlpp_stream_reader_sheet_count(xlpp_stream_reader r);
XLPP_API void               xlpp_stream_reader_sheet_name(xlpp_stream_reader r, int index, char* out, int outSize);
// Iterate all rows of the given sheet. Callback receives (rowNumber, cellCount,
// addresses, values, formula flags, styleIndexes) as raw arrays; return 0 to stop.
typedef int (*xlpp_stream_row_callback)(void* user, uint64_t rowNumber, int cellCount,
                                        const char** addresses, const double* numbers,
                                        const int* valueTypes, const char** strings,
                                        const char** formulas, const int* styleIndexes);
XLPP_API int xlpp_stream_reader_read_sheet(xlpp_stream_reader r, int index,
                                           xlpp_stream_row_callback callback, void* user);

// ============================================================
// Utility
// ============================================================
XLPP_API const char*  xlpp_version(void);
XLPP_API void         xlpp_free_string(const char* str);
XLPP_API const char*  xlpp_last_error(void);
XLPP_API void         xlpp_clear_error(void);

#ifdef __cplusplus
}
#endif
