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
//
// Lifetime rule: handles to surviving child model objects remain valid when the
// owning collection grows or when another element is erased. A handle becomes
// invalid when its own element is erased, its owning model is destroyed/replaced,
// or an API explicitly replaces the containing model.
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
typedef struct xlpp_dependency_graph_t* xlpp_dependency_graph;
typedef struct xlpp_validation_report_t* xlpp_validation_report;
typedef struct xlpp_external_value_t* xlpp_external_value;

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
#define XLPP_ERROR_SPILL            8
#define XLPP_ERROR_CALC             9

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

// SharedStringMode
#define XLPP_SSM_DISABLED   0
#define XLPP_SSM_HASH       1
#define XLPP_SSM_BOUNDED_LRU 2

// Formula / structural editing reports
typedef int (*xlpp_external_reference_resolver)(void* user, const char* workbook_token, const char* sheet_name,
                                                const char* address, xlpp_external_value output);

typedef struct xlpp_calculation_options_t {
    int recursive_dependencies;
    int update_cached_values;
    int evaluate_volatile_functions;
    int spill_dynamic_arrays;
    int iterative_calculation;
    uint64_t max_iterations;
    double max_change;
    uint64_t max_depth;
    xlpp_external_reference_resolver external_reference_resolver;
    void* external_reference_user;
} xlpp_calculation_options;

typedef struct xlpp_calculation_report_t {
    uint64_t formula_cells_visited;
    uint64_t formula_cells_evaluated;
    uint64_t cached_values_updated;
    uint64_t dependency_evaluations;
    uint64_t defined_names_resolved;
    uint64_t circular_references;
    uint64_t unsupported_formulas;
    uint64_t evaluation_errors;
    uint64_t dynamic_arrays_spilled;
    uint64_t spill_cells_updated;
    uint64_t spill_conflicts;
    uint64_t structured_references_resolved;
    uint64_t iterative_iterations;
    uint64_t iterative_convergence_failures;
    uint64_t external_references_resolved;
    uint64_t unresolved_external_references;
    int success;
} xlpp_calculation_report;

typedef struct xlpp_structural_report_t {
    uint64_t worksheets_visited;
    uint64_t cells_moved;
    uint64_t cells_removed;
    uint64_t formulas_updated;
    uint64_t formula_metadata_updated;
    uint64_t worksheet_references_updated;
    uint64_t defined_names_updated;
    uint64_t chart_references_updated;
    uint64_t pivot_references_updated;
    uint64_t drawing_anchors_updated;
    uint64_t hyperlinks_updated;
    uint64_t references_invalidated;
    uint64_t formulas_calculated;
    uint64_t chart_caches_updated;
    int success;
} xlpp_structural_report;

#define XLPP_STRUCT_INSERT_ROWS     0
#define XLPP_STRUCT_DELETE_ROWS     1
#define XLPP_STRUCT_INSERT_COLUMNS  2
#define XLPP_STRUCT_DELETE_COLUMNS  3

#define XLPP_ENCRYPTION_NONE                 0
#define XLPP_ENCRYPTION_AGILE_AES256_SHA512 1
#define XLPP_ENCRYPTION_STANDARD_AES_SHA1    2
#define XLPP_ENCRYPTION_UNSUPPORTED          3

#define XLPP_DEPENDENCY_CELL_OR_RANGE       0
#define XLPP_DEPENDENCY_DEFINED_NAME        1
#define XLPP_DEPENDENCY_TABLE               2
#define XLPP_DEPENDENCY_EXTERNAL_REFERENCE  3
#define XLPP_DEPENDENCY_VOLATILE_REFERENCE  4

#define XLPP_VALIDATION_WARNING 0
#define XLPP_VALIDATION_ERROR   1

#define XLPP_VBA_MODULE_STANDARD 0
#define XLPP_VBA_MODULE_DOCUMENT 1
#define XLPP_VBA_MODULE_CLASS    2

typedef int (*xlpp_cancel_callback)(void* user);
typedef void (*xlpp_progress_callback)(void* user, uint64_t done, uint64_t total);

typedef struct xlpp_load_options_t {
    int lenient;
    uint64_t max_entries;
    uint64_t max_entry_bytes;
    uint64_t max_total_bytes;
    uint64_t max_file_bytes;
    const char* password;
    int verify_encryption_integrity;
    xlpp_cancel_callback cancel;
    xlpp_progress_callback progress;
    void* callback_user;
} xlpp_load_options;

typedef struct xlpp_save_options_t {
    int compression_level;
    int compression_strategy;
    uint64_t parallel_workers;
    int parallel_sheets;
    int parallel_rows;
    int strict_namespace;
    int synchronize_chart_caches;
    int synchronize_changed_chart_caches_only;
    int calculate_formulas_before_save;
    int atomic_write;
    int validate_before_save;
    const char* encryption_password;
    int encryption_mode;
    uint64_t encryption_spin_count;
    uint64_t encryption_key_bits;
} xlpp_save_options;

typedef struct xlpp_structural_options_t {
    int transactional;
    int update_defined_names;
    int recalculate_formulas;
    int synchronize_chart_caches;
    int changed_chart_caches_only;
    int fail_on_invalid_reference;
} xlpp_structural_options;

typedef struct xlpp_worksheet_rename_report_t {
    uint64_t worksheets_visited;
    uint64_t formulas_updated;
    uint64_t formula_metadata_updated;
    uint64_t defined_names_updated;
    uint64_t chart_references_updated;
    uint64_t pivot_references_updated;
    uint64_t hyperlinks_updated;
    uint64_t references_updated;
    uint64_t formulas_calculated;
    uint64_t chart_caches_updated;
    int success;
} xlpp_worksheet_rename_report;

typedef struct xlpp_chart_cache_sync_options_t {
    int synchronize_titles;
    int synchronize_categories;
    int synchronize_values;
    int changed_references_only;
    int clear_unsupported_references;
} xlpp_chart_cache_sync_options;

typedef struct xlpp_chart_cache_sync_report_t {
    uint64_t charts_visited;
    uint64_t series_visited;
    uint64_t references_checked;
    uint64_t references_unchanged;
    uint64_t dependencies_registered;
    uint64_t dependencies_changed;
    uint64_t caches_updated;
    uint64_t caches_cleared;
    uint64_t references_skipped;
    int success;
} xlpp_chart_cache_sync_report;

typedef struct xlpp_dependency_report_t {
    uint64_t formula_cells;
    uint64_t edges;
    uint64_t cell_or_range_edges;
    uint64_t defined_name_edges;
    uint64_t table_edges;
    uint64_t external_edges;
    uint64_t volatile_references;
    uint64_t unresolved_symbols;
} xlpp_dependency_report;

typedef struct xlpp_validation_options_t {
    int validate_worksheet_names;
    int validate_defined_names;
    int validate_tables;
    int validate_pivots;
} xlpp_validation_options;

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
XLPP_API xlpp_worksheet xlpp_workbook_copy_sheet(xlpp_workbook wb, xlpp_worksheet src, const char* new_name);
XLPP_API int            xlpp_workbook_sheet_index(xlpp_workbook wb, xlpp_worksheet ws);
XLPP_API const char*    xlpp_workbook_sheet_name(xlpp_workbook wb, int index, char* out, int outSize);
XLPP_API int            xlpp_workbook_sheet_names_count(xlpp_workbook wb);

XLPP_API int  xlpp_workbook_load(xlpp_workbook wb, const char* path);
XLPP_API int  xlpp_workbook_save(xlpp_workbook wb, const char* path);
XLPP_API int  xlpp_workbook_save_durable(xlpp_workbook wb, const char* path, int durable_write);
XLPP_API int  xlpp_workbook_load_password(xlpp_workbook wb, const char* path, const char* password, int verify_integrity);
XLPP_API int  xlpp_workbook_save_encrypted(xlpp_workbook wb, const char* path, const char* password, uint64_t spin_count, int calculate_formulas);
XLPP_API int  xlpp_workbook_save_encrypted_ex(xlpp_workbook wb, const char* path, const char* password, int encryption_mode, uint64_t key_bits, uint64_t spin_count, int calculate_formulas);
XLPP_API int  xlpp_workbook_calculate(xlpp_workbook wb, xlpp_calculation_report* report);
XLPP_API int  xlpp_workbook_calculate_ex(xlpp_workbook wb, int iterative, uint64_t max_iterations, double max_change, xlpp_calculation_report* report);
XLPP_API int  xlpp_workbook_calculate_options(xlpp_workbook wb, const xlpp_calculation_options* options, xlpp_calculation_report* report);
XLPP_API void xlpp_external_value_set_empty(xlpp_external_value value);
XLPP_API void xlpp_external_value_set_number(xlpp_external_value value, double number);
XLPP_API void xlpp_external_value_set_bool(xlpp_external_value value, int boolean_value);
XLPP_API void xlpp_external_value_set_string(xlpp_external_value value, const char* string_value);
XLPP_API void xlpp_external_value_set_error(xlpp_external_value value, int error_code);
XLPP_API void xlpp_external_value_set_date(xlpp_external_value value, int year, int month, int day, int hour, int minute, double second, int has_time);
XLPP_API int  xlpp_workbook_structural_edit(xlpp_workbook wb, const char* sheet_name, int kind, uint64_t index, uint64_t amount, int fail_on_invalid_reference, xlpp_structural_report* report);
XLPP_API int  xlpp_workbook_load_ex(xlpp_workbook wb, const char* path, const xlpp_load_options* options);
XLPP_API int  xlpp_workbook_save_ex(xlpp_workbook wb, const char* path, const xlpp_save_options* options);
XLPP_API int  xlpp_workbook_load_bytes(xlpp_workbook wb, const unsigned char* bytes, uint64_t size, const xlpp_load_options* options);
XLPP_API int  xlpp_workbook_save_bytes(xlpp_workbook wb, const xlpp_save_options* options, unsigned char** bytes, uint64_t* size);
XLPP_API void xlpp_free_bytes(unsigned char* bytes);
XLPP_API int  xlpp_workbook_structural_edit_ex(xlpp_workbook wb, const char* sheet_name, int kind, uint64_t index, uint64_t amount, const xlpp_structural_options* options, xlpp_structural_report* report);
XLPP_API int  xlpp_workbook_rename_sheet(xlpp_workbook wb, const char* old_name, const char* new_name, int recalculate_formulas, int synchronize_chart_caches, int changed_chart_caches_only, xlpp_worksheet_rename_report* report);
XLPP_API int  xlpp_workbook_synchronize_chart_caches(xlpp_workbook wb, const xlpp_chart_cache_sync_options* options, xlpp_chart_cache_sync_report* report);
XLPP_API void xlpp_workbook_reset_chart_cache_tracking(xlpp_workbook wb);
XLPP_API uint64_t xlpp_workbook_tracked_chart_cache_dependencies(xlpp_workbook wb);

XLPP_API xlpp_dependency_graph xlpp_workbook_dependency_graph(xlpp_workbook wb);
XLPP_API void xlpp_dependency_graph_destroy(xlpp_dependency_graph graph);
XLPP_API int xlpp_dependency_graph_report(xlpp_dependency_graph graph, xlpp_dependency_report* report);
XLPP_API uint64_t xlpp_dependency_graph_edge_count(xlpp_dependency_graph graph);
XLPP_API int xlpp_dependency_graph_edge_kind(xlpp_dependency_graph graph, uint64_t index);
XLPP_API int xlpp_dependency_graph_edge_dependent_sheet(xlpp_dependency_graph graph, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_dependency_graph_edge_dependent_cell(xlpp_dependency_graph graph, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_dependency_graph_edge_precedent_sheet(xlpp_dependency_graph graph, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_dependency_graph_edge_precedent_reference(xlpp_dependency_graph graph, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_dependency_graph_edge_symbol(xlpp_dependency_graph graph, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_dependency_graph_depends_on(xlpp_dependency_graph graph, const char* dependent_sheet, const char* dependent_cell, const char* precedent_sheet, const char* precedent_cell);

XLPP_API xlpp_validation_report xlpp_workbook_validate(xlpp_workbook wb, const xlpp_validation_options* options);
XLPP_API void xlpp_validation_report_destroy(xlpp_validation_report report);
XLPP_API uint64_t xlpp_validation_error_count(xlpp_validation_report report);
XLPP_API uint64_t xlpp_validation_warning_count(xlpp_validation_report report);
XLPP_API uint64_t xlpp_validation_issue_count(xlpp_validation_report report);
XLPP_API int xlpp_validation_issue_severity(xlpp_validation_report report, uint64_t index);
XLPP_API int xlpp_validation_issue_code(xlpp_validation_report report, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_validation_issue_message(xlpp_validation_report report, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_validation_issue_worksheet(xlpp_validation_report report, uint64_t index, char* out, int out_size);

XLPP_API int xlpp_workbook_add_vba_project(xlpp_workbook wb, const char* path);
XLPP_API int xlpp_workbook_set_vba_project(xlpp_workbook wb, const unsigned char* bytes, uint64_t size);
XLPP_API int xlpp_workbook_has_vba_project(xlpp_workbook wb);
XLPP_API int xlpp_workbook_remove_vba_project(xlpp_workbook wb);
XLPP_API int xlpp_workbook_set_vba_module_text(xlpp_workbook wb, const char* module_name, const char* source);
XLPP_API int xlpp_workbook_set_vba_class_module_text(xlpp_workbook wb, const char* module_name, const char* source, int read_only, int private_module);
XLPP_API int xlpp_workbook_set_vba_document_module_text(xlpp_workbook wb, const char* module_name, const char* source);
XLPP_API int xlpp_workbook_vba_module_text(xlpp_workbook wb, const char* module_name, char* out, int out_size);
XLPP_API uint64_t xlpp_workbook_vba_module_count(xlpp_workbook wb);
XLPP_API int xlpp_workbook_vba_module_type(xlpp_workbook wb, uint64_t index);
XLPP_API int xlpp_workbook_vba_module_read_only(xlpp_workbook wb, uint64_t index);
XLPP_API int xlpp_workbook_vba_module_private(xlpp_workbook wb, uint64_t index);
XLPP_API int xlpp_workbook_vba_module_name(xlpp_workbook wb, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_workbook_vba_module_source(xlpp_workbook wb, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_workbook_remove_vba_module(xlpp_workbook wb, const char* module_name);
XLPP_API uint64_t xlpp_workbook_vba_project_bytes(xlpp_workbook wb, unsigned char* out, uint64_t out_size);
XLPP_API int xlpp_workbook_save_vba_project(xlpp_workbook wb, const char* path);
XLPP_API int xlpp_workbook_has_vba_signature(xlpp_workbook wb);
XLPP_API int xlpp_workbook_vba_source_editable(xlpp_workbook wb);
XLPP_API int xlpp_workbook_vba_project_name(xlpp_workbook wb, char* out, int out_size);
XLPP_API int xlpp_workbook_vba_project_description(xlpp_workbook wb, char* out, int out_size);
XLPP_API int xlpp_workbook_vba_project_help_file(xlpp_workbook wb, char* out, int out_size);
XLPP_API uint32_t xlpp_workbook_vba_project_help_context(xlpp_workbook wb);
XLPP_API int xlpp_workbook_vba_project_constants(xlpp_workbook wb, char* out, int out_size);
XLPP_API int xlpp_workbook_set_vba_project_properties(xlpp_workbook wb, const char* name, const char* description, const char* help_file, uint32_t help_context, const char* constants);
XLPP_API int  xlpp_inspect_office_encryption(const char* path, int* mode, uint64_t* key_bits, uint64_t* spin_count, char* cipher, int cipher_size, char* hash, int hash_size);
XLPP_API void xlpp_workbook_set_date1904(xlpp_workbook wb, int v);
XLPP_API int  xlpp_workbook_date1904(xlpp_workbook wb);
XLPP_API void xlpp_workbook_clear(xlpp_workbook wb);
XLPP_API int  xlpp_workbook_strict_namespaces(xlpp_workbook wb);
XLPP_API uint64_t xlpp_workbook_diagnostic_warning_count(xlpp_workbook wb);
XLPP_API uint64_t xlpp_workbook_diagnostic_error_count(xlpp_workbook wb);
XLPP_API int xlpp_workbook_diagnostic_warning(xlpp_workbook wb, uint64_t index, char* out, int out_size);
XLPP_API int xlpp_workbook_diagnostic_error(xlpp_workbook wb, uint64_t index, char* out, int out_size);

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
XLPP_API xlpp_definedname xlpp_workbook_add_defined_name_scoped(xlpp_workbook wb, const char* name, const char* value, uint64_t local_sheet_id, int* ok);
XLPP_API xlpp_definedname xlpp_workbook_defined_name(xlpp_workbook wb, const char* name);
XLPP_API xlpp_definedname xlpp_workbook_defined_name_scoped(xlpp_workbook wb, const char* name, int has_local_sheet_id, uint64_t local_sheet_id);
XLPP_API int              xlpp_workbook_defined_names_count(xlpp_workbook wb);
XLPP_API xlpp_definedname xlpp_workbook_defined_name_at(xlpp_workbook wb, int index);
XLPP_API int xlpp_workbook_preserved_relationships_count(xlpp_workbook wb);
XLPP_API int xlpp_workbook_preserved_relationship_source_part(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_relationship_id(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_relationship_type(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_relationship_target(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_relationship_target_mode(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_parts_count(xlpp_workbook wb);
XLPP_API int xlpp_workbook_preserved_part_name(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API uint64_t xlpp_workbook_preserved_part_data_size(xlpp_workbook wb, int index);
XLPP_API uint64_t xlpp_workbook_preserved_part_data(xlpp_workbook wb, int index, unsigned char* out, uint64_t out_size);
XLPP_API int xlpp_workbook_preserved_part_override_type(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_part_extension(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_part_default_type(xlpp_workbook wb, int index, char* out, int out_size);
XLPP_API int xlpp_workbook_preserved_part_compress(xlpp_workbook wb, int index);

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
XLPP_API int          xlpp_sheet_vba_code_name(xlpp_worksheet ws, char* out, int out_size);
XLPP_API int          xlpp_sheet_set_vba_code_name(xlpp_worksheet ws, const char* code_name);

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
XLPP_API uint64_t xlpp_definedname_local_sheet_id(xlpp_definedname d, int* has_value);
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
XLPP_API int  xlpp_chart_type(xlpp_chart c);
XLPP_API int  xlpp_chart_is_modern(xlpp_chart c);
XLPP_API int  xlpp_chart_plot_count(xlpp_chart c);
XLPP_API int  xlpp_chart_add_plot(xlpp_chart c, int type, uint64_t first_series, uint64_t series_count, int secondary_axes);
XLPP_API void xlpp_chart_plot_set_grouping(xlpp_chart c, int plot_index, int grouping);
XLPP_API void xlpp_chart_plot_set_bar_direction(xlpp_chart c, int plot_index, int direction);
XLPP_API void xlpp_chart_plot_set_scatter_style(xlpp_chart c, int plot_index, int style);
XLPP_API void xlpp_chart_plot_set_bubble_options(xlpp_chart c, int plot_index, int scale, int show_negative, int size_represents, int bubble3d);
XLPP_API void xlpp_chart_plot_set_histogram_bins(xlpp_chart c, int plot_index, double bin_width, int bin_count, int automatic_bins);
XLPP_API void xlpp_chart_plot_set_histogram_bounds(xlpp_chart c, int plot_index, int has_underflow, double underflow, int has_overflow, double overflow);
XLPP_API void xlpp_chart_plot_set_box_whisker_options(xlpp_chart c, int plot_index, int inner_points, int outliers, int mean_line, int mean_marker, int quartile_inclusive);
XLPP_API void xlpp_chart_plot_set_waterfall_connector_lines(xlpp_chart c, int plot_index, int enabled);

XLPP_API void xlpp_chartseries_set_title(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_values_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_categories_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_bubble_size_reference(xlpp_chartseries s, const char* v);
XLPP_API void xlpp_chartseries_set_smooth(xlpp_chartseries s, int v);
XLPP_API void xlpp_chartseries_clear_smooth(xlpp_chartseries s);

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
XLPP_API void xlpp_pivottable_add_data_field_named(xlpp_pivottable p, const char* name, const char* subtotal);
XLPP_API void xlpp_pivottable_set_layout(xlpp_pivottable p, int layout);
XLPP_API int  xlpp_pivottable_layout(xlpp_pivottable p);
XLPP_API void xlpp_pivottable_set_flags(xlpp_pivottable p, int show_empty_row, int show_empty_column, int show_drill, int enable_drill, int multiple_field_filters, int show_values_row, int subtotal_hidden_items);
XLPP_API void xlpp_pivottable_set_page_layout(xlpp_pivottable p, int page_wrap, int over_then_down);
XLPP_API void xlpp_pivottable_set_style(xlpp_pivottable p, const char* style_name, const char* data_caption);
XLPP_API void xlpp_pivottable_set_display_options(xlpp_pivottable p, int row_grand_totals, int column_grand_totals, int preserve_formatting, int use_auto_formatting, int show_row_headers, int show_column_headers, int show_row_stripes, int show_column_stripes, int show_last_column);
XLPP_API int  xlpp_pivottable_data_field_count(xlpp_pivottable p);
XLPP_API void xlpp_pivottable_configure_data_field(xlpp_pivottable p, int index, const char* subtotal, const char* caption, uint32_t number_format_id, const char* show_data_as, int base_field, int base_item);
XLPP_API int  xlpp_pivottable_row_field_count(xlpp_pivottable p);
XLPP_API int  xlpp_pivottable_column_field_count(xlpp_pivottable p);
XLPP_API int  xlpp_pivottable_page_field_count(xlpp_pivottable p);
XLPP_API xlpp_pivotfield xlpp_pivottable_row_field_at(xlpp_pivottable p, int index);
XLPP_API xlpp_pivotfield xlpp_pivottable_column_field_at(xlpp_pivottable p, int index);
XLPP_API xlpp_pivotfield xlpp_pivottable_page_field_at(xlpp_pivottable p, int index);
XLPP_API void xlpp_pivottable_add_filter(xlpp_pivottable p, const char* type, int field_index, int measure_field_index, const char* value1, const char* value2, double top10_value, int top10_percent, int top10_top);

XLPP_API void xlpp_pivotcache_set_cache_id(xlpp_pivotcache c, int v);
XLPP_API int  xlpp_pivotcache_cache_id(xlpp_pivotcache c);
XLPP_API void xlpp_pivotcache_set_source_data(xlpp_pivotcache c, const char* v);
XLPP_API void xlpp_pivotcache_source_data(xlpp_pivotcache c, char* out, int outSize);
XLPP_API void xlpp_pivotcache_set_refresh_on_load(xlpp_pivotcache c, int v);
XLPP_API void xlpp_pivotcache_set_save_data(xlpp_pivotcache c, int v);
XLPP_API void xlpp_pivotcache_set_enable_refresh(xlpp_pivotcache c, int v);
XLPP_API void xlpp_pivotcache_set_missing_items_limit(xlpp_pivotcache c, int v);
XLPP_API void xlpp_pivotcache_set_advanced_flags(xlpp_pivotcache c, int background_query, int optimize_memory, int upgrade_on_refresh, int support_subquery, int support_advanced_drill);
XLPP_API void xlpp_pivotcache_set_refreshed_by(xlpp_pivotcache c, const char* v);

XLPP_API void xlpp_pivotfield_set_repeat_item_labels(xlpp_pivotfield f, int v);
XLPP_API void xlpp_pivotfield_set_compact(xlpp_pivotfield f, int v);
XLPP_API void xlpp_pivotfield_set_outline(xlpp_pivotfield f, int v);
XLPP_API void xlpp_pivotfield_set_show_dropdowns(xlpp_pivotfield f, int v);
XLPP_API void xlpp_pivotfield_set_behavior(xlpp_pivotfield f, int show_all, int sort_type, int subtotal_top, int insert_blank_row, int include_new_items_in_filter, int multiple_item_selection_allowed, int selected_item_index, int insert_page_break, int default_subtotal);
XLPP_API void xlpp_pivotfield_add_subtotal(xlpp_pivotfield f, const char* subtotal);
XLPP_API void xlpp_pivotfield_set_item_hidden(xlpp_pivotfield f, int item_index, int hidden);
XLPP_API void xlpp_pivotfield_set_numeric_grouping(xlpp_pivotfield f, int auto_start, int auto_end, double start, double end, double interval);
XLPP_API void xlpp_pivotfield_set_date_grouping(xlpp_pivotfield f, int date_part, int auto_start, int auto_end, const char* start_date, const char* end_date);
XLPP_API void xlpp_pivotfield_clear_grouping(xlpp_pivotfield f);

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
XLPP_API uint64_t     xlpp_c_abi_version(void);
// Stable feature bits for runtime capability negotiation across FFI boundaries.
// Additive only: existing bits never change meaning.
enum xlpp_capability_bits {
    XLPP_CAP_FORMULA_ENGINE = 1ull << 0,
    XLPP_CAP_STREAMING = 1ull << 1,
    XLPP_CAP_ENCRYPTION = 1ull << 2,
    XLPP_CAP_CHARTS = 1ull << 3,
    XLPP_CAP_PIVOT = 1ull << 4,
    XLPP_CAP_VBA = 1ull << 5,
    XLPP_CAP_EXTERNAL_DATA_INSPECTION = 1ull << 6,
    XLPP_CAP_DATA_MODEL_INSPECTION = 1ull << 7,
    XLPP_CAP_DIRTY_RECALC = 1ull << 8,
    XLPP_CAP_ADVANCED_AUTOFILTER = 1ull << 9
};
XLPP_API uint64_t     xlpp_capabilities(void);

typedef struct xlpp_external_data_summary {
    uint64_t external_workbooks;
    uint64_t connections;
    uint64_t query_tables;
    uint64_t power_query_parts;
    uint64_t web_query_parts;
    uint64_t unknown_connection_parts;
} xlpp_external_data_summary;

typedef struct xlpp_data_model_summary {
    int present;
    int has_olap_pivot_caches;
    uint64_t model_parts;
    uint64_t model_relationships;
    uint64_t olap_pivot_cache_parts;
    uint64_t warnings;
} xlpp_data_model_summary;

XLPP_API int xlpp_workbook_inspect_external_data(xlpp_workbook wb, xlpp_external_data_summary* out);
XLPP_API int xlpp_workbook_inspect_data_model(xlpp_workbook wb, xlpp_data_model_summary* out);
XLPP_API void         xlpp_free_string(const char* str);
XLPP_API const char*  xlpp_last_error(void);
XLPP_API void         xlpp_clear_error(void);

#ifdef __cplusplus
}
#endif
