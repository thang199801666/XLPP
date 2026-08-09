#include "xlpp_capi.h"
#include <cstring>
#include <cstdio>
#include <cmath>
#include <string>

// CI-friendly checks: return 1 (nonzero) on the first failure instead of
// aborting, so the failure is visible in the workflow log.
#define CHECK(cond, msg) do { if (!(cond)) { fprintf(stderr, "FAIL: %s\n", msg); return 1; } } while (0)
#define CHECK_STR(a, b, msg) do { if (std::strcmp((a), (b)) != 0) { fprintf(stderr, "FAIL: %s (got '%s')\n", msg, (a)); return 1; } } while (0)

struct LoadCallbackState { uint64_t calls{0}; uint64_t lastDone{0}; uint64_t lastTotal{0}; };
static int neverCancel(void*) { return 0; }
static void progressCallback(void* user, uint64_t done, uint64_t total) {
    auto* state = static_cast<LoadCallbackState*>(user);
    if (!state) return;
    ++state->calls; state->lastDone = done; state->lastTotal = total;
}

static int externalReferenceResolver(void*, const char* workbookToken, const char* sheetName,
                                     const char* address, xlpp_external_value output) {
    if (!workbookToken || !sheetName || !address || !output) return 0;
    if (std::strcmp(workbookToken, "External.xlsx") != 0 || std::strcmp(sheetName, "Data") != 0) return 0;
    if (std::strcmp(address, "A1") == 0) { xlpp_external_value_set_number(output, 41.0); return 1; }
    if (std::strcmp(address, "A2") == 0) { xlpp_external_value_set_string(output, "resolved"); return 1; }
    return 0;
}

int main() {
    CHECK_STR(xlpp_version(), "1.12.0", "version");
    CHECK(xlpp_c_abi_version() >= 2, "c_abi_version");
    const auto capabilities = xlpp_capabilities();
    CHECK((capabilities & XLPP_CAP_DIRTY_RECALC) != 0, "capability_dirty_recalc");
    CHECK((capabilities & XLPP_CAP_ADVANCED_AUTOFILTER) != 0, "capability_advanced_autofilter");
    CHECK((capabilities & XLPP_CAP_EXTERNAL_DATA_INSPECTION) != 0, "capability_external_data");
    CHECK((capabilities & XLPP_CAP_DATA_MODEL_INSPECTION) != 0, "capability_data_model");

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

    // Scoped defined names preserve Excel workbook/local-name semantics.
    auto otherSheet = xlpp_workbook_add_sheet(wb2, "Other");
    CHECK(otherSheet != nullptr, "add_other_sheet");
    int scopedOk = 0;
    auto local0 = xlpp_workbook_add_defined_name_scoped(wb2, "Rate", "Smoke!$A$1", 0, &scopedOk);
    CHECK(scopedOk == 1 && local0 != nullptr, "add_scoped_defined_name_0");
    auto local1 = xlpp_workbook_add_defined_name_scoped(wb2, "Rate", "Other!$A$1", 1, &scopedOk);
    CHECK(scopedOk == 1 && local1 != nullptr, "add_scoped_defined_name_1");
    local0 = xlpp_workbook_defined_name_scoped(wb2, "Rate", 1, 0);
    local1 = xlpp_workbook_defined_name_scoped(wb2, "Rate", 1, 1);
    CHECK(local0 != nullptr && std::strcmp(xlpp_definedname_value(local0), "Smoke!$A$1") == 0, "lookup_scoped_defined_name_0");
    CHECK(local1 != nullptr && std::strcmp(xlpp_definedname_value(local1), "Other!$A$1") == 0, "lookup_scoped_defined_name_1");
    int hasLocalId = 0;
    CHECK(xlpp_definedname_local_sheet_id(local1, &hasLocalId) == 1 && hasLocalId == 1, "defined_name_local_sheet_id");

    // P0W/P1 parity: formula calculation through the C ABI.
    xlpp_cell_set_number(xlpp_sheet_cell(s2, "G1"), 10.0);
    xlpp_cell_set_number(xlpp_sheet_cell(s2, "G2"), 5.0);
    auto calcCell = xlpp_sheet_cell(s2, "G3");
    xlpp_cell_set_formula(calcCell, "=SUM(G1:G2)");
    xlpp_calculation_report calcReport{};
    CHECK(xlpp_workbook_calculate(wb2, &calcReport) == 1, "calculate_formulas");
    CHECK(calcReport.formula_cells_evaluated >= 1, "calculation_report");
    CHECK(std::fabs(xlpp_cell_get_number(calcCell) - 15.0) < 1e-9, "calculated_cache");

    auto iterA = xlpp_sheet_cell(s2, "H1"); auto iterB = xlpp_sheet_cell(s2, "H2");
    xlpp_cell_set_formula(iterA, "(H2+1)/2"); xlpp_cell_set_number(iterA, 0.0);
    xlpp_cell_set_formula(iterB, "(H1+1)/2"); xlpp_cell_set_number(iterB, 0.0);
    xlpp_calculation_report iterativeReport{};
    CHECK(xlpp_workbook_calculate_ex(wb2, 1, 100, 1e-9, &iterativeReport) == 1, "calculate_iterative");
    CHECK(iterativeReport.iterative_convergence_failures == 0 && iterativeReport.iterative_iterations > 1, "iterative_report");
    CHECK(std::fabs(xlpp_cell_get_number(iterA) - 1.0) < 1e-7, "iterative_result");

    // Workbook-level structural transaction updates formula dependencies.
    xlpp_structural_report structuralReport{};
    CHECK(xlpp_workbook_structural_edit(wb2, "Smoke", XLPP_STRUCT_INSERT_ROWS, 2, 1, 0, &structuralReport) == 1, "structural_insert_rows");
    CHECK(structuralReport.cells_moved >= 1, "structural_report");
    auto movedFormula = xlpp_sheet_cell(s2, "G4");
    CHECK_STR(xlpp_cell_get_formula(movedFormula), "=SUM(G1:G3)", "structural_formula_translation");

    // P0Z-D binding parity: full calculation/structural options and workbook services.
    xlpp_calculation_options fullCalc{};
    fullCalc.recursive_dependencies = 1; fullCalc.update_cached_values = 1; fullCalc.evaluate_volatile_functions = 1;
    fullCalc.spill_dynamic_arrays = 1; fullCalc.max_iterations = 50; fullCalc.max_change = 1e-8; fullCalc.max_depth = 256;
    fullCalc.external_reference_resolver = externalReferenceResolver;
    xlpp_calculation_report fullCalcReport{};
    auto externalFormula = xlpp_sheet_cell(s2, "J1");
    xlpp_cell_set_formula(externalFormula, "='[External.xlsx]Data'!A1+1");
    CHECK(xlpp_workbook_calculate_options(wb2, &fullCalc, &fullCalcReport) == 1, "calculate_options");
    CHECK(std::fabs(xlpp_cell_get_number(externalFormula) - 42.0) < 1e-9, "external_reference_callback");
    CHECK(fullCalcReport.external_references_resolved >= 1, "external_reference_report");

    auto refs = xlpp_workbook_add_sheet(wb2, "Refs");
    CHECK(refs != nullptr, "add_refs_sheet");
    auto refCell = xlpp_sheet_cell(refs, "A1");
    xlpp_cell_set_formula(refCell, "=Smoke!G4");
    auto graph = xlpp_workbook_dependency_graph(wb2);
    CHECK(graph != nullptr, "dependency_graph");
    xlpp_dependency_report depReport{};
    CHECK(xlpp_dependency_graph_report(graph, &depReport) == 1 && depReport.edges >= 1, "dependency_report");
    CHECK(xlpp_dependency_graph_depends_on(graph, "Refs", "A1", "Smoke", "G4") == 1, "dependency_query");
    CHECK(xlpp_dependency_graph_edge_count(graph) >= 1, "dependency_edges");
    xlpp_dependency_graph_destroy(graph);

    xlpp_worksheet_rename_report renameReport{};
    CHECK(xlpp_workbook_rename_sheet(wb2, "Smoke", "Renamed", 0, 1, 0, &renameReport) == 1, "dependency_aware_rename");
    CHECK(renameReport.formulas_updated >= 1, "rename_report");
    CHECK_STR(xlpp_cell_get_formula(refCell), "='Renamed'!G4", "rename_formula_translation");

    xlpp_chart_cache_sync_options cacheOptions{1,1,1,1,0};
    xlpp_chart_cache_sync_report cacheReport{};
    CHECK(xlpp_workbook_synchronize_chart_caches(wb2, &cacheOptions, &cacheReport) == 1, "chart_cache_sync");
    xlpp_workbook_reset_chart_cache_tracking(wb2);
    CHECK(xlpp_workbook_tracked_chart_cache_dependencies(wb2) == 0, "chart_cache_reset");

    xlpp_validation_options validationOptions{1,1,1,1};
    auto validation = xlpp_workbook_validate(wb2, &validationOptions);
    CHECK(validation != nullptr, "workbook_validate");
    CHECK(xlpp_validation_error_count(validation) == 0, "validation_errors");
    xlpp_validation_report_destroy(validation);

    CHECK(xlpp_workbook_diagnostic_error_count(wb2) == 0, "load_diagnostics_errors");
    xlpp_external_data_summary externalDataSummary{};
    CHECK(xlpp_workbook_inspect_external_data(wb2, &externalDataSummary) == 1, "external_data_inspection");
    xlpp_data_model_summary dataModelSummary{};
    CHECK(xlpp_workbook_inspect_data_model(wb2, &dataModelSummary) == 1, "data_model_inspection");

    CHECK(xlpp_sheet_set_vba_code_name(s2, "SmokeSheet") == 1, "set_vba_code_name");
    char codeName[64] = {};
    CHECK(xlpp_sheet_vba_code_name(s2, codeName, sizeof(codeName)) > 1 && std::strcmp(codeName, "SmokeSheet") == 0, "vba_code_name");
    CHECK(xlpp_workbook_set_vba_project_properties(wb2, "SmokeMacros", "C ABI VBA project", "smoke.chm", 7, "SmokeFlag = 1") == 1, "set_vba_project_properties");
    char projectName[64] = {};
    CHECK(xlpp_workbook_vba_project_name(wb2, projectName, sizeof(projectName)) > 1 && std::strcmp(projectName, "SmokeMacros") == 0, "vba_project_name");
    CHECK(xlpp_workbook_vba_project_help_context(wb2) == 7, "vba_project_help_context");
    CHECK(xlpp_workbook_set_vba_document_module_text(wb2, "SmokeSheet", "Private Sub Worksheet_Activate()\r\nEnd Sub\r\n") == 1, "set_vba_document_module");
    CHECK(xlpp_workbook_set_vba_class_module_text(wb2, "SmokeClass", "Public Function Value() As Long\r\nValue = 1\r\nEnd Function\r\n", 1, 1) == 1, "set_vba_class_module");
    CHECK(xlpp_workbook_set_vba_module_text(wb2, "Module1", "Sub Hello()\r\nEnd Sub\r\n") == 1, "set_vba_module");
    CHECK(xlpp_workbook_has_vba_project(wb2) == 1, "has_vba_project");
    const int vbaTextSize = xlpp_workbook_vba_module_text(wb2, "Module1", nullptr, 0);
    CHECK(vbaTextSize > 1, "vba_module_text_size");
    char vbaText[128] = {};
    CHECK(xlpp_workbook_vba_module_text(wb2, "Module1", vbaText, sizeof(vbaText)) > 1, "vba_module_text");
    CHECK(std::strstr(vbaText, "Sub Hello()") != nullptr, "vba_module_content");
    CHECK(xlpp_workbook_vba_module_count(wb2) >= 4, "vba_module_count");
    int classFlagsSeen = 0;
    for (uint64_t i = 0; i < xlpp_workbook_vba_module_count(wb2); ++i) {
        char moduleName[64] = {};
        xlpp_workbook_vba_module_name(wb2, i, moduleName, sizeof(moduleName));
        if (std::strcmp(moduleName, "SmokeClass") == 0)
            classFlagsSeen = xlpp_workbook_vba_module_read_only(wb2, i) && xlpp_workbook_vba_module_private(wb2, i);
    }
    CHECK(classFlagsSeen == 1, "vba_class_flags");
    CHECK(xlpp_workbook_vba_project_bytes(wb2, nullptr, 0) > 512, "vba_project_bytes");
    CHECK(xlpp_workbook_vba_source_editable(wb2) == 1, "vba_source_editable");
    CHECK(xlpp_workbook_has_vba_signature(wb2) == 0, "vba_signature_state");

    // Extended load/save options bridge used by managed bindings.
    const char* optionsPath = "smoke-options.xlsx";
    xlpp_save_options saveOptions{};
    saveOptions.compression_level = XLPP_COMPRESS_DEFAULT; saveOptions.compression_strategy = XLPP_STRATEGY_DEFAULT;
    saveOptions.parallel_sheets = 1; saveOptions.synchronize_changed_chart_caches_only = 1;
    saveOptions.atomic_write = 1; saveOptions.validate_before_save = 1;
    saveOptions.encryption_mode = XLPP_ENCRYPTION_AGILE_AES256_SHA512; saveOptions.encryption_spin_count = 100000; saveOptions.encryption_key_bits = 256;
    CHECK(xlpp_workbook_save_ex(wb2, optionsPath, &saveOptions) == 1, "save_options");
    const char* lowLatencyPath = "smoke-nondurable.xlsx";
    CHECK(xlpp_workbook_save_durable(wb2, lowLatencyPath, 0) == 1, "save_durable_opt_out");
    auto optionsWb = xlpp_workbook_create();
    xlpp_load_options loadOptions{}; loadOptions.verify_encryption_integrity = 1;
    CHECK(xlpp_workbook_load_ex(optionsWb, optionsPath, &loadOptions) == 1, "load_options");
    CHECK(xlpp_workbook_sheet_by_name(optionsWb, "Renamed") != nullptr, "options_roundtrip");
    xlpp_workbook_destroy(optionsWb);

    unsigned char* memoryBytes = nullptr; uint64_t memorySize = 0;
    CHECK(xlpp_workbook_save_bytes(wb2, &saveOptions, &memoryBytes, &memorySize) == 1 && memoryBytes != nullptr && memorySize > 0, "save_bytes");
    auto memoryWb = xlpp_workbook_create();
    LoadCallbackState callbackState{};
    xlpp_load_options memoryLoad{}; memoryLoad.verify_encryption_integrity = 1;
    memoryLoad.cancel = neverCancel; memoryLoad.progress = progressCallback; memoryLoad.callback_user = &callbackState;
    CHECK(xlpp_workbook_load_bytes(memoryWb, memoryBytes, memorySize, &memoryLoad) == 1, "load_bytes");
    CHECK(xlpp_workbook_sheet_by_name(memoryWb, "Renamed") != nullptr, "memory_roundtrip");
    CHECK(callbackState.calls > 0, "load_progress_callback");
    xlpp_workbook_destroy(memoryWb); xlpp_free_bytes(memoryBytes);

    // Password encryption is available from the C ABI and can be inspected/read back.
    const char* encryptedPath = "smoke-encrypted.xlsx";
    CHECK(xlpp_workbook_save_encrypted(wb2, encryptedPath, "capi-password", 1000, 1) == 1, "save_encrypted");
    int mode = -1; uint64_t keyBits = 0, spinCount = 0; char cipher[32] = {}, hash[32] = {};
    CHECK(xlpp_inspect_office_encryption(encryptedPath, &mode, &keyBits, &spinCount, cipher, sizeof(cipher), hash, sizeof(hash)) == 1, "inspect_encryption");
    CHECK(mode == XLPP_ENCRYPTION_AGILE_AES256_SHA512 && keyBits == 256, "encryption_profile");
    const char* standardEncryptedPath = "smoke-standard-encrypted.xlsx";
    CHECK(xlpp_workbook_save_encrypted_ex(wb2, standardEncryptedPath, "capi-standard", XLPP_ENCRYPTION_STANDARD_AES_SHA1, 128, 50000, 1) == 1, "save_standard_encrypted");
    int standardMode = -1; uint64_t standardBits = 0, standardSpin = 0;
    CHECK(xlpp_inspect_office_encryption(standardEncryptedPath, &standardMode, &standardBits, &standardSpin, nullptr, 0, nullptr, 0) == 1, "inspect_standard_encryption");
    CHECK(standardMode == XLPP_ENCRYPTION_STANDARD_AES_SHA1 && standardBits == 128 && standardSpin == 50000, "standard_encryption_profile");

    auto wb3 = xlpp_workbook_create();
    CHECK(xlpp_workbook_load_password(wb3, encryptedPath, "capi-password", 1) == 1, "load_password");
    CHECK(xlpp_workbook_sheet_by_name(wb3, "Renamed") != nullptr, "encrypted_roundtrip");
    xlpp_workbook_destroy(wb3);


    // P0Z-E stable-child-handle regression: child handles exposed by the C
    // ABI must survive growth of the owning native collection.
    int stableOk = 0;
    auto stableStyle = xlpp_workbook_add_named_style(wb2, "StableStyle", &stableOk);
    CHECK(stableOk == 1 && stableStyle != nullptr, "stable_named_style_create");
    for (int i = 0; i < 256; ++i) {
        const std::string name = "GrowStyle" + std::to_string(i);
        (void)xlpp_workbook_add_named_style(wb2, name.c_str(), &stableOk);
        CHECK(stableOk == 1, "stable_named_style_grow");
    }
    xlpp_namedstyle_set_name(stableStyle, "StableStyleRenamed");
    CHECK_STR(xlpp_namedstyle_name(stableStyle), "StableStyleRenamed", "stable_named_style_handle");

    auto stableName = xlpp_workbook_add_defined_name(wb2, "StableName", "Renamed!$A$1", &stableOk);
    CHECK(stableOk == 1 && stableName != nullptr, "stable_defined_name_create");
    for (int i = 0; i < 256; ++i) {
        const std::string name = "GrowName" + std::to_string(i);
        (void)xlpp_workbook_add_defined_name(wb2, name.c_str(), "Renamed!$A$1", &stableOk);
        CHECK(stableOk == 1, "stable_defined_name_grow");
    }
    xlpp_definedname_set_comment(stableName, "still-live");
    CHECK_STR(xlpp_definedname_comment(stableName), "still-live", "stable_defined_name_handle");

    auto stableTable = xlpp_sheet_add_table(refs, "StableTable", "A1:B2", &stableOk);
    CHECK(stableOk == 1 && stableTable != nullptr, "stable_table_create");
    for (int i = 0; i < 128; ++i) {
        const std::string name = "GrowTable" + std::to_string(i);
        (void)xlpp_sheet_add_table(refs, name.c_str(), "A1:B2", &stableOk);
        CHECK(stableOk == 1, "stable_table_grow");
    }
    auto stableColumn = xlpp_table_add_column(stableTable, "RootColumn");
    CHECK(stableColumn != nullptr, "stable_table_column_create");
    for (int i = 0; i < 128; ++i) {
        const std::string name = "C" + std::to_string(i);
        (void)xlpp_table_add_column(stableTable, name.c_str());
    }
    xlpp_tablecolumn_set_name(stableColumn, "RootColumnRenamed");
    char stableColumnName[64] = {};
    xlpp_tablecolumn_name(stableColumn, stableColumnName, sizeof(stableColumnName));
    CHECK_STR(stableColumnName, "RootColumnRenamed", "stable_table_column_handle");
    CHECK(xlpp_table_column_at(stableTable, 0) == stableColumn, "stable_table_column_identity");

    xlpp_sheet_add_chart(refs, 1);
    auto stableChart = xlpp_sheet_chart_at(refs, xlpp_sheet_chart_count(refs) - 1);
    CHECK(stableChart != nullptr, "stable_chart_create");
    auto stableSeries = xlpp_chart_add_series(stableChart, "RootSeries");
    CHECK(stableSeries != nullptr, "stable_chart_series_create");
    for (int i = 0; i < 128; ++i) {
        const std::string title = "S" + std::to_string(i);
        (void)xlpp_chart_add_series(stableChart, title.c_str());
    }
    xlpp_chartseries_set_title(stableSeries, "RootSeriesRenamed");
    CHECK(xlpp_chart_series_at(stableChart, 0) == stableSeries, "stable_chart_series_identity");

    auto dvCollection = xlpp_sheet_data_validations(refs);
    auto stableValidation = xlpp_dvcollection_add(dvCollection, 1, "D1");
    CHECK(stableValidation != nullptr, "stable_validation_create");
    for (int i = 0; i < 128; ++i) {
        const std::string ref = "D" + std::to_string(i + 2);
        (void)xlpp_dvcollection_add(dvCollection, 1, ref.c_str());
    }
    xlpp_datavalidation_set_formula1(stableValidation, "1");
    char stableValidationFormula[16] = {};
    xlpp_datavalidation_formula1(stableValidation, stableValidationFormula, sizeof(stableValidationFormula));
    CHECK_STR(stableValidationFormula, "1", "stable_validation_handle");

    auto customProps = xlpp_workbook_custom_properties(wb2);
    auto stableProperty = xlpp_customprops_add(customProps, "RootProperty", "root", 0);
    CHECK(stableProperty != nullptr, "stable_custom_property_create");
    for (int i = 0; i < 128; ++i) {
        const std::string name = "P" + std::to_string(i);
        (void)xlpp_customprops_add(customProps, name.c_str(), "v", 0);
    }
    CHECK_STR(xlpp_customprop_name(stableProperty), "RootProperty", "stable_custom_property_handle");

    auto richText = xlpp_richtext_create();
    auto stableRun = xlpp_richtext_add_run(richText, "root");
    CHECK(stableRun != nullptr, "stable_richtext_run_create");
    for (int i = 0; i < 128; ++i) (void)xlpp_richtext_add_run(richText, "grow");
    xlpp_richtextrun_set_text(stableRun, "still-live");
    CHECK_STR(xlpp_richtextrun_text(stableRun), "still-live", "stable_richtext_run_handle");
    xlpp_richtext_destroy(richText);

    xlpp_workbook_destroy(wb2);
    std::remove(path);
    std::remove(encryptedPath);
    std::remove(standardEncryptedPath);
    std::remove(optionsPath);
    std::remove(lowLatencyPath);

    fprintf(stderr, "XLPP C API smoke test OK\n");
    return 0;
}
