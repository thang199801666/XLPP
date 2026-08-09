// XLPP C API Implementation
// Binding package version: 1.12.0
#include "xlpp_capi.h"
#include <XLPP/XLPP.h>
#include <XLPP/Version.h>
#include <XLPP/Cell/RichText.h>
#include <string>
#include <vector>
#include <cstring>
#include <memory>
#include <cstdio>
#include <algorithm>
#include <limits>
#include <sstream>

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

struct xlpp_dependency_graph_t {
    explicit xlpp_dependency_graph_t(xlpp::FormulaDependencyGraph value) : graph(std::move(value)) {}
    xlpp::FormulaDependencyGraph graph;
};

struct xlpp_validation_report_t {
    explicit xlpp_validation_report_t(xlpp::WorkbookValidationReport value) : report(std::move(value)) {}
    xlpp::WorkbookValidationReport report;
};

struct xlpp_external_value_t {
    std::optional<xlpp::CellValue> value;
};

namespace {
thread_local std::string g_lastError;
void clearError() noexcept { g_lastError.clear(); }
void setError(const char* message) noexcept { g_lastError = message ? message : "XLPP C API error"; }
void setError(const std::exception& error) noexcept { g_lastError = error.what(); }

xlpp::LoadOptions toLoadOptions(const xlpp_load_options* input) {
    xlpp::LoadOptions options;
    if (!input) return options;
    options.lenient = input->lenient != 0;
    options.maxEntries = static_cast<std::size_t>(input->max_entries);
    options.maxEntryBytes = static_cast<std::size_t>(input->max_entry_bytes);
    options.maxTotalBytes = static_cast<std::size_t>(input->max_total_bytes);
    options.maxFileBytes = static_cast<std::size_t>(input->max_file_bytes);
    if (input->password) options.password = input->password;
    options.verifyEncryptionIntegrity = input->verify_encryption_integrity != 0;
    if (input->cancel) {
        const auto callback = input->cancel; void* user = input->callback_user;
        options.cancel = [callback, user]() { return callback(user) != 0; };
    }
    if (input->progress) {
        const auto callback = input->progress; void* user = input->callback_user;
        options.progress = [callback, user](std::size_t done, std::size_t total) {
            callback(user, static_cast<uint64_t>(done), static_cast<uint64_t>(total));
        };
    }
    return options;
}

xlpp::SaveOptions toSaveOptions(const xlpp_save_options* input) {
    xlpp::SaveOptions options;
    if (!input) return options;
    switch (input->compression_level) {
        case XLPP_COMPRESS_STORE: options.compressionLevel = xlpp::CompressionLevel::Store; break;
        case XLPP_COMPRESS_FASTEST: options.compressionLevel = xlpp::CompressionLevel::Fastest; break;
        case XLPP_COMPRESS_BEST: options.compressionLevel = xlpp::CompressionLevel::Best; break;
        default: options.compressionLevel = xlpp::CompressionLevel::Default; break;
    }
    switch (input->compression_strategy) {
        case XLPP_STRATEGY_FILTERED: options.compressionStrategy = xlpp::CompressionStrategy::Filtered; break;
        case XLPP_STRATEGY_HUFFMAN_ONLY: options.compressionStrategy = xlpp::CompressionStrategy::HuffmanOnly; break;
        case XLPP_STRATEGY_RLE: options.compressionStrategy = xlpp::CompressionStrategy::Rle; break;
        case XLPP_STRATEGY_FIXED: options.compressionStrategy = xlpp::CompressionStrategy::Fixed; break;
        default: options.compressionStrategy = xlpp::CompressionStrategy::Default; break;
    }
    options.parallelWorkers = static_cast<std::size_t>(input->parallel_workers);
    options.parallelSheets = input->parallel_sheets != 0;
    options.parallelRows = input->parallel_rows != 0;
    options.strictNamespace = input->strict_namespace != 0;
    options.synchronizeChartCaches = input->synchronize_chart_caches != 0;
    options.synchronizeChangedChartCachesOnly = input->synchronize_changed_chart_caches_only != 0;
    options.calculateFormulasBeforeSave = input->calculate_formulas_before_save != 0;
    options.atomicWrite = input->atomic_write != 0;
    // Preserve the stable C ABI: durable file synchronization remains enabled
    // by the C++ SaveOptions default rather than extending xlpp_save_options.
    options.validateBeforeSave = input->validate_before_save != 0;
    if (input->encryption_password) options.encryptionPassword = input->encryption_password;
    switch (input->encryption_mode) {
        case XLPP_ENCRYPTION_NONE: options.encryptionMode = xlpp::OfficeEncryptionMode::None; break;
        case XLPP_ENCRYPTION_STANDARD_AES_SHA1: options.encryptionMode = xlpp::OfficeEncryptionMode::StandardAesSha1; break;
        default: options.encryptionMode = xlpp::OfficeEncryptionMode::AgileAes256Sha512; break;
    }
    if (input->encryption_spin_count) options.encryptionSpinCount = static_cast<std::uint32_t>(input->encryption_spin_count);
    if (input->encryption_key_bits) options.encryptionKeyBits = static_cast<std::uint32_t>(input->encryption_key_bits);
    return options;
}

xlpp::StructuralEditOptions toStructuralOptions(const xlpp_structural_options* input) {
    xlpp::StructuralEditOptions options;
    if (!input) return options;
    options.transactional = input->transactional != 0;
    options.updateDefinedNames = input->update_defined_names != 0;
    options.recalculateFormulas = input->recalculate_formulas != 0;
    options.synchronizeChartCaches = input->synchronize_chart_caches != 0;
    options.changedChartCachesOnly = input->changed_chart_caches_only != 0;
    options.failOnInvalidReference = input->fail_on_invalid_reference != 0;
    return options;
}

xlpp::CalculationOptions toCalculationOptions(const xlpp_calculation_options* input) {
    xlpp::CalculationOptions options;
    if (!input) return options;
    options.recursiveDependencies = input->recursive_dependencies != 0;
    options.updateCachedValues = input->update_cached_values != 0;
    options.evaluateVolatileFunctions = input->evaluate_volatile_functions != 0;
    options.spillDynamicArrays = input->spill_dynamic_arrays != 0;
    options.iterativeCalculation = input->iterative_calculation != 0;
    if (input->max_iterations) options.maxIterations = static_cast<std::size_t>(input->max_iterations);
    if (input->max_change >= 0.0) options.maxChange = input->max_change;
    if (input->max_depth) options.maxDepth = static_cast<std::size_t>(input->max_depth);
    if (input->external_reference_resolver) {
        const auto callback = input->external_reference_resolver;
        void* user = input->external_reference_user;
        options.externalReferenceResolver = [callback, user](const std::string& workbookToken,
                                                             const std::string& sheetName,
                                                             const std::string& address) -> std::optional<xlpp::CellValue> {
            xlpp_external_value_t output;
            if (!callback(user, workbookToken.c_str(), sheetName.c_str(), address.c_str(), &output))
                return std::nullopt;
            return output.value;
        };
    }
    return options;
}

void fillCalculationReport(const xlpp::CalculationReport& r, xlpp_calculation_report* out) {
    if (!out) return;
    out->formula_cells_visited = r.formulaCellsVisited;
    out->formula_cells_evaluated = r.formulaCellsEvaluated;
    out->cached_values_updated = r.cachedValuesUpdated;
    out->dependency_evaluations = r.dependencyEvaluations;
    out->defined_names_resolved = r.definedNamesResolved;
    out->circular_references = r.circularReferences;
    out->unsupported_formulas = r.unsupportedFormulas;
    out->evaluation_errors = r.evaluationErrors;
    out->dynamic_arrays_spilled = r.dynamicArraysSpilled;
    out->spill_cells_updated = r.spillCellsUpdated;
    out->spill_conflicts = r.spillConflicts;
    out->structured_references_resolved = r.structuredReferencesResolved;
    out->iterative_iterations = r.iterativeIterations;
    out->iterative_convergence_failures = r.iterativeConvergenceFailures;
    out->external_references_resolved = r.externalReferencesResolved;
    out->unresolved_external_references = r.unresolvedExternalReferences;
    out->success = r.success() ? 1 : 0;
}

void fillStructuralReport(const xlpp::StructuralEditReport& r, xlpp_structural_report* out) {
    if (!out) return;
    out->worksheets_visited = r.worksheetsVisited;
    out->cells_moved = r.cellsMoved;
    out->cells_removed = r.cellsRemoved;
    out->formulas_updated = r.formulasUpdated;
    out->formula_metadata_updated = r.formulaMetadataUpdated;
    out->worksheet_references_updated = r.worksheetReferencesUpdated;
    out->defined_names_updated = r.definedNamesUpdated;
    out->chart_references_updated = r.chartReferencesUpdated;
    out->pivot_references_updated = r.pivotReferencesUpdated;
    out->drawing_anchors_updated = r.drawingAnchorsUpdated;
    out->hyperlinks_updated = r.hyperlinksUpdated;
    out->references_invalidated = r.referencesInvalidated;
    out->formulas_calculated = r.formulasCalculated;
    out->chart_caches_updated = r.chartCachesUpdated;
    out->success = r.success() ? 1 : 0;
}
}

// String copy helper into caller buffer.
static void copyStr(const std::string& s, char* out, int outSize) {
    if (!out || outSize <= 0) return;
    const auto n = (std::min)(s.size(), static_cast<std::size_t>(outSize - 1));
    std::memcpy(out, s.c_str(), n);
    out[n] = '\0';
}

static int copyStrSized(const std::string& s, char* out, int outSize) {
    copyStr(s, out, outSize);
    const auto required = s.size() + 1;
    return required > static_cast<std::size_t>((std::numeric_limits<int>::max)())
        ? (std::numeric_limits<int>::max)() : static_cast<int>(required);
}

// ============================================================
// Workbook
// ============================================================
extern "C" {

XLPP_API const char* xlpp_version(void) {
    return XLPP_VERSION_STRING;
}

XLPP_API uint64_t xlpp_c_abi_version(void) {
    return static_cast<uint64_t>(XLPP_C_ABI_VERSION);
}

XLPP_API uint64_t xlpp_capabilities(void) {
    return XLPP_CAP_FORMULA_ENGINE | XLPP_CAP_STREAMING | XLPP_CAP_ENCRYPTION |
           XLPP_CAP_CHARTS | XLPP_CAP_PIVOT | XLPP_CAP_VBA |
           XLPP_CAP_EXTERNAL_DATA_INSPECTION | XLPP_CAP_DATA_MODEL_INSPECTION |
           XLPP_CAP_DIRTY_RECALC | XLPP_CAP_ADVANCED_AUTOFILTER;
}

XLPP_API int xlpp_workbook_inspect_external_data(xlpp_workbook wb, xlpp_external_data_summary* out) {
    if (!wb || !out) return 0;
    try {
        const auto info = WB(wb)->inspectExternalData();
        out->external_workbooks = static_cast<uint64_t>(info.externalWorkbooks.size());
        out->connections = static_cast<uint64_t>(info.connections.size());
        out->query_tables = static_cast<uint64_t>(info.queryTables.size());
        out->power_query_parts = static_cast<uint64_t>(info.powerQueryParts.size());
        out->web_query_parts = static_cast<uint64_t>(info.webQueryParts.size());
        out->unknown_connection_parts = static_cast<uint64_t>(info.unknownConnectionParts.size());
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_inspect_data_model(xlpp_workbook wb, xlpp_data_model_summary* out) {
    if (!wb || !out) return 0;
    try {
        const auto info = WB(wb)->inspectDataModel();
        out->present = info.present ? 1 : 0;
        out->has_olap_pivot_caches = info.hasOlapPivotCaches ? 1 : 0;
        out->model_parts = static_cast<uint64_t>(info.modelParts.size());
        out->model_relationships = static_cast<uint64_t>(info.modelRelationships.size());
        out->olap_pivot_cache_parts = static_cast<uint64_t>(info.olapPivotCacheParts.size());
        out->warnings = static_cast<uint64_t>(info.warnings.size());
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API xlpp_workbook xlpp_workbook_create(void) {
    return reinterpret_cast<xlpp_workbook>(new xlpp::Workbook());
}

XLPP_API void xlpp_workbook_destroy(xlpp_workbook wb) {
    delete WB(wb);
}

XLPP_API xlpp_worksheet xlpp_workbook_add_sheet(xlpp_workbook wb, const char* name) {
    if (!wb) { setError("Workbook handle is null"); return nullptr; }
    if (!name || !name[0]) { setError("Workbook and sheet name are required"); return nullptr; }
    try {
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->addWorksheet(name));
    } catch (const std::exception& e) { setError(e); return nullptr; }
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

XLPP_API xlpp_worksheet xlpp_workbook_copy_sheet(xlpp_workbook wb, xlpp_worksheet src, const char* new_name) {
    try {
        return reinterpret_cast<xlpp_worksheet>(&WB(wb)->copyWorksheet(*WS(src), new_name));
    } catch (...) { return nullptr; }
}

XLPP_API int xlpp_workbook_sheet_index(xlpp_workbook wb, xlpp_worksheet ws) {
    try {
        return static_cast<int>(WB(wb)->index(*WS(ws)));
    } catch (...) { return -1; }
}

XLPP_API const char* xlpp_workbook_sheet_name(xlpp_workbook wb, int index, char* out, int outSize) {
    try {
        copyStr((*WB(wb))[static_cast<std::size_t>(index)].name(), out, outSize);
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
    return out;
}

XLPP_API int xlpp_workbook_sheet_names_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->sheetNames().size());
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

XLPP_API int xlpp_workbook_save_durable(xlpp_workbook wb, const char* path, int durable_write) {
    clearError();
    if (!wb || !path) { setError("Workbook handle and path are required"); return 0; }
    try {
        xlpp::SaveOptions options;
        options.durableWrite = durable_write != 0;
        WB(wb)->save(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_load_password(xlpp_workbook wb, const char* path, const char* password, int verify_integrity) {
    clearError();
    if (!wb || !path) { setError("Workbook handle and path are required"); return 0; }
    try {
        xlpp::LoadOptions options;
        if (password) options.password = password;
        options.verifyEncryptionIntegrity = verify_integrity != 0;
        WB(wb)->load(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_encrypted(xlpp_workbook wb, const char* path, const char* password, uint64_t spin_count, int calculate_formulas) {
    clearError();
    if (!wb || !path || !password || !password[0]) { setError("Workbook, path, and non-empty password are required"); return 0; }
    try {
        xlpp::SaveOptions options;
        options.encryptionPassword = password;
        if (spin_count) options.encryptionSpinCount = static_cast<std::uint32_t>(spin_count);
        options.calculateFormulasBeforeSave = calculate_formulas != 0;
        WB(wb)->save(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_encrypted_ex(xlpp_workbook wb, const char* path, const char* password, int encryption_mode, uint64_t key_bits, uint64_t spin_count, int calculate_formulas) {
    clearError();
    if (!wb || !path || !password || !password[0]) { setError("Workbook, path, and non-empty password are required"); return 0; }
    try {
        xlpp::SaveOptions options;
        options.encryptionPassword = password;
        switch (encryption_mode) {
            case XLPP_ENCRYPTION_AGILE_AES256_SHA512:
                options.encryptionMode = xlpp::OfficeEncryptionMode::AgileAes256Sha512;
                break;
            case XLPP_ENCRYPTION_STANDARD_AES_SHA1:
                options.encryptionMode = xlpp::OfficeEncryptionMode::StandardAesSha1;
                break;
            default:
                setError("Unsupported encryption mode");
                return 0;
        }
        if (key_bits) options.encryptionKeyBits = static_cast<std::uint32_t>(key_bits);
        if (spin_count) options.encryptionSpinCount = static_cast<std::uint32_t>(spin_count);
        options.calculateFormulasBeforeSave = calculate_formulas != 0;
        WB(wb)->save(std::filesystem::path(path), options);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_load_ex(xlpp_workbook wb, const char* path, const xlpp_load_options* input) {
    clearError();
    if (!wb || !path) { setError("Workbook handle and path are required"); return 0; }
    try {
        WB(wb)->load(std::filesystem::path(path), toLoadOptions(input));
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_ex(xlpp_workbook wb, const char* path, const xlpp_save_options* input) {
    clearError();
    if (!wb || !path) { setError("Workbook handle and path are required"); return 0; }
    try {
        WB(wb)->save(std::filesystem::path(path), toSaveOptions(input));
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_load_bytes(xlpp_workbook wb, const unsigned char* bytes, uint64_t size, const xlpp_load_options* input) {
    clearError();
    if (!wb || (!bytes && size != 0)) { setError("Workbook handle and byte buffer are required"); return 0; }
    try {
        std::string raw;
        if (size != 0) raw.assign(reinterpret_cast<const char*>(bytes), static_cast<std::size_t>(size));
        std::istringstream stream(raw, std::ios::binary);
        WB(wb)->load(stream, toLoadOptions(input));
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_save_bytes(xlpp_workbook wb, const xlpp_save_options* input, unsigned char** bytes, uint64_t* size) {
    clearError();
    if (!wb || !bytes || !size) { setError("Workbook handle and output pointers are required"); return 0; }
    *bytes = nullptr; *size = 0;
    try {
        std::ostringstream stream(std::ios::binary);
        WB(wb)->save(stream, toSaveOptions(input));
        const auto raw = stream.str();
        auto* output = new unsigned char[raw.size()];
        if (!raw.empty()) std::memcpy(output, raw.data(), raw.size());
        *bytes = output; *size = static_cast<uint64_t>(raw.size());
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API void xlpp_free_bytes(unsigned char* bytes) { delete[] bytes; }

XLPP_API int xlpp_workbook_calculate(xlpp_workbook wb, xlpp_calculation_report* out) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return 0; }
    try {
        const auto r = WB(wb)->calculateFormulas();
        fillCalculationReport(r, out);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_calculate_ex(xlpp_workbook wb, int iterative, uint64_t max_iterations, double max_change, xlpp_calculation_report* out) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return 0; }
    try {
        xlpp::CalculationOptions options;
        options.iterativeCalculation = iterative != 0;
        if (max_iterations) options.maxIterations = static_cast<std::size_t>(max_iterations);
        if (max_change >= 0.0) options.maxChange = max_change;
        const auto r = WB(wb)->calculateFormulas(options);
        fillCalculationReport(r, out);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API void xlpp_external_value_set_empty(xlpp_external_value value) { if (value) value->value = xlpp::CellValue{std::monostate{}}; }
XLPP_API void xlpp_external_value_set_number(xlpp_external_value value, double number) { if (value) value->value = xlpp::CellValue{number}; }
XLPP_API void xlpp_external_value_set_bool(xlpp_external_value value, int boolean_value) { if (value) value->value = xlpp::CellValue{boolean_value != 0}; }
XLPP_API void xlpp_external_value_set_string(xlpp_external_value value, const char* string_value) { if (value) value->value = xlpp::CellValue{std::string(string_value ? string_value : "")}; }
XLPP_API void xlpp_external_value_set_error(xlpp_external_value value, int error_code) {
    if (!value) return;
    if (error_code < XLPP_ERROR_NULL || error_code > XLPP_ERROR_CALC) error_code = XLPP_ERROR_VALUE;
    value->value = xlpp::CellValue{static_cast<xlpp::CellError>(error_code)};
}
XLPP_API void xlpp_external_value_set_date(xlpp_external_value value, int year, int month, int day, int hour, int minute, double second, int /*has_time*/) {
    if (value) value->value = xlpp::CellValue{xlpp::DateTime{year, month, day, hour, minute, second}};
}

XLPP_API int xlpp_workbook_calculate_options(xlpp_workbook wb, const xlpp_calculation_options* input, xlpp_calculation_report* out) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return 0; }
    try {
        const auto r = WB(wb)->calculateFormulas(toCalculationOptions(input));
        fillCalculationReport(r, out);
        return 1;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_structural_edit(xlpp_workbook wb, const char* sheet_name, int kind, uint64_t index, uint64_t amount, int fail_on_invalid_reference, xlpp_structural_report* out) {
    clearError();
    if (!wb || !sheet_name || !sheet_name[0]) { setError("Workbook and sheet name are required"); return 0; }
    try {
        xlpp::StructuralEditKind editKind;
        switch (kind) {
            case XLPP_STRUCT_INSERT_ROWS: editKind = xlpp::StructuralEditKind::InsertRows; break;
            case XLPP_STRUCT_DELETE_ROWS: editKind = xlpp::StructuralEditKind::DeleteRows; break;
            case XLPP_STRUCT_INSERT_COLUMNS: editKind = xlpp::StructuralEditKind::InsertColumns; break;
            case XLPP_STRUCT_DELETE_COLUMNS: editKind = xlpp::StructuralEditKind::DeleteColumns; break;
            default: setError("Invalid structural edit kind"); return 0;
        }
        xlpp::StructuralEditOptions options;
        options.failOnInvalidReference = fail_on_invalid_reference != 0;
        const auto r = WB(wb)->applyStructuralEdit(xlpp::StructuralEdit{sheet_name, editKind, static_cast<std::size_t>(index), static_cast<std::size_t>(amount)}, options);
        fillStructuralReport(r, out);
        return r.success() ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_structural_edit_ex(xlpp_workbook wb, const char* sheet_name, int kind, uint64_t index, uint64_t amount, const xlpp_structural_options* input, xlpp_structural_report* out) {
    clearError();
    if (!wb || !sheet_name || !sheet_name[0]) { setError("Workbook and sheet name are required"); return 0; }
    try {
        xlpp::StructuralEditKind editKind;
        switch (kind) {
            case XLPP_STRUCT_INSERT_ROWS: editKind = xlpp::StructuralEditKind::InsertRows; break;
            case XLPP_STRUCT_DELETE_ROWS: editKind = xlpp::StructuralEditKind::DeleteRows; break;
            case XLPP_STRUCT_INSERT_COLUMNS: editKind = xlpp::StructuralEditKind::InsertColumns; break;
            case XLPP_STRUCT_DELETE_COLUMNS: editKind = xlpp::StructuralEditKind::DeleteColumns; break;
            default: setError("Invalid structural edit kind"); return 0;
        }
        const auto r = WB(wb)->applyStructuralEdit(
            xlpp::StructuralEdit{sheet_name, editKind, static_cast<std::size_t>(index), static_cast<std::size_t>(amount)},
            toStructuralOptions(input));
        fillStructuralReport(r, out);
        return r.success() ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_rename_sheet(xlpp_workbook wb, const char* old_name, const char* new_name, int recalculate_formulas, int synchronize_chart_caches, int changed_chart_caches_only, xlpp_worksheet_rename_report* out) {
    clearError();
    if (!wb || !old_name || !old_name[0] || !new_name || !new_name[0]) { setError("Workbook, old name, and new name are required"); return 0; }
    try {
        xlpp::WorksheetRenameOptions options;
        options.recalculateFormulas = recalculate_formulas != 0;
        options.synchronizeChartCaches = synchronize_chart_caches != 0;
        options.changedChartCachesOnly = changed_chart_caches_only != 0;
        const auto r = WB(wb)->renameWorksheet(old_name, new_name, options);
        if (out) {
            out->worksheets_visited = r.worksheetsVisited;
            out->formulas_updated = r.formulasUpdated;
            out->formula_metadata_updated = r.formulaMetadataUpdated;
            out->defined_names_updated = r.definedNamesUpdated;
            out->chart_references_updated = r.chartReferencesUpdated;
            out->pivot_references_updated = r.pivotReferencesUpdated;
            out->hyperlinks_updated = r.hyperlinksUpdated;
            out->references_updated = r.referencesUpdated;
            out->formulas_calculated = r.formulasCalculated;
            out->chart_caches_updated = r.chartCachesUpdated;
            out->success = r.success() ? 1 : 0;
        }
        return r.success() ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API int xlpp_workbook_synchronize_chart_caches(xlpp_workbook wb, const xlpp_chart_cache_sync_options* input, xlpp_chart_cache_sync_report* out) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return 0; }
    try {
        xlpp::ChartCacheSyncOptions options;
        if (input) {
            options.synchronizeTitles = input->synchronize_titles != 0;
            options.synchronizeCategories = input->synchronize_categories != 0;
            options.synchronizeValues = input->synchronize_values != 0;
            options.changedReferencesOnly = input->changed_references_only != 0;
            options.clearUnsupportedReferences = input->clear_unsupported_references != 0;
        }
        const auto r = WB(wb)->synchronizeChartCaches(options);
        if (out) {
            out->charts_visited = r.chartsVisited;
            out->series_visited = r.seriesVisited;
            out->references_checked = r.referencesChecked;
            out->references_unchanged = r.referencesUnchanged;
            out->dependencies_registered = r.dependenciesRegistered;
            out->dependencies_changed = r.dependenciesChanged;
            out->caches_updated = r.cachesUpdated;
            out->caches_cleared = r.cachesCleared;
            out->references_skipped = r.referencesSkipped;
            out->success = r.success() ? 1 : 0;
        }
        return r.success() ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API void xlpp_workbook_reset_chart_cache_tracking(xlpp_workbook wb) { if (wb) WB(wb)->resetChartCacheDependencyTracking(); }
XLPP_API uint64_t xlpp_workbook_tracked_chart_cache_dependencies(xlpp_workbook wb) { return wb ? static_cast<uint64_t>(WB(wb)->trackedChartCacheDependencyCount()) : 0; }

XLPP_API xlpp_dependency_graph xlpp_workbook_dependency_graph(xlpp_workbook wb) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return nullptr; }
    try { return new xlpp_dependency_graph_t(WB(wb)->dependencyGraph()); }
    catch (const std::exception& e) { setError(e); return nullptr; }
}
XLPP_API void xlpp_dependency_graph_destroy(xlpp_dependency_graph graph) { delete graph; }
XLPP_API int xlpp_dependency_graph_report(xlpp_dependency_graph graph, xlpp_dependency_report* out) {
    if (!graph || !out) return 0;
    const auto& r = graph->graph.report();
    out->formula_cells = r.formulaCells; out->edges = r.edges; out->cell_or_range_edges = r.cellOrRangeEdges;
    out->defined_name_edges = r.definedNameEdges; out->table_edges = r.tableEdges; out->external_edges = r.externalEdges;
    out->volatile_references = r.volatileReferences; out->unresolved_symbols = r.unresolvedSymbols;
    return 1;
}
XLPP_API uint64_t xlpp_dependency_graph_edge_count(xlpp_dependency_graph graph) { return graph ? static_cast<uint64_t>(graph->graph.edges().size()) : 0; }
static const xlpp::FormulaDependency* dependencyEdge(xlpp_dependency_graph graph, uint64_t index) {
    if (!graph || index >= graph->graph.edges().size()) return nullptr;
    return &graph->graph.edges()[static_cast<std::size_t>(index)];
}
XLPP_API int xlpp_dependency_graph_edge_kind(xlpp_dependency_graph graph, uint64_t index) { const auto* e = dependencyEdge(graph,index); return e ? static_cast<int>(e->kind) : -1; }
XLPP_API int xlpp_dependency_graph_edge_dependent_sheet(xlpp_dependency_graph graph, uint64_t index, char* out, int size) { const auto* e=dependencyEdge(graph,index); return copyStrSized(e?e->dependentSheet:std::string{},out,size); }
XLPP_API int xlpp_dependency_graph_edge_dependent_cell(xlpp_dependency_graph graph, uint64_t index, char* out, int size) { const auto* e=dependencyEdge(graph,index); return copyStrSized(e?e->dependentCell:std::string{},out,size); }
XLPP_API int xlpp_dependency_graph_edge_precedent_sheet(xlpp_dependency_graph graph, uint64_t index, char* out, int size) { const auto* e=dependencyEdge(graph,index); return copyStrSized(e?e->precedentSheet:std::string{},out,size); }
XLPP_API int xlpp_dependency_graph_edge_precedent_reference(xlpp_dependency_graph graph, uint64_t index, char* out, int size) { const auto* e=dependencyEdge(graph,index); return copyStrSized(e?e->precedentReference:std::string{},out,size); }
XLPP_API int xlpp_dependency_graph_edge_symbol(xlpp_dependency_graph graph, uint64_t index, char* out, int size) { const auto* e=dependencyEdge(graph,index); return copyStrSized(e?e->symbol:std::string{},out,size); }
XLPP_API int xlpp_dependency_graph_depends_on(xlpp_dependency_graph graph, const char* ds, const char* dc, const char* ps, const char* pc) {
    if (!graph || !ds || !dc || !ps || !pc) return 0;
    return graph->graph.dependsOn(ds,dc,ps,pc) ? 1 : 0;
}

XLPP_API xlpp_validation_report xlpp_workbook_validate(xlpp_workbook wb, const xlpp_validation_options* input) {
    clearError();
    if (!wb) { setError("Workbook handle is null"); return nullptr; }
    try {
        xlpp::WorkbookValidationOptions options;
        if (input) {
            options.validateWorksheetNames = input->validate_worksheet_names != 0;
            options.validateDefinedNames = input->validate_defined_names != 0;
            options.validateTables = input->validate_tables != 0;
            options.validatePivots = input->validate_pivots != 0;
        }
        return new xlpp_validation_report_t(WB(wb)->validate(options));
    } catch (const std::exception& e) { setError(e); return nullptr; }
}
XLPP_API void xlpp_validation_report_destroy(xlpp_validation_report report) { delete report; }
XLPP_API uint64_t xlpp_validation_error_count(xlpp_validation_report report) { return report ? report->report.errorCount : 0; }
XLPP_API uint64_t xlpp_validation_warning_count(xlpp_validation_report report) { return report ? report->report.warningCount : 0; }
XLPP_API uint64_t xlpp_validation_issue_count(xlpp_validation_report report) { return report ? static_cast<uint64_t>(report->report.issues.size()) : 0; }
static const xlpp::WorkbookValidationIssue* validationIssue(xlpp_validation_report report, uint64_t index) {
    if (!report || index >= report->report.issues.size()) return nullptr;
    return &report->report.issues[static_cast<std::size_t>(index)];
}
XLPP_API int xlpp_validation_issue_severity(xlpp_validation_report report, uint64_t index) { const auto* i=validationIssue(report,index); return i ? static_cast<int>(i->severity) : -1; }
XLPP_API int xlpp_validation_issue_code(xlpp_validation_report report, uint64_t index, char* out, int size) { const auto* i=validationIssue(report,index); return copyStrSized(i?i->code:std::string{},out,size); }
XLPP_API int xlpp_validation_issue_message(xlpp_validation_report report, uint64_t index, char* out, int size) { const auto* i=validationIssue(report,index); return copyStrSized(i?i->message:std::string{},out,size); }
XLPP_API int xlpp_validation_issue_worksheet(xlpp_validation_report report, uint64_t index, char* out, int size) { const auto* i=validationIssue(report,index); return copyStrSized(i?i->worksheet:std::string{},out,size); }

XLPP_API int xlpp_workbook_add_vba_project(xlpp_workbook wb, const char* path) { clearError(); if(!wb||!path)return 0; try{WB(wb)->addVbaProject(std::filesystem::path(path));return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_set_vba_project(xlpp_workbook wb, const unsigned char* bytes, uint64_t size) { clearError(); if(!wb||(!bytes&&size))return 0; try{WB(wb)->setVbaProject(std::vector<unsigned char>(bytes,bytes+size));return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_has_vba_project(xlpp_workbook wb) { return wb && WB(wb)->hasVbaProject() ? 1 : 0; }
XLPP_API int xlpp_workbook_remove_vba_project(xlpp_workbook wb) { return wb && WB(wb)->removeVbaProject() ? 1 : 0; }
XLPP_API int xlpp_workbook_set_vba_module_text(xlpp_workbook wb, const char* name, const char* source) { clearError(); if(!wb||!name||!source)return 0; try{WB(wb)->setVbaModuleText(name,source);return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_set_vba_class_module_text(xlpp_workbook wb, const char* name, const char* source, int readOnly, int privateModule) { clearError(); if(!wb||!name||!source)return 0; try{WB(wb)->setVbaClassModuleText(name,source,readOnly!=0,privateModule!=0);return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_set_vba_document_module_text(xlpp_workbook wb, const char* name, const char* source) { clearError(); if(!wb||!name||!source)return 0; try{WB(wb)->setVbaDocumentModuleText(name,source);return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_vba_module_text(xlpp_workbook wb, const char* name, char* out, int size) { if(!wb||!name)return 0; const auto v=WB(wb)->vbaModuleText(name); if(!v)return 0; return copyStrSized(*v,out,size); }
XLPP_API uint64_t xlpp_workbook_vba_module_count(xlpp_workbook wb) { return wb ? static_cast<uint64_t>(WB(wb)->vbaModules().size()) : 0; }
XLPP_API int xlpp_workbook_vba_module_type(xlpp_workbook wb, uint64_t index) { if(!wb)return -1; const auto v=WB(wb)->vbaModules(); return index<v.size()?static_cast<int>(v[static_cast<std::size_t>(index)].type):-1; }
XLPP_API int xlpp_workbook_vba_module_read_only(xlpp_workbook wb, uint64_t index) { if(!wb)return 0; const auto v=WB(wb)->vbaModules(); return index<v.size()&&v[static_cast<std::size_t>(index)].readOnly?1:0; }
XLPP_API int xlpp_workbook_vba_module_private(xlpp_workbook wb, uint64_t index) { if(!wb)return 0; const auto v=WB(wb)->vbaModules(); return index<v.size()&&v[static_cast<std::size_t>(index)].privateModule?1:0; }
XLPP_API int xlpp_workbook_vba_module_name(xlpp_workbook wb, uint64_t index, char* out, int size) { if(!wb)return copyStrSized({},out,size); const auto v=WB(wb)->vbaModules(); return copyStrSized(index<v.size()?v[static_cast<std::size_t>(index)].name:std::string{},out,size); }
XLPP_API int xlpp_workbook_vba_module_source(xlpp_workbook wb, uint64_t index, char* out, int size) { if(!wb)return copyStrSized({},out,size); const auto v=WB(wb)->vbaModules(); return copyStrSized(index<v.size()?v[static_cast<std::size_t>(index)].source:std::string{},out,size); }
XLPP_API int xlpp_workbook_remove_vba_module(xlpp_workbook wb, const char* name) { clearError(); if(!wb||!name)return 0; try{return WB(wb)->removeVbaModule(name)?1:0;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API uint64_t xlpp_workbook_vba_project_bytes(xlpp_workbook wb, unsigned char* out, uint64_t size) { if(!wb)return 0; const auto bytes=WB(wb)->vbaProjectBytes(); const auto required=static_cast<uint64_t>(bytes.size()); if(out&&size){const auto count=static_cast<std::size_t>(std::min<uint64_t>(required,size));std::copy_n(bytes.begin(),count,out);} return required; }
XLPP_API int xlpp_workbook_save_vba_project(xlpp_workbook wb, const char* path) { clearError(); if(!wb||!path)return 0; try{WB(wb)->saveVbaProject(std::filesystem::path(path));return 1;}catch(const std::exception&e){setError(e);return 0;} }
XLPP_API int xlpp_workbook_has_vba_signature(xlpp_workbook wb) { return wb&&WB(wb)->hasVbaSignature()?1:0; }
XLPP_API int xlpp_workbook_vba_source_editable(xlpp_workbook wb) { return wb&&WB(wb)->vbaSourceEditable()?1:0; }
XLPP_API int xlpp_workbook_vba_project_name(xlpp_workbook wb, char* out, int size) { return wb?copyStrSized(WB(wb)->vbaProjectProperties().name,out,size):copyStrSized({},out,size); }
XLPP_API int xlpp_workbook_vba_project_description(xlpp_workbook wb, char* out, int size) { return wb?copyStrSized(WB(wb)->vbaProjectProperties().description,out,size):copyStrSized({},out,size); }
XLPP_API int xlpp_workbook_vba_project_help_file(xlpp_workbook wb, char* out, int size) { return wb?copyStrSized(WB(wb)->vbaProjectProperties().helpFile,out,size):copyStrSized({},out,size); }
XLPP_API uint32_t xlpp_workbook_vba_project_help_context(xlpp_workbook wb) { return wb?WB(wb)->vbaProjectProperties().helpContextId:0; }
XLPP_API int xlpp_workbook_vba_project_constants(xlpp_workbook wb, char* out, int size) { return wb?copyStrSized(WB(wb)->vbaProjectProperties().constants,out,size):copyStrSized({},out,size); }
XLPP_API int xlpp_workbook_set_vba_project_properties(xlpp_workbook wb, const char* name, const char* description, const char* helpFile, uint32_t helpContext, const char* constants) {
    clearError();
    if(!wb||!name)return 0;
    try{ xlpp::VbaProjectProperties properties; properties.name=name; properties.description=description?description:""; properties.helpFile=helpFile?helpFile:""; properties.helpContextId=helpContext; properties.constants=constants?constants:""; WB(wb)->setVbaProjectProperties(std::move(properties)); return 1; }
    catch(const std::exception&e){setError(e);return 0;}
}

XLPP_API int xlpp_inspect_office_encryption(const char* path, int* mode, uint64_t* key_bits, uint64_t* spin_count, char* cipher, int cipher_size, char* hash, int hash_size) {
    clearError();
    if (!path) { setError("Path is required"); return 0; }
    try {
        const auto info = xlpp::inspectOfficeEncryption(std::filesystem::path(path));
        if (mode) *mode = static_cast<int>(info.mode);
        if (key_bits) *key_bits = info.keyBits;
        if (spin_count) *spin_count = info.spinCount;
        copyStr(info.cipherAlgorithm, cipher, cipher_size);
        copyStr(info.hashAlgorithm, hash, hash_size);
        return info.encrypted ? 1 : 0;
    } catch (const std::exception& e) { setError(e); return 0; }
}

XLPP_API void xlpp_workbook_set_date1904(xlpp_workbook wb, int v) { WB(wb)->setDate1904(v != 0); }
XLPP_API int  xlpp_workbook_date1904(xlpp_workbook wb) { return WB(wb)->date1904() ? 1 : 0; }
XLPP_API void xlpp_workbook_clear(xlpp_workbook wb) { WB(wb)->clear(); }
XLPP_API int  xlpp_workbook_strict_namespaces(xlpp_workbook wb) { return WB(wb)->strictNamespaces() ? 1 : 0; }
XLPP_API uint64_t xlpp_workbook_diagnostic_warning_count(xlpp_workbook wb) { return wb ? static_cast<uint64_t>(WB(wb)->diagnostics().warnings.size()) : 0; }
XLPP_API uint64_t xlpp_workbook_diagnostic_error_count(xlpp_workbook wb) { return wb ? static_cast<uint64_t>(WB(wb)->diagnostics().errors.size()) : 0; }
XLPP_API int xlpp_workbook_diagnostic_warning(xlpp_workbook wb, uint64_t index, char* out, int size) {
    if (!wb || index >= WB(wb)->diagnostics().warnings.size()) return 0;
    return copyStrSized(WB(wb)->diagnostics().warnings[static_cast<std::size_t>(index)], out, size);
}
XLPP_API int xlpp_workbook_diagnostic_error(xlpp_workbook wb, uint64_t index, char* out, int size) {
    if (!wb || index >= WB(wb)->diagnostics().errors.size()) return 0;
    return copyStrSized(WB(wb)->diagnostics().errors[static_cast<std::size_t>(index)], out, size);
}

XLPP_API xlpp_properties xlpp_workbook_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_properties>(&WB(wb)->properties());
}
XLPP_API xlpp_wbprotection xlpp_workbook_protection(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_wbprotection>(&WB(wb)->protection());
}
XLPP_API xlpp_calcprops xlpp_workbook_calc_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_calcprops>(&WB(wb)->calcProperties());
}
XLPP_API xlpp_customprops xlpp_workbook_custom_properties(xlpp_workbook wb) {
    return reinterpret_cast<xlpp_customprops>(&WB(wb)->customProperties());
}

XLPP_API xlpp_namedstyle xlpp_workbook_add_named_style(xlpp_workbook wb, const char* name, int* ok) {
    try {
        auto& s = WB(wb)->addNamedStyle(xlpp::NamedStyle(name));
        if (ok) *ok = 1;
        return reinterpret_cast<xlpp_namedstyle>(&s);
    } catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_namedstyle xlpp_workbook_named_style(xlpp_workbook wb, const char* name) {
    return reinterpret_cast<xlpp_namedstyle>(WB(wb)->namedStyle(name));
}
XLPP_API int xlpp_workbook_named_styles_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->namedStyles().size());
}
XLPP_API xlpp_namedstyle xlpp_workbook_named_style_at(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_namedstyle>(const_cast<xlpp::NamedStyle*>(&WB(wb)->namedStyles()[static_cast<std::size_t>(index)]));
    } catch (...) { return nullptr; }
}
XLPP_API void xlpp_workbook_apply_named_style(xlpp_workbook wb, xlpp_cell c, const char* name) {
    try { WB(wb)->applyNamedStyle(*CELL(c), name); } catch (...) {}
}

XLPP_API xlpp_definedname xlpp_workbook_add_defined_name(xlpp_workbook wb, const char* name, const char* value, int* ok) {
    try {
        auto& d = WB(wb)->addDefinedName(xlpp::DefinedName(name, value));
        if (ok) *ok = 1;
        return reinterpret_cast<xlpp_definedname>(&d);
    } catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_definedname xlpp_workbook_add_defined_name_scoped(xlpp_workbook wb, const char* name, const char* value, uint64_t local_sheet_id, int* ok) {
    try {
        xlpp::DefinedName defined(name, value);
        defined.setLocalSheetId(static_cast<std::size_t>(local_sheet_id));
        auto& d = WB(wb)->addDefinedName(std::move(defined));
        if (ok) *ok = 1;
        return reinterpret_cast<xlpp_definedname>(&d);
    } catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_definedname xlpp_workbook_defined_name(xlpp_workbook wb, const char* name) {
    return reinterpret_cast<xlpp_definedname>(WB(wb)->definedName(name));
}
XLPP_API xlpp_definedname xlpp_workbook_defined_name_scoped(xlpp_workbook wb, const char* name, int has_local_sheet_id, uint64_t local_sheet_id) {
    if (!wb || !name) return nullptr;
    const std::optional<std::size_t> scope = has_local_sheet_id != 0
        ? std::optional<std::size_t>{static_cast<std::size_t>(local_sheet_id)}
        : std::nullopt;
    return reinterpret_cast<xlpp_definedname>(WB(wb)->definedName(name, scope));
}
XLPP_API int xlpp_workbook_defined_names_count(xlpp_workbook wb) {
    return static_cast<int>(WB(wb)->definedNames().size());
}
XLPP_API xlpp_definedname xlpp_workbook_defined_name_at(xlpp_workbook wb, int index) {
    try {
        return reinterpret_cast<xlpp_definedname>(&WB(wb)->definedNames()[static_cast<std::size_t>(index)]);
    } catch (...) { return nullptr; }
}
XLPP_API int xlpp_workbook_preserved_relationships_count(xlpp_workbook wb) {
    return wb ? static_cast<int>(WB(wb)->preservedRelationships().size()) : 0;
}
XLPP_API int xlpp_workbook_preserved_relationship_source_part(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedRelationships().at(static_cast<std::size_t>(index)).sourcePart, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_relationship_id(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedRelationships().at(static_cast<std::size_t>(index)).id, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_relationship_type(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedRelationships().at(static_cast<std::size_t>(index)).type, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_relationship_target(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedRelationships().at(static_cast<std::size_t>(index)).target, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_relationship_target_mode(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedRelationships().at(static_cast<std::size_t>(index)).targetMode, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_parts_count(xlpp_workbook wb) { return wb ? static_cast<int>(WB(wb)->preservedParts().size()) : 0; }
XLPP_API int xlpp_workbook_preserved_part_name(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).name, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API uint64_t xlpp_workbook_preserved_part_data_size(xlpp_workbook wb, int index) {
    try { return static_cast<uint64_t>(WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).data.size()); } catch (...) { return 0; }
}
XLPP_API uint64_t xlpp_workbook_preserved_part_data(xlpp_workbook wb, int index, unsigned char* out, uint64_t out_size) {
    try { const auto& data = WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).data; const auto n = (std::min)(data.size(), static_cast<std::size_t>(out_size)); if (out && n) std::memcpy(out, data.data(), n); return static_cast<uint64_t>(n); } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_part_override_type(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).overrideType, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_part_extension(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).extension, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_part_default_type(xlpp_workbook wb, int index, char* out, int out_size) {
    try { copyStr(WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).defaultType, out, out_size); return 1; } catch (...) { return 0; }
}
XLPP_API int xlpp_workbook_preserved_part_compress(xlpp_workbook wb, int index) {
    try { return WB(wb)->preservedParts().at(static_cast<std::size_t>(index)).compress ? 1 : 0; } catch (...) { return 0; }
}

// ============================================================
// Properties
// ============================================================
XLPP_API void xlpp_properties_set_title(xlpp_properties p, const char* v)    { PROP(p)->setTitle(v); }
XLPP_API void xlpp_properties_set_creator(xlpp_properties p, const char* v)   { PROP(p)->setCreator(v); }
XLPP_API void xlpp_properties_set_subject(xlpp_properties p, const char* v)   { PROP(p)->setSubject(v); }
XLPP_API void xlpp_properties_set_description(xlpp_properties p, const char* v) { PROP(p)->setDescription(v); }
XLPP_API void xlpp_properties_set_keywords(xlpp_properties p, const char* v)  { PROP(p)->setKeywords(v); }
XLPP_API void xlpp_properties_set_category(xlpp_properties p, const char* v)  { PROP(p)->setCategory(v); }
XLPP_API void xlpp_properties_set_last_modified_by(xlpp_properties p, const char* v) { PROP(p)->setLastModifiedBy(v); }
XLPP_API const char* xlpp_properties_get_title(xlpp_properties p)             { return PROP(p)->title().c_str(); }
XLPP_API const char* xlpp_properties_get_creator(xlpp_properties p)           { return PROP(p)->creator().c_str(); }
XLPP_API const char* xlpp_properties_get_subject(xlpp_properties p)           { return PROP(p)->subject().c_str(); }
XLPP_API const char* xlpp_properties_get_description(xlpp_properties p)       { return PROP(p)->description().c_str(); }
XLPP_API const char* xlpp_properties_get_keywords(xlpp_properties p)          { return PROP(p)->keywords().c_str(); }
XLPP_API const char* xlpp_properties_get_category(xlpp_properties p)          { return PROP(p)->category().c_str(); }
XLPP_API const char* xlpp_properties_get_last_modified_by(xlpp_properties p)  { return PROP(p)->lastModifiedBy().c_str(); }

// ============================================================
// Workbook protection
// ============================================================
#define WBP(h) reinterpret_cast<xlpp::WorkbookProtection*>(h)
XLPP_API void xlpp_wbprotection_set_lock_structure(xlpp_wbprotection p, int v) { WBP(p)->setLockStructure(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_structure(xlpp_wbprotection p) { return WBP(p)->lockStructure() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_lock_windows(xlpp_wbprotection p, int v) { WBP(p)->setLockWindows(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_windows(xlpp_wbprotection p) { return WBP(p)->lockWindows() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_lock_revision(xlpp_wbprotection p, int v) { WBP(p)->setLockRevision(v != 0); }
XLPP_API int  xlpp_wbprotection_lock_revision(xlpp_wbprotection p) { return WBP(p)->lockRevision() ? 1 : 0; }
XLPP_API void xlpp_wbprotection_set_password_hash(xlpp_wbprotection p, const char* v) { WBP(p)->setWorkbookPasswordHash(v); }
XLPP_API const char* xlpp_wbprotection_password_hash(xlpp_wbprotection p) { return WBP(p)->workbookPasswordHash().c_str(); }

// ============================================================
// CalcProperties
// ============================================================
#define CP(h) reinterpret_cast<xlpp::CalcProperties*>(h)
XLPP_API void xlpp_calcprops_set_calc_id(xlpp_calcprops p, int v) { CP(p)->setCalcId(v); }
XLPP_API int  xlpp_calcprops_calc_id(xlpp_calcprops p) { return CP(p)->calcId(); }
XLPP_API void xlpp_calcprops_set_calc_mode(xlpp_calcprops p, const char* v) { CP(p)->setCalcMode(v); }
XLPP_API const char* xlpp_calcprops_calc_mode(xlpp_calcprops p) { return CP(p)->calcMode().c_str(); }
XLPP_API void xlpp_calcprops_set_calc_on_save(xlpp_calcprops p, int v) { CP(p)->setCalcOnSave(v != 0); }
XLPP_API int  xlpp_calcprops_calc_on_save(xlpp_calcprops p) { return CP(p)->calcOnSave() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_full_calc_on_load(xlpp_calcprops p, int v) { CP(p)->setFullCalcOnLoad(v != 0); }
XLPP_API int  xlpp_calcprops_full_calc_on_load(xlpp_calcprops p) { return CP(p)->fullCalcOnLoad() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_full_precision(xlpp_calcprops p, int v) { CP(p)->setFullPrecision(v != 0); }
XLPP_API int  xlpp_calcprops_full_precision(xlpp_calcprops p) { return CP(p)->fullPrecision() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_iterate(xlpp_calcprops p, int v) { CP(p)->setIterate(v != 0); }
XLPP_API int  xlpp_calcprops_iterate(xlpp_calcprops p) { return CP(p)->iterate() ? 1 : 0; }
XLPP_API void xlpp_calcprops_set_iterate_count(xlpp_calcprops p, int v) { CP(p)->setIterateCount(v); }
XLPP_API int  xlpp_calcprops_iterate_count(xlpp_calcprops p) { return CP(p)->iterateCount(); }
XLPP_API void xlpp_calcprops_set_iterate_delta(xlpp_calcprops p, double v) { CP(p)->setIterateDelta(v); }
XLPP_API double xlpp_calcprops_iterate_delta(xlpp_calcprops p) { return CP(p)->iterateDelta(); }

// ============================================================
// Custom properties
// ============================================================
#define CUST(h) reinterpret_cast<xlpp::CustomProperties*>(h)
#define CUSTP(h) reinterpret_cast<xlpp::CustomProperty*>(h)
XLPP_API xlpp_customprop xlpp_customprops_add(xlpp_customprops c, const char* name, const char* value, int type) {
    try {
        xlpp::CustomProperty p;
        const std::string vstr = value ? value : "";
        switch (type) {
        case 0: p = xlpp::CustomProperty(name, vstr); break;
        case 1: p = xlpp::CustomProperty(name, vstr.empty() ? 0 : std::atoi(vstr.c_str())); break;
        case 2: p = xlpp::CustomProperty(name, vstr.empty() ? 0.0 : std::atof(vstr.c_str())); break;
        case 3: p = xlpp::CustomProperty(name, vstr == "true"); break;
        default: p = xlpp::CustomProperty(name, vstr); break;
        }
        CUST(c)->add(std::move(p));
        return reinterpret_cast<xlpp_customprop>(&CUST(c)->items().back());
    } catch (...) { return nullptr; }
}
XLPP_API int xlpp_customprops_count(xlpp_customprops c) { return static_cast<int>(CUST(c)->items().size()); }
XLPP_API xlpp_customprop xlpp_customprops_at(xlpp_customprops c, int index) {
    try { return reinterpret_cast<xlpp_customprop>(&CUST(c)->items()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API const char* xlpp_customprop_name(xlpp_customprop p) { return CUSTP(p)->name().c_str(); }
XLPP_API const char* xlpp_customprop_value(xlpp_customprop p) { return CUSTP(p)->value().c_str(); }
XLPP_API const char* xlpp_customprop_type(xlpp_customprop p) { return CUSTP(p)->type().c_str(); }

// ============================================================
// Worksheet
// ============================================================
XLPP_API const char* xlpp_sheet_name(xlpp_worksheet ws) { return WS(ws)->name().c_str(); }
XLPP_API void xlpp_sheet_rename(xlpp_worksheet ws, const char* name) { WS(ws)->rename(name); }
XLPP_API int xlpp_sheet_vba_code_name(xlpp_worksheet ws, char* out, int size) { return ws?copyStrSized(WS(ws)->vbaCodeName(),out,size):copyStrSized({},out,size); }
XLPP_API int xlpp_sheet_set_vba_code_name(xlpp_worksheet ws, const char* codeName) { clearError(); if(!ws||!codeName)return 0; try{WS(ws)->setVbaCodeName(codeName);return 1;}catch(const std::exception&e){setError(e);return 0;} }

XLPP_API xlpp_cell xlpp_sheet_cell(xlpp_worksheet ws, const char* address) {
    try { return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(address)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_cell xlpp_sheet_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col) {
    try { return reinterpret_cast<xlpp_cell>(&WS(ws)->cell(static_cast<std::size_t>(row), static_cast<std::size_t>(col))); } catch (...) { return nullptr; }
}
XLPP_API int xlpp_sheet_has_cell(xlpp_worksheet ws, const char* address) {
    return WS(ws)->tryCell(address) != nullptr ? 1 : 0;
}
XLPP_API int xlpp_sheet_has_cell_rc(xlpp_worksheet ws, uint64_t row, uint64_t col) {
    return WS(ws)->tryCell(static_cast<std::size_t>(row), static_cast<std::size_t>(col)) != nullptr ? 1 : 0;
}

XLPP_API void xlpp_sheet_append_row(xlpp_worksheet ws, const char** values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        if (values[i] && values[i][0]) cv.push_back(std::string(values[i]));
        else cv.push_back(std::monostate{});
    }
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_append_doubles(xlpp_worksheet ws, const double* values, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) cv.push_back(values[i]);
    WS(ws)->append(cv);
}

XLPP_API void xlpp_sheet_append_values(xlpp_worksheet ws, const double* nums, const int* types, int count) {
    std::vector<xlpp::CellValue> cv;
    cv.reserve(static_cast<std::size_t>(count));
    for (int i = 0; i < count; ++i) {
        switch (types[i]) {
        case XLPP_VALUE_NUMBER: cv.push_back(nums[i]); break;
        case XLPP_VALUE_BOOL: cv.push_back(nums[i] != 0.0); break;
        case XLPP_VALUE_STRING: cv.push_back(std::string(reinterpret_cast<const char*>(static_cast<uintptr_t>(static_cast<uint64_t>(nums[i]))))); break;
        default: cv.push_back(std::monostate{}); break;
        }
    }
    WS(ws)->append(cv);
}

XLPP_API xlpp_cellrange xlpp_sheet_range(xlpp_worksheet ws, const char* address) {
    try {
        auto r = WS(ws)->range(address);
        return reinterpret_cast<xlpp_cellrange>(new xlpp::CellRange(std::move(r)));
    } catch (...) { return nullptr; }
}
XLPP_API xlpp_cellrange xlpp_sheet_range_rc(xlpp_worksheet ws, uint64_t minRow, uint64_t minCol, uint64_t maxRow, uint64_t maxCol) {
    try {
        auto r = WS(ws)->range(static_cast<std::size_t>(minRow), static_cast<std::size_t>(minCol),
                               static_cast<std::size_t>(maxRow), static_cast<std::size_t>(maxCol));
        return reinterpret_cast<xlpp_cellrange>(new xlpp::CellRange(std::move(r)));
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_merge_cells(xlpp_worksheet ws, const char* range)     { try { WS(ws)->mergeCells(range); } catch (...) {} }
XLPP_API void xlpp_sheet_unmerge_cells(xlpp_worksheet ws, const char* range)   { try { WS(ws)->unmergeCells(range); } catch (...) {} }
XLPP_API int xlpp_sheet_is_merged(xlpp_worksheet ws, const char* cell)         { try { return WS(ws)->isMerged(cell) ? 1 : 0; } catch (...) { return 0; } }
XLPP_API int xlpp_sheet_merged_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->mergedRanges().size()); }
XLPP_API void xlpp_sheet_merged_at(xlpp_worksheet ws, int index, char* out, int outSize) {
    try { copyStr(WS(ws)->mergedRanges()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}

XLPP_API void xlpp_sheet_freeze_panes(xlpp_worksheet ws, const char* cell)     { try { WS(ws)->freezePanes(cell); } catch (...) {} }
XLPP_API void xlpp_sheet_clear_freeze_panes(xlpp_worksheet ws) { WS(ws)->clearFreezePanes(); }
XLPP_API void xlpp_sheet_frozen_pane(xlpp_worksheet ws, char* out, int outSize) {
    const auto& fp = WS(ws)->frozenPane();
    if (fp) copyStr(*fp, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}

XLPP_API xlpp_rowdim xlpp_sheet_row_dimension(xlpp_worksheet ws, uint64_t row) {
    return reinterpret_cast<xlpp_rowdim>(&WS(ws)->rowDimension(static_cast<std::size_t>(row)));
}
XLPP_API xlpp_coldim xlpp_sheet_col_dimension(xlpp_worksheet ws, uint64_t col) {
    return reinterpret_cast<xlpp_coldim>(&WS(ws)->columnDimension(static_cast<std::size_t>(col)));
}

XLPP_API uint64_t xlpp_sheet_max_row(xlpp_worksheet ws) { return WS(ws)->maxRow(); }
XLPP_API uint64_t xlpp_sheet_max_col(xlpp_worksheet ws) { return WS(ws)->maxColumn(); }
XLPP_API void xlpp_sheet_dimensions(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->dimensions(), out, outSize); }
XLPP_API int xlpp_sheet_empty(xlpp_worksheet ws) { return WS(ws)->empty() ? 1 : 0; }
XLPP_API uint64_t xlpp_sheet_row_count(xlpp_worksheet ws) { return WS(ws)->rowCount(); }
XLPP_API uint64_t xlpp_sheet_col_count(xlpp_worksheet ws) { return WS(ws)->columnCount(); }

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

XLPP_API void xlpp_sheet_set_print_area(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintArea(v); }
XLPP_API void xlpp_sheet_print_area(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printArea(), out, outSize); }
XLPP_API void xlpp_sheet_set_print_titles_rows(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintTitlesRows(v); }
XLPP_API void xlpp_sheet_print_titles_rows(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printTitlesRows(), out, outSize); }
XLPP_API void xlpp_sheet_set_print_titles_cols(xlpp_worksheet ws, const char* v) { WS(ws)->setPrintTitlesCols(v); }
XLPP_API void xlpp_sheet_print_titles_cols(xlpp_worksheet ws, char* out, int outSize) { copyStr(WS(ws)->printTitlesCols(), out, outSize); }

XLPP_API xlpp_autofilter xlpp_sheet_auto_filter(xlpp_worksheet ws) { return reinterpret_cast<xlpp_autofilter>(&WS(ws)->autoFilter()); }
XLPP_API xlpp_cfcollection xlpp_sheet_conditional_formatting(xlpp_worksheet ws) { return reinterpret_cast<xlpp_cfcollection>(&WS(ws)->conditionalFormatting()); }
XLPP_API xlpp_dvcollection xlpp_sheet_data_validations(xlpp_worksheet ws) { return reinterpret_cast<xlpp_dvcollection>(&WS(ws)->dataValidations()); }
XLPP_API xlpp_pagesetup xlpp_sheet_page_setup(xlpp_worksheet ws) { return reinterpret_cast<xlpp_pagesetup>(&WS(ws)->pageSetup()); }
XLPP_API xlpp_pagemargins xlpp_sheet_page_margins(xlpp_worksheet ws) { return reinterpret_cast<xlpp_pagemargins>(&WS(ws)->pageMargins()); }
XLPP_API xlpp_printopts xlpp_sheet_print_options(xlpp_worksheet ws) { return reinterpret_cast<xlpp_printopts>(&WS(ws)->printOptions()); }
XLPP_API xlpp_headerfooter xlpp_sheet_header_footer(xlpp_worksheet ws) { return reinterpret_cast<xlpp_headerfooter>(&WS(ws)->headerFooter()); }
XLPP_API xlpp_wssprotection xlpp_sheet_protection(xlpp_worksheet ws) { return reinterpret_cast<xlpp_wssprotection>(&WS(ws)->protection()); }
XLPP_API xlpp_sheetview xlpp_sheet_view(xlpp_worksheet ws) { return reinterpret_cast<xlpp_sheetview>(&WS(ws)->sheetView()); }

XLPP_API xlpp_table xlpp_sheet_add_table(xlpp_worksheet ws, const char* name, const char* reference, int* ok) {
    try { auto& t = WS(ws)->addTable(name, reference); if (ok) *ok = 1; return reinterpret_cast<xlpp_table>(&t); }
    catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API xlpp_table xlpp_sheet_table(xlpp_worksheet ws, const char* name) { return reinterpret_cast<xlpp_table>(WS(ws)->table(name)); }
XLPP_API int xlpp_sheet_table_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->tables().size()); }
XLPP_API xlpp_table xlpp_sheet_table_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_table>(&WS(ws)->tables()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API xlpp_image xlpp_sheet_add_image(xlpp_worksheet ws, const char* path, const char* anchor, int* ok) {
    try { auto& img = WS(ws)->addImage(std::filesystem::path(path), anchor); if (ok) *ok = 1; return reinterpret_cast<xlpp_image>(&img); }
    catch (...) { if (ok) *ok = 0; return nullptr; }
}
XLPP_API int xlpp_sheet_image_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->images().size()); }
XLPP_API xlpp_image xlpp_sheet_image_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_image>(&WS(ws)->images()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_add_chart(xlpp_worksheet ws, int type) {
    WS(ws)->addChart(xlpp::Chart(static_cast<xlpp::Chart::Type>(type)));
}
XLPP_API int xlpp_sheet_chart_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->chartCount()); }
XLPP_API xlpp_chart xlpp_sheet_chart_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_chart>(&WS(ws)->chart(static_cast<std::size_t>(index))); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_sheet_add_pivot(xlpp_worksheet ws, const char* name, const char* location) {
    xlpp::PivotTable pt(name);
    pt.setLocation(location);
    WS(ws)->addPivotTable(std::move(pt));
}
XLPP_API int xlpp_sheet_pivot_count(xlpp_worksheet ws) { return static_cast<int>(WS(ws)->pivotTables().size()); }
XLPP_API xlpp_pivottable xlpp_sheet_pivot_at(xlpp_worksheet ws, int index) {
    try { return reinterpret_cast<xlpp_pivottable>(&WS(ws)->pivotTables()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

// ============================================================
// Row / Column dimensions
// ============================================================
#define RD(h) reinterpret_cast<xlpp::RowDimension*>(h)
#define CD(h) reinterpret_cast<xlpp::ColumnDimension*>(h)
XLPP_API void xlpp_rowdim_set_height(xlpp_rowdim d, double v) { RD(d)->height = v; }
XLPP_API int  xlpp_rowdim_has_height(xlpp_rowdim d) { return RD(d)->height.has_value() ? 1 : 0; }
XLPP_API double xlpp_rowdim_height(xlpp_rowdim d) { return RD(d)->height.value_or(0.0); }
XLPP_API void xlpp_rowdim_set_hidden(xlpp_rowdim d, int v) { RD(d)->hidden = v != 0; }
XLPP_API int  xlpp_rowdim_hidden(xlpp_rowdim d) { return RD(d)->hidden ? 1 : 0; }
XLPP_API void xlpp_rowdim_set_outline_level(xlpp_rowdim d, int v) { RD(d)->outlineLevel = v; }
XLPP_API int  xlpp_rowdim_outline_level(xlpp_rowdim d) { return RD(d)->outlineLevel; }
XLPP_API void xlpp_rowdim_set_collapsed(xlpp_rowdim d, int v) { RD(d)->collapsed = v != 0; }
XLPP_API int  xlpp_rowdim_collapsed(xlpp_rowdim d) { return RD(d)->collapsed ? 1 : 0; }

XLPP_API void xlpp_coldim_set_width(xlpp_coldim d, double v) { CD(d)->width = v; }
XLPP_API int  xlpp_coldim_has_width(xlpp_coldim d) { return CD(d)->width.has_value() ? 1 : 0; }
XLPP_API double xlpp_coldim_width(xlpp_coldim d) { return CD(d)->width.value_or(0.0); }
XLPP_API void xlpp_coldim_set_hidden(xlpp_coldim d, int v) { CD(d)->hidden = v != 0; }
XLPP_API int  xlpp_coldim_hidden(xlpp_coldim d) { return CD(d)->hidden ? 1 : 0; }
XLPP_API void xlpp_coldim_set_best_fit(xlpp_coldim d, int v) { CD(d)->bestFit = v != 0; }
XLPP_API int  xlpp_coldim_best_fit(xlpp_coldim d) { return CD(d)->bestFit ? 1 : 0; }
XLPP_API void xlpp_coldim_set_outline_level(xlpp_coldim d, int v) { CD(d)->outlineLevel = v; }
XLPP_API int  xlpp_coldim_outline_level(xlpp_coldim d) { return CD(d)->outlineLevel; }
XLPP_API void xlpp_coldim_set_collapsed(xlpp_coldim d, int v) { CD(d)->collapsed = v != 0; }
XLPP_API int  xlpp_coldim_collapsed(xlpp_coldim d) { return CD(d)->collapsed ? 1 : 0; }

// ============================================================
// CellRange
// ============================================================
#define RNG(h) reinterpret_cast<xlpp::CellRange*>(h)
XLPP_API uint64_t xlpp_range_min_row(xlpp_cellrange r) { return RNG(r)->minRow(); }
XLPP_API uint64_t xlpp_range_min_col(xlpp_cellrange r) { return RNG(r)->minColumn(); }
XLPP_API uint64_t xlpp_range_max_row(xlpp_cellrange r) { return RNG(r)->maxRow(); }
XLPP_API uint64_t xlpp_range_max_col(xlpp_cellrange r) { return RNG(r)->maxColumn(); }
XLPP_API uint64_t xlpp_range_row_count(xlpp_cellrange r) { return RNG(r)->rowCount(); }
XLPP_API uint64_t xlpp_range_col_count(xlpp_cellrange r) { return RNG(r)->columnCount(); }
XLPP_API void xlpp_range_address(xlpp_cellrange r, char* out, int outSize) { copyStr(RNG(r)->address(), out, outSize); }
XLPP_API xlpp_cell xlpp_range_cell(xlpp_cellrange r, uint64_t relRow, uint64_t relCol) {
    return reinterpret_cast<xlpp_cell>(&RNG(r)->cell(static_cast<std::size_t>(relRow), static_cast<std::size_t>(relCol)));
}
XLPP_API void xlpp_range_set_value(xlpp_cellrange r, double v) { RNG(r)->setValue(v); }
XLPP_API void xlpp_range_set_string(xlpp_cellrange r, const char* v) { RNG(r)->setValue(std::string(v)); }
XLPP_API void xlpp_range_clear(xlpp_cellrange r) { RNG(r)->clear(); }
XLPP_API void xlpp_range_values(xlpp_cellrange r, double* out, int* outCount) {
    auto values = RNG(r)->values();
    if (outCount) *outCount = static_cast<int>(values.size());
    if (out) {
        for (std::size_t i = 0; i < values.size() && out; ++i)
            out[i] = std::holds_alternative<double>(values[i]) ? std::get<double>(values[i]) : 0.0;
    }
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
    try { return std::get<bool>(CELL(c)->value()) ? 1 : 0; } catch (...) { return 0; }
}
XLPP_API double xlpp_cell_get_number(xlpp_cell c) {
    if (auto* v = std::get_if<double>(&CELL(c)->value())) return *v;
    if (auto* v = std::get_if<xlpp::DateTime>(&CELL(c)->value())) return xlpp::toExcelSerial(*v, false);
    return 0.0;
}
XLPP_API const char* xlpp_cell_get_string(xlpp_cell c) {
    if (auto* v = std::get_if<std::string>(&CELL(c)->value())) return v->c_str();
    return "";
}
XLPP_API int xlpp_cell_is_empty(xlpp_cell c) { return CELL(c)->empty() ? 1 : 0; }
XLPP_API int xlpp_cell_has_value(xlpp_cell c) { return CELL(c)->hasValue() ? 1 : 0; }
XLPP_API int xlpp_cell_is_numeric(xlpp_cell c) { return CELL(c)->isNumeric() ? 1 : 0; }
XLPP_API int xlpp_cell_is_string(xlpp_cell c) { return CELL(c)->isString() ? 1 : 0; }
XLPP_API int xlpp_cell_is_bool(xlpp_cell c) { return CELL(c)->isBoolean() ? 1 : 0; }
XLPP_API int xlpp_cell_is_date(xlpp_cell c) { return CELL(c)->isDate() ? 1 : 0; }
XLPP_API int xlpp_cell_is_error(xlpp_cell c) { return CELL(c)->isError() ? 1 : 0; }
XLPP_API int xlpp_cell_error_code(xlpp_cell c) {
    auto e = CELL(c)->error();
    if (!e) return -1;
    switch (*e) {
    case xlpp::CellError::Null: return XLPP_ERROR_NULL;
    case xlpp::CellError::DivisionByZero: return XLPP_ERROR_DIV0;
    case xlpp::CellError::Value: return XLPP_ERROR_VALUE;
    case xlpp::CellError::Reference: return XLPP_ERROR_REF;
    case xlpp::CellError::Name: return XLPP_ERROR_NAME;
    case xlpp::CellError::Number: return XLPP_ERROR_NUM;
    case xlpp::CellError::NotAvailable: return XLPP_ERROR_NA;
    case xlpp::CellError::GettingData: return XLPP_ERROR_GETTING_DATA;
    case xlpp::CellError::Spill: return XLPP_ERROR_SPILL;
    case xlpp::CellError::Calculation: return XLPP_ERROR_CALC;
    }
    return XLPP_ERROR_VALUE;
}
XLPP_API int xlpp_cell_date(xlpp_cell c, int* year, int* month, int* day, int* hour, int* minute, double* second) {
    auto d = CELL(c)->date();
    if (!d) return 0;
    if (year) *year = d->year;
    if (month) *month = d->month;
    if (day) *day = d->day;
    if (hour) *hour = d->hour;
    if (minute) *minute = d->minute;
    if (second) *second = d->second;
    return 1;
}

XLPP_API void xlpp_cell_set_string(xlpp_cell c, const char* v)   { CELL(c)->setStringValue(v ? v : ""); }
XLPP_API void xlpp_cell_set_number(xlpp_cell c, double v)        { CELL(c)->setNumericValue(v); }
XLPP_API void xlpp_cell_set_bool(xlpp_cell c, int v)             { CELL(c)->setBoolValue(v != 0); }
XLPP_API void xlpp_cell_set_empty(xlpp_cell c)                   { CELL(c)->setValue(std::monostate{}); }
XLPP_API void xlpp_cell_set_error(xlpp_cell c, int errorCode) {
    switch (errorCode) {
    case XLPP_ERROR_NULL: CELL(c)->setError(xlpp::CellError::Null); break;
    case XLPP_ERROR_DIV0: CELL(c)->setError(xlpp::CellError::DivisionByZero); break;
    case XLPP_ERROR_REF: CELL(c)->setError(xlpp::CellError::Reference); break;
    case XLPP_ERROR_NAME: CELL(c)->setError(xlpp::CellError::Name); break;
    case XLPP_ERROR_NUM: CELL(c)->setError(xlpp::CellError::Number); break;
    case XLPP_ERROR_NA: CELL(c)->setError(xlpp::CellError::NotAvailable); break;
    case XLPP_ERROR_GETTING_DATA: CELL(c)->setError(xlpp::CellError::GettingData); break;
    case XLPP_ERROR_SPILL: CELL(c)->setError(xlpp::CellError::Spill); break;
    case XLPP_ERROR_CALC: CELL(c)->setError(xlpp::CellError::Calculation); break;
    default: CELL(c)->setError(xlpp::CellError::Value); break;
    }
}
XLPP_API void xlpp_cell_set_date(xlpp_cell c, int year, int month, int day, int hour, int minute, double second, int hasTime) {
    xlpp::DateTime d{year, month, day, hour, minute, second};
    if (hasTime) CELL(c)->setDateTime(d);
    else CELL(c)->setDate(d);
}
XLPP_API void xlpp_cell_clear(xlpp_cell c)                       { CELL(c)->clear(); }

XLPP_API const char* xlpp_cell_get_formula(xlpp_cell c)          { return CELL(c)->formula().c_str(); }
XLPP_API void xlpp_cell_set_formula(xlpp_cell c, const char* f)   { CELL(c)->setFormula(f ? f : ""); }
XLPP_API void xlpp_cell_set_shared_formula(xlpp_cell c, const char* f, unsigned sharedIndex, const char* reference) {
    CELL(c)->setSharedFormula(f ? f : "", sharedIndex, reference ? reference : "");
}
XLPP_API void xlpp_cell_set_array_formula(xlpp_cell c, const char* f, const char* reference) {
    CELL(c)->setArrayFormula(f ? f : "", reference ? reference : "");
}
XLPP_API void xlpp_cell_set_dynamic_array_formula(xlpp_cell c, const char* f, const char* reference) {
    CELL(c)->setDynamicArrayFormula(f ? f : "", reference ? reference : "");
}
XLPP_API int xlpp_cell_has_formula(xlpp_cell c)                  { return CELL(c)->hasFormula() ? 1 : 0; }
XLPP_API void xlpp_cell_clear_formula(xlpp_cell c)               { CELL(c)->clearFormula(); }

XLPP_API xlpp_style xlpp_cell_style(xlpp_cell c)         { return reinterpret_cast<xlpp_style>(&CELL(c)->style()); }
XLPP_API xlpp_font xlpp_cell_font(xlpp_cell c)           { return reinterpret_cast<xlpp_font>(&CELL(c)->font()); }
XLPP_API xlpp_fill xlpp_cell_fill(xlpp_cell c)           { return reinterpret_cast<xlpp_fill>(&CELL(c)->fill()); }
XLPP_API xlpp_border xlpp_cell_border(xlpp_cell c)       { return reinterpret_cast<xlpp_border>(&CELL(c)->border()); }
XLPP_API xlpp_alignment xlpp_cell_alignment(xlpp_cell c)  { return reinterpret_cast<xlpp_alignment>(&CELL(c)->alignment()); }
XLPP_API void xlpp_cell_set_number_format(xlpp_cell c, const char* v) { CELL(c)->setNumberFormat(v); }
XLPP_API void xlpp_cell_number_format(xlpp_cell c, char* out, int outSize) { copyStr(CELL(c)->numberFormat(), out, outSize); }
XLPP_API void xlpp_cell_set_named_style(xlpp_cell c, const char* name) { CELL(c)->setNamedStyle(name); }
XLPP_API void xlpp_cell_named_style(xlpp_cell c, char* out, int outSize) {
    const auto& ns = CELL(c)->namedStyle();
    if (ns) copyStr(*ns, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}

XLPP_API int xlpp_cell_has_hyperlink(xlpp_cell c) { return CELL(c)->hasHyperlink() ? 1 : 0; }
XLPP_API xlpp_hyperlink xlpp_cell_hyperlink(xlpp_cell c) { return reinterpret_cast<xlpp_hyperlink>(&CELL(c)->hyperlink()); }
XLPP_API void xlpp_cell_set_hyperlink(xlpp_cell c, const char* url) {
    CELL(c)->setHyperlink(xlpp::Hyperlink(url ? url : ""));
}
XLPP_API void xlpp_cell_set_hyperlink_full(xlpp_cell c, const char* url, const char* display, const char* tooltip, int external) {
    xlpp::Hyperlink h(url ? url : "");
    if (display) h.setDisplay(display);
    if (tooltip) h.setTooltip(tooltip);
    h.setExternal(external != 0);
    CELL(c)->setHyperlink(std::move(h));
}
XLPP_API void xlpp_cell_clear_hyperlink(xlpp_cell c) { CELL(c)->clearHyperlink(); }

XLPP_API int xlpp_cell_has_comment(xlpp_cell c) { return CELL(c)->hasComment() ? 1 : 0; }
XLPP_API xlpp_comment xlpp_cell_comment(xlpp_cell c) { return reinterpret_cast<xlpp_comment>(&CELL(c)->comment()); }
XLPP_API void xlpp_cell_set_comment(xlpp_cell c, const char* text, const char* author) {
    CELL(c)->setComment(xlpp::Comment(text ? text : "", author ? author : ""));
}
XLPP_API void xlpp_cell_clear_comment(xlpp_cell c) { CELL(c)->clearComment(); }

// ============================================================
// Hyperlink / Comment
// ============================================================
#define HL(h) reinterpret_cast<xlpp::Hyperlink*>(h)
#define CM(h) reinterpret_cast<xlpp::Comment*>(h)
XLPP_API void xlpp_hyperlink_set_target(xlpp_hyperlink h, const char* v) { HL(h)->setTarget(v); }
XLPP_API void xlpp_hyperlink_set_display(xlpp_hyperlink h, const char* v) { HL(h)->setDisplay(v); }
XLPP_API void xlpp_hyperlink_set_tooltip(xlpp_hyperlink h, const char* v) { HL(h)->setTooltip(v); }
XLPP_API void xlpp_hyperlink_set_external(xlpp_hyperlink h, int v) { HL(h)->setExternal(v != 0); }
XLPP_API const char* xlpp_hyperlink_target(xlpp_hyperlink h) { return HL(h)->target().c_str(); }
XLPP_API const char* xlpp_hyperlink_display(xlpp_hyperlink h) { return HL(h)->display().c_str(); }
XLPP_API const char* xlpp_hyperlink_tooltip(xlpp_hyperlink h) { return HL(h)->tooltip().c_str(); }
XLPP_API int xlpp_hyperlink_external(xlpp_hyperlink h) { return HL(h)->external() ? 1 : 0; }

XLPP_API void xlpp_comment_set_text(xlpp_comment c, const char* v) { CM(c)->setText(v); }
XLPP_API void xlpp_comment_set_author(xlpp_comment c, const char* v) { CM(c)->setAuthor(v); }
XLPP_API const char* xlpp_comment_text(xlpp_comment c) { return CM(c)->text().c_str(); }
XLPP_API const char* xlpp_comment_author(xlpp_comment c) { return CM(c)->author().c_str(); }

// ============================================================
// RichText
// ============================================================
#define RT(h) reinterpret_cast<xlpp::RichText*>(h)
#define RTR(h) reinterpret_cast<xlpp::RichTextRun*>(h)
XLPP_API xlpp_richtext xlpp_richtext_create(void) {
    return reinterpret_cast<xlpp_richtext>(new xlpp::RichText());
}
XLPP_API void xlpp_richtext_destroy(xlpp_richtext rt) { delete RT(rt); }
XLPP_API int xlpp_richtext_run_count(xlpp_richtext rt) { return static_cast<int>(RT(rt)->runs().size()); }
XLPP_API xlpp_richtextrun xlpp_richtext_add_run(xlpp_richtext rt, const char* text) {
    RT(rt)->addRun(xlpp::RichTextRun(text ? text : ""));
    return reinterpret_cast<xlpp_richtextrun>(&RT(rt)->runs().back());
}
XLPP_API xlpp_richtextrun xlpp_richtext_run_at(xlpp_richtext rt, int index) {
    try { return reinterpret_cast<xlpp_richtextrun>(&RT(rt)->runs()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API int xlpp_richtext_empty(xlpp_richtext rt) { return RT(rt)->empty() ? 1 : 0; }
XLPP_API void xlpp_richtext_plain_text(xlpp_richtext rt, char* out, int outSize) { copyStr(RT(rt)->plainText(), out, outSize); }

XLPP_API void xlpp_richtextrun_set_text(xlpp_richtextrun r, const char* v) { RTR(r)->setText(v); }
XLPP_API void xlpp_richtextrun_set_bold(xlpp_richtextrun r, int v) { RTR(r)->setBold(v != 0); }
XLPP_API void xlpp_richtextrun_set_italic(xlpp_richtextrun r, int v) { RTR(r)->setItalic(v != 0); }
XLPP_API void xlpp_richtextrun_set_underline(xlpp_richtextrun r, int v) { RTR(r)->setUnderline(v != 0); }
XLPP_API void xlpp_richtextrun_set_strike(xlpp_richtextrun r, int v) { RTR(r)->setStrike(v != 0); }
XLPP_API void xlpp_richtextrun_set_color(xlpp_richtextrun r, const char* v) { RTR(r)->setColor(v); }
XLPP_API void xlpp_richtextrun_set_size(xlpp_richtextrun r, double v) { RTR(r)->setSize(v); }
XLPP_API void xlpp_richtextrun_set_font(xlpp_richtextrun r, const char* v) { RTR(r)->setFontName(v); }
XLPP_API const char* xlpp_richtextrun_text(xlpp_richtextrun r) { return RTR(r)->text().c_str(); }

// ============================================================
// Named style / Defined name
// ============================================================
#define NS(h) reinterpret_cast<xlpp::NamedStyle*>(h)
#define DN(h) reinterpret_cast<xlpp::DefinedName*>(h)
XLPP_API void xlpp_namedstyle_set_name(xlpp_namedstyle s, const char* v) { NS(s)->setName(v); }
XLPP_API const char* xlpp_namedstyle_name(xlpp_namedstyle s) { return NS(s)->name().c_str(); }
XLPP_API xlpp_style xlpp_namedstyle_style(xlpp_namedstyle s) { return reinterpret_cast<xlpp_style>(&NS(s)->style()); }

XLPP_API void xlpp_definedname_set_value(xlpp_definedname d, const char* v) { try { DN(d)->setValue(v); } catch (...) {} }
XLPP_API const char* xlpp_definedname_name(xlpp_definedname d) { return DN(d)->name().c_str(); }
XLPP_API const char* xlpp_definedname_value(xlpp_definedname d) { return DN(d)->value().c_str(); }
XLPP_API void xlpp_definedname_set_local_sheet_id(xlpp_definedname d, uint64_t v) { DN(d)->setLocalSheetId(static_cast<std::size_t>(v)); }
XLPP_API void xlpp_definedname_clear_local_sheet_id(xlpp_definedname d) { DN(d)->clearLocalSheetId(); }
XLPP_API int xlpp_definedname_has_local_sheet_id(xlpp_definedname d) { return DN(d)->localSheetId().has_value() ? 1 : 0; }
XLPP_API uint64_t xlpp_definedname_local_sheet_id(xlpp_definedname d, int* has_value) {
    if (!d) { if (has_value) *has_value = 0; return 0; }
    const auto& value = DN(d)->localSheetId();
    if (has_value) *has_value = value.has_value() ? 1 : 0;
    return value ? static_cast<uint64_t>(*value) : 0;
}
XLPP_API void xlpp_definedname_set_hidden(xlpp_definedname d, int v) { DN(d)->setHidden(v != 0); }
XLPP_API int xlpp_definedname_hidden(xlpp_definedname d) { return DN(d)->hidden() ? 1 : 0; }
XLPP_API void xlpp_definedname_set_comment(xlpp_definedname d, const char* v) { DN(d)->setComment(v); }
XLPP_API const char* xlpp_definedname_comment(xlpp_definedname d) { return DN(d)->comment().c_str(); }

// ============================================================
// AutoFilter / Sort
// ============================================================
#define AF(h) reinterpret_cast<xlpp::AutoFilter*>(h)
#define FC(h) reinterpret_cast<xlpp::FilterColumn*>(h)
#define SS(h) reinterpret_cast<xlpp::SortState*>(h)
XLPP_API void xlpp_autofilter_set_reference(xlpp_autofilter f, const char* v) { AF(f)->setReference(v); }
XLPP_API void xlpp_autofilter_reference(xlpp_autofilter f, char* out, int outSize) { copyStr(AF(f)->reference(), out, outSize); }
XLPP_API int xlpp_autofilter_enabled(xlpp_autofilter f) { return AF(f)->enabled() ? 1 : 0; }
XLPP_API void xlpp_autofilter_clear(xlpp_autofilter f) { AF(f)->clear(); }
XLPP_API xlpp_filtercol xlpp_autofilter_column(xlpp_autofilter f, uint64_t columnId) {
    return reinterpret_cast<xlpp_filtercol>(&AF(f)->column(static_cast<std::size_t>(columnId)));
}
XLPP_API xlpp_sortstate xlpp_autofilter_sort_state(xlpp_autofilter f) { return reinterpret_cast<xlpp_sortstate>(&AF(f)->sortState()); }

XLPP_API uint64_t xlpp_filtercol_column_id(xlpp_filtercol c) { return FC(c)->columnId(); }
XLPP_API void xlpp_filtercol_add_value(xlpp_filtercol c, const char* v) { FC(c)->addValue(v); }
XLPP_API void xlpp_filtercol_clear_values(xlpp_filtercol c) { FC(c)->clearValues(); }
XLPP_API int xlpp_filtercol_value_count(xlpp_filtercol c) { return static_cast<int>(FC(c)->values().size()); }
XLPP_API void xlpp_filtercol_value_at(xlpp_filtercol c, int index, char* out, int outSize) {
    try { copyStr(FC(c)->values()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_filtercol_set_and_mode(xlpp_filtercol c, int v) { FC(c)->setAndMode(v != 0); }
XLPP_API int xlpp_filtercol_and_mode(xlpp_filtercol c) { return FC(c)->andMode() ? 1 : 0; }
XLPP_API void xlpp_filtercol_set_include_blank(xlpp_filtercol c, int v) { FC(c)->setIncludeBlank(v != 0); }
XLPP_API int xlpp_filtercol_include_blank(xlpp_filtercol c) { return FC(c)->includeBlank() ? 1 : 0; }

XLPP_API void xlpp_sortstate_set_reference(xlpp_sortstate s, const char* v) { SS(s)->setReference(v); }
XLPP_API void xlpp_sortstate_reference(xlpp_sortstate s, char* out, int outSize) { copyStr(SS(s)->reference(), out, outSize); }
XLPP_API void xlpp_sortstate_set_case_sensitive(xlpp_sortstate s, int v) { SS(s)->setCaseSensitive(v != 0); }
XLPP_API int xlpp_sortstate_case_sensitive(xlpp_sortstate s) { return SS(s)->caseSensitive() ? 1 : 0; }
XLPP_API void xlpp_sortstate_add_condition(xlpp_sortstate s, const char* reference, int descending) { SS(s)->addCondition(reference, descending != 0); }
XLPP_API void xlpp_sortstate_clear(xlpp_sortstate s) { SS(s)->clear(); }

// ============================================================
// Page setup / margins / print options / header footer
// ============================================================
#define PS(h) reinterpret_cast<xlpp::PageSetup*>(h)
#define PM(h) reinterpret_cast<xlpp::PageMargins*>(h)
#define PO(h) reinterpret_cast<xlpp::PrintOptions*>(h)
#define HF(h) reinterpret_cast<xlpp::HeaderFooter*>(h)
XLPP_API void xlpp_pagesetup_set_orientation(xlpp_pagesetup p, int v) { PS(p)->setOrientation(static_cast<xlpp::PageOrientation>(v)); }
XLPP_API int xlpp_pagesetup_orientation(xlpp_pagesetup p) { return static_cast<int>(PS(p)->orientation()); }
XLPP_API void xlpp_pagesetup_set_paper_size(xlpp_pagesetup p, int v) { PS(p)->setPaperSize(static_cast<xlpp::PaperSize>(v)); }
XLPP_API int xlpp_pagesetup_paper_size(xlpp_pagesetup p) { return static_cast<int>(PS(p)->paperSize()); }
XLPP_API void xlpp_pagesetup_set_scale(xlpp_pagesetup p, unsigned v) { PS(p)->setScale(v); }
XLPP_API unsigned xlpp_pagesetup_scale(xlpp_pagesetup p) { return PS(p)->scale(); }
XLPP_API void xlpp_pagesetup_set_fit_to_width(xlpp_pagesetup p, unsigned v) { PS(p)->setFitToWidth(v); }
XLPP_API unsigned xlpp_pagesetup_fit_to_width(xlpp_pagesetup p) { return PS(p)->fitToWidth(); }
XLPP_API void xlpp_pagesetup_set_fit_to_height(xlpp_pagesetup p, unsigned v) { PS(p)->setFitToHeight(v); }
XLPP_API unsigned xlpp_pagesetup_fit_to_height(xlpp_pagesetup p) { return PS(p)->fitToHeight(); }
XLPP_API void xlpp_pagesetup_set_fit_to_page(xlpp_pagesetup p, int v) { PS(p)->setFitToPage(v != 0); }
XLPP_API int xlpp_pagesetup_fit_to_page(xlpp_pagesetup p) { return PS(p)->fitToPage() ? 1 : 0; }
XLPP_API void xlpp_pagesetup_set_black_and_white(xlpp_pagesetup p, int v) { PS(p)->setBlackAndWhite(v != 0); }
XLPP_API int xlpp_pagesetup_black_and_white(xlpp_pagesetup p) { return PS(p)->blackAndWhite() ? 1 : 0; }
XLPP_API void xlpp_pagesetup_set_draft(xlpp_pagesetup p, int v) { PS(p)->setDraft(v != 0); }
XLPP_API int xlpp_pagesetup_draft(xlpp_pagesetup p) { return PS(p)->draft() ? 1 : 0; }

XLPP_API void xlpp_pagemargins_set_left(xlpp_pagemargins m, double v) { PM(m)->setLeft(v); }
XLPP_API double xlpp_pagemargins_left(xlpp_pagemargins m) { return PM(m)->left(); }
XLPP_API void xlpp_pagemargins_set_right(xlpp_pagemargins m, double v) { PM(m)->setRight(v); }
XLPP_API double xlpp_pagemargins_right(xlpp_pagemargins m) { return PM(m)->right(); }
XLPP_API void xlpp_pagemargins_set_top(xlpp_pagemargins m, double v) { PM(m)->setTop(v); }
XLPP_API double xlpp_pagemargins_top(xlpp_pagemargins m) { return PM(m)->top(); }
XLPP_API void xlpp_pagemargins_set_bottom(xlpp_pagemargins m, double v) { PM(m)->setBottom(v); }
XLPP_API double xlpp_pagemargins_bottom(xlpp_pagemargins m) { return PM(m)->bottom(); }
XLPP_API void xlpp_pagemargins_set_header(xlpp_pagemargins m, double v) { PM(m)->setHeader(v); }
XLPP_API double xlpp_pagemargins_header(xlpp_pagemargins m) { return PM(m)->header(); }
XLPP_API void xlpp_pagemargins_set_footer(xlpp_pagemargins m, double v) { PM(m)->setFooter(v); }
XLPP_API double xlpp_pagemargins_footer(xlpp_pagemargins m) { return PM(m)->footer(); }

XLPP_API void xlpp_printopts_set_horizontal_centered(xlpp_printopts p, int v) { PO(p)->setHorizontalCentered(v != 0); }
XLPP_API int xlpp_printopts_horizontal_centered(xlpp_printopts p) { return PO(p)->horizontalCentered() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_vertical_centered(xlpp_printopts p, int v) { PO(p)->setVerticalCentered(v != 0); }
XLPP_API int xlpp_printopts_vertical_centered(xlpp_printopts p) { return PO(p)->verticalCentered() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_headings(xlpp_printopts p, int v) { PO(p)->setHeadings(v != 0); }
XLPP_API int xlpp_printopts_headings(xlpp_printopts p) { return PO(p)->headings() ? 1 : 0; }
XLPP_API void xlpp_printopts_set_grid_lines(xlpp_printopts p, int v) { PO(p)->setGridLines(v != 0); }
XLPP_API int xlpp_printopts_grid_lines(xlpp_printopts p) { return PO(p)->gridLines() ? 1 : 0; }

XLPP_API void xlpp_headerfooter_set_odd_header(xlpp_headerfooter h, const char* v) { HF(h)->setOddHeader(v); }
XLPP_API void xlpp_headerfooter_set_odd_footer(xlpp_headerfooter h, const char* v) { HF(h)->setOddFooter(v); }
XLPP_API void xlpp_headerfooter_set_even_header(xlpp_headerfooter h, const char* v) { HF(h)->setEvenHeader(v); }
XLPP_API void xlpp_headerfooter_set_even_footer(xlpp_headerfooter h, const char* v) { HF(h)->setEvenFooter(v); }
XLPP_API void xlpp_headerfooter_set_different_odd_even(xlpp_headerfooter h, int v) { HF(h)->setDifferentOddEven(v != 0); }
XLPP_API int xlpp_headerfooter_different_odd_even(xlpp_headerfooter h) { return HF(h)->differentOddEven() ? 1 : 0; }
XLPP_API void xlpp_headerfooter_set_different_first(xlpp_headerfooter h, int v) { HF(h)->setDifferentFirst(v != 0); }
XLPP_API int xlpp_headerfooter_different_first(xlpp_headerfooter h) { return HF(h)->differentFirst() ? 1 : 0; }

// ============================================================
// Worksheet protection
// ============================================================
#define WSP(h) reinterpret_cast<xlpp::WorksheetProtection*>(h)
XLPP_API void xlpp_wssprotection_set_enabled(xlpp_wssprotection p, int v) { WSP(p)->setEnabled(v != 0); }
XLPP_API int xlpp_wssprotection_enabled(xlpp_wssprotection p) { return WSP(p)->enabled() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_password_hash(xlpp_wssprotection p, const char* v) { WSP(p)->setPasswordHash(v); }
XLPP_API const char* xlpp_wssprotection_password_hash(xlpp_wssprotection p) { return WSP(p)->passwordHash().c_str(); }
XLPP_API void xlpp_wssprotection_set_select_locked(xlpp_wssprotection p, int v) { WSP(p)->setSelectLockedCells(v != 0); }
XLPP_API int xlpp_wssprotection_select_locked(xlpp_wssprotection p) { return WSP(p)->selectLockedCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_select_unlocked(xlpp_wssprotection p, int v) { WSP(p)->setSelectUnlockedCells(v != 0); }
XLPP_API int xlpp_wssprotection_select_unlocked(xlpp_wssprotection p) { return WSP(p)->selectUnlockedCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_cells(xlpp_wssprotection p, int v) { WSP(p)->setFormatCells(v != 0); }
XLPP_API int xlpp_wssprotection_format_cells(xlpp_wssprotection p) { return WSP(p)->formatCells() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_columns(xlpp_wssprotection p, int v) { WSP(p)->setFormatColumns(v != 0); }
XLPP_API int xlpp_wssprotection_format_columns(xlpp_wssprotection p) { return WSP(p)->formatColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_format_rows(xlpp_wssprotection p, int v) { WSP(p)->setFormatRows(v != 0); }
XLPP_API int xlpp_wssprotection_format_rows(xlpp_wssprotection p) { return WSP(p)->formatRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_insert_rows(xlpp_wssprotection p, int v) { WSP(p)->setInsertRows(v != 0); }
XLPP_API int xlpp_wssprotection_insert_rows(xlpp_wssprotection p) { return WSP(p)->insertRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_insert_columns(xlpp_wssprotection p, int v) { WSP(p)->setInsertColumns(v != 0); }
XLPP_API int xlpp_wssprotection_insert_columns(xlpp_wssprotection p) { return WSP(p)->insertColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_delete_rows(xlpp_wssprotection p, int v) { WSP(p)->setDeleteRows(v != 0); }
XLPP_API int xlpp_wssprotection_delete_rows(xlpp_wssprotection p) { return WSP(p)->deleteRows() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_delete_columns(xlpp_wssprotection p, int v) { WSP(p)->setDeleteColumns(v != 0); }
XLPP_API int xlpp_wssprotection_delete_columns(xlpp_wssprotection p) { return WSP(p)->deleteColumns() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_sort(xlpp_wssprotection p, int v) { WSP(p)->setSort(v != 0); }
XLPP_API int xlpp_wssprotection_sort(xlpp_wssprotection p) { return WSP(p)->sort() ? 1 : 0; }
XLPP_API void xlpp_wssprotection_set_auto_filter(xlpp_wssprotection p, int v) { WSP(p)->setAutoFilter(v != 0); }
XLPP_API int xlpp_wssprotection_auto_filter(xlpp_wssprotection p) { return WSP(p)->autoFilter() ? 1 : 0; }

// ============================================================
// SheetView
// ============================================================
#define SV(h) reinterpret_cast<xlpp::SheetView*>(h)
XLPP_API void xlpp_sheetview_set_workbook_view_id(xlpp_sheetview s, int v) { SV(s)->setWorkbookViewId(v); }
XLPP_API int xlpp_sheetview_workbook_view_id(xlpp_sheetview s) { return SV(s)->workbookViewId(); }
XLPP_API void xlpp_sheetview_set_tab_color(xlpp_sheetview s, const char* v) { SV(s)->setTabColor(v); }
XLPP_API void xlpp_sheetview_clear_tab_color(xlpp_sheetview s) { SV(s)->clearTabColor(); }
XLPP_API void xlpp_sheetview_tab_color(xlpp_sheetview s, char* out, int outSize) {
    const auto& c = SV(s)->tabColor();
    if (c) copyStr(*c, out, outSize);
    else if (out && outSize > 0) out[0] = '\0';
}
XLPP_API void xlpp_sheetview_set_zoom_scale(xlpp_sheetview s, int v) { SV(s)->setZoomScale(v); }
XLPP_API int xlpp_sheetview_zoom_scale(xlpp_sheetview s) { return SV(s)->zoomScale(); }
XLPP_API void xlpp_sheetview_set_zoom_scale_normal(xlpp_sheetview s, int v) { SV(s)->setZoomScaleNormal(v); }
XLPP_API int xlpp_sheetview_zoom_scale_normal(xlpp_sheetview s) { return SV(s)->zoomScaleNormal(); }
XLPP_API void xlpp_sheetview_set_show_grid_lines(xlpp_sheetview s, int v) { SV(s)->setShowGridLines(v != 0); }
XLPP_API int xlpp_sheetview_show_grid_lines(xlpp_sheetview s) { return SV(s)->showGridLines() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_tab_selected(xlpp_sheetview s, int v) { SV(s)->setTabSelected(v != 0); }
XLPP_API int xlpp_sheetview_tab_selected(xlpp_sheetview s) { return SV(s)->tabSelected() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_right_to_left(xlpp_sheetview s, int v) { SV(s)->setRightToLeft(v != 0); }
XLPP_API int xlpp_sheetview_right_to_left(xlpp_sheetview s) { return SV(s)->rightToLeft() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_show_outline_symbols(xlpp_sheetview s, int v) { SV(s)->setShowOutlineSymbols(v != 0); }
XLPP_API int xlpp_sheetview_show_outline_symbols(xlpp_sheetview s) { return SV(s)->showOutlineSymbols() ? 1 : 0; }
XLPP_API void xlpp_sheetview_set_pane(xlpp_sheetview s, const char* v) { SV(s)->setPane(v); }
XLPP_API void xlpp_sheetview_pane(xlpp_sheetview s, char* out, int outSize) { copyStr(SV(s)->pane(), out, outSize); }
XLPP_API void xlpp_sheetview_set_top_left_cell(xlpp_sheetview s, const char* v) { SV(s)->setTopLeftCell(v); }
XLPP_API void xlpp_sheetview_top_left_cell(xlpp_sheetview s, char* out, int outSize) { copyStr(SV(s)->topLeftCell(), out, outSize); }
XLPP_API void xlpp_sheetview_set_x_split(xlpp_sheetview s, int v) { SV(s)->setXSplit(v); }
XLPP_API int xlpp_sheetview_x_split(xlpp_sheetview s) { return SV(s)->xSplit(); }
XLPP_API void xlpp_sheetview_set_y_split(xlpp_sheetview s, int v) { SV(s)->setYSplit(v); }
XLPP_API int xlpp_sheetview_y_split(xlpp_sheetview s) { return SV(s)->ySplit(); }

// ============================================================
// Image
// ============================================================
#define IMG(h) reinterpret_cast<xlpp::Image*>(h)
XLPP_API void xlpp_image_set_anchor(xlpp_image img, const char* v) { IMG(img)->setAnchor(v); }
XLPP_API void xlpp_image_anchor(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->anchor(), out, outSize); }
XLPP_API void xlpp_image_extension(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->extension(), out, outSize); }
XLPP_API void xlpp_image_set_width(xlpp_image img, double v) { IMG(img)->setWidthPixels(v); }
XLPP_API double xlpp_image_width(xlpp_image img) { return IMG(img)->widthPixels(); }
XLPP_API void xlpp_image_set_height(xlpp_image img, double v) { IMG(img)->setHeightPixels(v); }
XLPP_API double xlpp_image_height(xlpp_image img) { return IMG(img)->heightPixels(); }
XLPP_API void xlpp_image_set_name(xlpp_image img, const char* v) { IMG(img)->setName(v); }
XLPP_API void xlpp_image_name(xlpp_image img, char* out, int outSize) { copyStr(IMG(img)->name(), out, outSize); }

// ============================================================
// Table
// ============================================================
#define TBL(h) reinterpret_cast<xlpp::Table*>(h)
#define TC(h) reinterpret_cast<xlpp::TableColumn*>(h)
#define TSI(h) reinterpret_cast<xlpp::TableStyleInfo*>(h)
XLPP_API void xlpp_table_set_display_name(xlpp_table t, const char* v) { try { TBL(t)->setDisplayName(v); } catch (...) {} }
XLPP_API void xlpp_table_name(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->name(), out, outSize); }
XLPP_API void xlpp_table_display_name(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->displayName(), out, outSize); }
XLPP_API void xlpp_table_set_reference(xlpp_table t, const char* v) { try { TBL(t)->setReference(v); } catch (...) {} }
XLPP_API void xlpp_table_reference(xlpp_table t, char* out, int outSize) { copyStr(TBL(t)->reference(), out, outSize); }
XLPP_API void xlpp_table_set_show_header_row(xlpp_table t, int v) { TBL(t)->setShowHeaderRow(v != 0); }
XLPP_API int xlpp_table_show_header_row(xlpp_table t) { return TBL(t)->showHeaderRow() ? 1 : 0; }
XLPP_API void xlpp_table_set_show_totals_row(xlpp_table t, int v) { TBL(t)->setShowTotalsRow(v != 0); }
XLPP_API int xlpp_table_show_totals_row(xlpp_table t) { return TBL(t)->showTotalsRow() ? 1 : 0; }
XLPP_API int xlpp_table_column_count(xlpp_table t) { return static_cast<int>(TBL(t)->columns().size()); }
XLPP_API xlpp_tablecolumn xlpp_table_column_at(xlpp_table t, int index) {
    try { return reinterpret_cast<xlpp_tablecolumn>(&TBL(t)->columns()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API xlpp_tablecolumn xlpp_table_add_column(xlpp_table t, const char* name) {
    return reinterpret_cast<xlpp_tablecolumn>(&TBL(t)->addColumn(name));
}
XLPP_API xlpp_tablestyle xlpp_table_style_info(xlpp_table t) { return reinterpret_cast<xlpp_tablestyle>(&TBL(t)->styleInfo()); }

XLPP_API void xlpp_tablecolumn_set_name(xlpp_tablecolumn c, const char* v) { try { TC(c)->setName(v); } catch (...) {} }
XLPP_API void xlpp_tablecolumn_name(xlpp_tablecolumn c, char* out, int outSize) { copyStr(TC(c)->name(), out, outSize); }
XLPP_API uint64_t xlpp_tablecolumn_id(xlpp_tablecolumn c) { return TC(c)->id(); }

XLPP_API void xlpp_tablestyle_set_name(xlpp_tablestyle s, const char* v) { TSI(s)->setName(v); }
XLPP_API void xlpp_tablestyle_name(xlpp_tablestyle s, char* out, int outSize) { copyStr(TSI(s)->name(), out, outSize); }
XLPP_API void xlpp_tablestyle_set_show_first(xlpp_tablestyle s, int v) { TSI(s)->setShowFirstColumn(v != 0); }
XLPP_API int xlpp_tablestyle_show_first(xlpp_tablestyle s) { return TSI(s)->showFirstColumn() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_last(xlpp_tablestyle s, int v) { TSI(s)->setShowLastColumn(v != 0); }
XLPP_API int xlpp_tablestyle_show_last(xlpp_tablestyle s) { return TSI(s)->showLastColumn() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_row_stripes(xlpp_tablestyle s, int v) { TSI(s)->setShowRowStripes(v != 0); }
XLPP_API int xlpp_tablestyle_show_row_stripes(xlpp_tablestyle s) { return TSI(s)->showRowStripes() ? 1 : 0; }
XLPP_API void xlpp_tablestyle_set_show_column_stripes(xlpp_tablestyle s, int v) { TSI(s)->setShowColumnStripes(v != 0); }
XLPP_API int xlpp_tablestyle_show_column_stripes(xlpp_tablestyle s) { return TSI(s)->showColumnStripes() ? 1 : 0; }

// ============================================================
// Chart
// ============================================================
#define CH(h) reinterpret_cast<xlpp::Chart*>(h)
#define CS(h) reinterpret_cast<xlpp::ChartSeries*>(h)
XLPP_API void xlpp_chart_set_grouping(xlpp_chart c, int v) { CH(c)->setGrouping(static_cast<xlpp::Chart::Grouping>(v)); }
XLPP_API int xlpp_chart_grouping(xlpp_chart c) { return static_cast<int>(CH(c)->grouping()); }
XLPP_API void xlpp_chart_set_title(xlpp_chart c, const char* v) { CH(c)->setTitle(v); }
XLPP_API void xlpp_chart_title(xlpp_chart c, char* out, int outSize) { copyStr(CH(c)->title(), out, outSize); }
XLPP_API void xlpp_chart_set_x_axis_title(xlpp_chart c, const char* v) { CH(c)->setXAxisTitle(v); }
XLPP_API void xlpp_chart_set_y_axis_title(xlpp_chart c, const char* v) { CH(c)->setYAxisTitle(v); }
XLPP_API void xlpp_chart_set_style(xlpp_chart c, const char* v) { CH(c)->setStyle(v); }
XLPP_API void xlpp_chart_set_width(xlpp_chart c, int v) { CH(c)->setWidth(v); }
XLPP_API int xlpp_chart_width(xlpp_chart c) { return CH(c)->width(); }
XLPP_API void xlpp_chart_set_height(xlpp_chart c, int v) { CH(c)->setHeight(v); }
XLPP_API int xlpp_chart_height(xlpp_chart c) { return CH(c)->height(); }
XLPP_API void xlpp_chart_set_show_legend(xlpp_chart c, int v) { CH(c)->setShowLegend(v != 0); }
XLPP_API int xlpp_chart_show_legend(xlpp_chart c) { return CH(c)->showLegend() ? 1 : 0; }
XLPP_API void xlpp_chart_set_legend_position(xlpp_chart c, const char* v) { CH(c)->setLegendPosition(v); }
XLPP_API int xlpp_chart_series_count(xlpp_chart c) { return static_cast<int>(CH(c)->series().size()); }
XLPP_API xlpp_chartseries xlpp_chart_series_at(xlpp_chart c, int index) {
    try { return reinterpret_cast<xlpp_chartseries>(&CH(c)->series()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API xlpp_chartseries xlpp_chart_add_series(xlpp_chart c, const char* title) {
    return reinterpret_cast<xlpp_chartseries>(&CH(c)->addSeries(xlpp::ChartSeries(title ? title : "")));
}
XLPP_API int xlpp_chart_type(xlpp_chart c) { return static_cast<int>(CH(c)->type()); }
XLPP_API int xlpp_chart_is_modern(xlpp_chart c) { return CH(c)->modern() ? 1 : 0; }
XLPP_API int xlpp_chart_plot_count(xlpp_chart c) { return static_cast<int>(CH(c)->plots().size()); }
XLPP_API int xlpp_chart_add_plot(xlpp_chart c, int type, uint64_t first_series, uint64_t series_count, int secondary_axes) {
    try { CH(c)->addPlot(static_cast<xlpp::Chart::Type>(type), static_cast<std::size_t>(first_series), static_cast<std::size_t>(series_count), secondary_axes != 0); return static_cast<int>(CH(c)->plots().size() - 1); } catch (...) { return -1; }
}
XLPP_API void xlpp_chart_plot_set_grouping(xlpp_chart c, int index, int grouping) { try { CH(c)->plots().at(static_cast<std::size_t>(index)).grouping = static_cast<xlpp::Chart::Grouping>(grouping); } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_bar_direction(xlpp_chart c, int index, int direction) { try { CH(c)->plots().at(static_cast<std::size_t>(index)).barDirection = static_cast<xlpp::Chart::BarDirection>(direction); } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_scatter_style(xlpp_chart c, int index, int style) { try { CH(c)->plots().at(static_cast<std::size_t>(index)).scatterStyle = static_cast<xlpp::Chart::ScatterStyle>(style); } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_bubble_options(xlpp_chart c, int index, int scale, int show_negative, int size_represents, int bubble3d) { try { auto& p = CH(c)->plots().at(static_cast<std::size_t>(index)); p.hasBubbleScale = true; p.bubbleScale = scale; p.showNegativeBubbles = show_negative != 0; p.bubbleSizeRepresents = static_cast<xlpp::Chart::BubbleSizeRepresents>(size_represents); p.bubble3D = bubble3d != 0; } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_histogram_bins(xlpp_chart c, int index, double bin_width, int bin_count, int automatic_bins) { try { auto& p = CH(c)->plots().at(static_cast<std::size_t>(index)); p.histogramBinWidth = bin_width; p.histogramBinCount = bin_count; p.histogramAutomaticBins = automatic_bins != 0; } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_histogram_bounds(xlpp_chart c, int index, int has_underflow, double underflow, int has_overflow, double overflow) { try { auto& p = CH(c)->plots().at(static_cast<std::size_t>(index)); p.histogramHasUnderflow = has_underflow != 0; p.histogramUnderflow = underflow; p.histogramHasOverflow = has_overflow != 0; p.histogramOverflow = overflow; } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_box_whisker_options(xlpp_chart c, int index, int inner_points, int outliers, int mean_line, int mean_marker, int quartile_inclusive) { try { auto& p = CH(c)->plots().at(static_cast<std::size_t>(index)); p.boxWhiskerShowInnerPoints = inner_points != 0; p.boxWhiskerShowOutlierPoints = outliers != 0; p.boxWhiskerShowMeanLine = mean_line != 0; p.boxWhiskerShowMeanMarker = mean_marker != 0; p.boxWhiskerQuartileInclusive = quartile_inclusive != 0; } catch (...) {} }
XLPP_API void xlpp_chart_plot_set_waterfall_connector_lines(xlpp_chart c, int index, int enabled) { try { CH(c)->plots().at(static_cast<std::size_t>(index)).waterfallShowConnectorLines = enabled != 0; } catch (...) {} }

XLPP_API void xlpp_chartseries_set_title(xlpp_chartseries s, const char* v) { CS(s)->setTitle(v); }
XLPP_API void xlpp_chartseries_set_values_reference(xlpp_chartseries s, const char* v) { CS(s)->setValuesReference(v); }
XLPP_API void xlpp_chartseries_set_categories_reference(xlpp_chartseries s, const char* v) { CS(s)->setCategoriesReference(v); }
XLPP_API void xlpp_chartseries_set_bubble_size_reference(xlpp_chartseries s, const char* v) { CS(s)->setBubbleSizeReference(v ? v : ""); }
XLPP_API void xlpp_chartseries_set_smooth(xlpp_chartseries s, int v) { CS(s)->setSmooth(v != 0); }
XLPP_API void xlpp_chartseries_clear_smooth(xlpp_chartseries s) { CS(s)->clearSmooth(); }

// ============================================================
// Pivot table
// ============================================================
#define PT(h) reinterpret_cast<xlpp::PivotTable*>(h)
#define PC(h) reinterpret_cast<xlpp::PivotCache*>(h)
XLPP_API void xlpp_pivottable_set_name(xlpp_pivottable p, const char* v) { PT(p)->setName(v); }
XLPP_API void xlpp_pivottable_name(xlpp_pivottable p, char* out, int outSize) { copyStr(PT(p)->name(), out, outSize); }
XLPP_API void xlpp_pivottable_set_location(xlpp_pivottable p, const char* v) { PT(p)->setLocation(v); }
XLPP_API void xlpp_pivottable_location(xlpp_pivottable p, char* out, int outSize) { copyStr(PT(p)->location(), out, outSize); }
XLPP_API xlpp_pivotcache xlpp_pivottable_cache(xlpp_pivottable p) { return reinterpret_cast<xlpp_pivotcache>(&PT(p)->cache()); }
XLPP_API void xlpp_pivottable_add_row_field(xlpp_pivottable p, const char* name) { PT(p)->addRowField(name); }
XLPP_API void xlpp_pivottable_add_column_field(xlpp_pivottable p, const char* name) { PT(p)->addColumnField(name); }
XLPP_API void xlpp_pivottable_add_page_field(xlpp_pivottable p, const char* name) { PT(p)->addPageField(name); }
XLPP_API void xlpp_pivottable_add_data_field(xlpp_pivottable p) { PT(p)->addDataField(); }
XLPP_API void xlpp_pivottable_add_data_field_named(xlpp_pivottable p, const char* name, const char* subtotal) { try { PT(p)->addDataField(name ? name : "", subtotal ? subtotal : "sum"); } catch (...) {} }
XLPP_API void xlpp_pivottable_set_layout(xlpp_pivottable p, int layout) { PT(p)->setLayout(static_cast<xlpp::PivotLayout>(layout)); }
XLPP_API int xlpp_pivottable_layout(xlpp_pivottable p) { return static_cast<int>(PT(p)->layout()); }
XLPP_API void xlpp_pivottable_set_flags(xlpp_pivottable p, int show_empty_row, int show_empty_column, int show_drill, int enable_drill, int multiple_field_filters, int show_values_row, int subtotal_hidden_items) { PT(p)->setShowEmptyRow(show_empty_row != 0); PT(p)->setShowEmptyColumn(show_empty_column != 0); PT(p)->setShowDrill(show_drill != 0); PT(p)->setEnableDrill(enable_drill != 0); PT(p)->setMultipleFieldFilters(multiple_field_filters != 0); PT(p)->setShowValuesRow(show_values_row != 0); PT(p)->setSubtotalHiddenItems(subtotal_hidden_items != 0); }
XLPP_API void xlpp_pivottable_set_page_layout(xlpp_pivottable p, int page_wrap, int over_then_down) { try { PT(p)->setPageWrap(page_wrap); PT(p)->setPageOverThenDown(over_then_down != 0); } catch (...) {} }
XLPP_API void xlpp_pivottable_set_style(xlpp_pivottable p, const char* style_name, const char* data_caption) { PT(p)->setStyleName(style_name ? style_name : ""); PT(p)->setDataCaption(data_caption ? data_caption : "Values"); }
XLPP_API void xlpp_pivottable_set_display_options(xlpp_pivottable p, int row_grand_totals, int column_grand_totals, int preserve_formatting, int use_auto_formatting, int show_row_headers, int show_column_headers, int show_row_stripes, int show_column_stripes, int show_last_column) { PT(p)->setRowGrandTotals(row_grand_totals != 0); PT(p)->setColumnGrandTotals(column_grand_totals != 0); PT(p)->setPreserveFormatting(preserve_formatting != 0); PT(p)->setUseAutoFormatting(use_auto_formatting != 0); PT(p)->setShowRowHeaders(show_row_headers != 0); PT(p)->setShowColumnHeaders(show_column_headers != 0); PT(p)->setShowRowStripes(show_row_stripes != 0); PT(p)->setShowColumnStripes(show_column_stripes != 0); PT(p)->setShowLastColumn(show_last_column != 0); }
XLPP_API int xlpp_pivottable_data_field_count(xlpp_pivottable p) { return static_cast<int>(PT(p)->dataFields().size()); }
XLPP_API void xlpp_pivottable_configure_data_field(xlpp_pivottable p, int index, const char* subtotal, const char* caption, uint32_t number_format_id, const char* show_data_as, int base_field, int base_item) { try { auto& f = PT(p)->dataFields().at(static_cast<std::size_t>(index)); if (subtotal && *subtotal) f.setSubtotal(subtotal); if (caption) f.setCaption(caption); f.setNumberFormatId(number_format_id); if (show_data_as && *show_data_as) f.setShowDataAs(show_data_as); f.setBaseField(base_field); f.setBaseItem(base_item); } catch (...) {} }
XLPP_API int xlpp_pivottable_row_field_count(xlpp_pivottable p) { return static_cast<int>(PT(p)->rowFields().size()); }
XLPP_API int xlpp_pivottable_column_field_count(xlpp_pivottable p) { return static_cast<int>(PT(p)->columnFields().size()); }
XLPP_API int xlpp_pivottable_page_field_count(xlpp_pivottable p) { return static_cast<int>(PT(p)->pageFields().size()); }
XLPP_API xlpp_pivotfield xlpp_pivottable_row_field_at(xlpp_pivottable p, int index) { try { return reinterpret_cast<xlpp_pivotfield>(&PT(p)->rowFields().at(static_cast<std::size_t>(index))); } catch (...) { return nullptr; } }
XLPP_API xlpp_pivotfield xlpp_pivottable_column_field_at(xlpp_pivottable p, int index) { try { return reinterpret_cast<xlpp_pivotfield>(&PT(p)->columnFields().at(static_cast<std::size_t>(index))); } catch (...) { return nullptr; } }
XLPP_API xlpp_pivotfield xlpp_pivottable_page_field_at(xlpp_pivottable p, int index) { try { return reinterpret_cast<xlpp_pivotfield>(&PT(p)->pageFields().at(static_cast<std::size_t>(index))); } catch (...) { return nullptr; } }
XLPP_API void xlpp_pivottable_add_filter(xlpp_pivottable p, const char* type, int field_index, int measure_field_index, const char* value1, const char* value2, double top10_value, int top10_percent, int top10_top) { xlpp::PivotFilter f; f.type = type ? type : "unknown"; f.fieldIndex = field_index; f.measureFieldIndex = measure_field_index; f.value1 = value1 ? value1 : ""; f.value2 = value2 ? value2 : ""; f.top10Value = top10_value; f.top10Percent = top10_percent != 0; f.top10Top = top10_top != 0; PT(p)->addFilter(std::move(f)); }

XLPP_API void xlpp_pivotcache_set_cache_id(xlpp_pivotcache c, int v) { PC(c)->setCacheId(v); }
XLPP_API int xlpp_pivotcache_cache_id(xlpp_pivotcache c) { return PC(c)->cacheId(); }
XLPP_API void xlpp_pivotcache_set_source_data(xlpp_pivotcache c, const char* v) { PC(c)->setSourceData(v); }
XLPP_API void xlpp_pivotcache_source_data(xlpp_pivotcache c, char* out, int outSize) { copyStr(PC(c)->sourceData(), out, outSize); }
XLPP_API void xlpp_pivotcache_set_refresh_on_load(xlpp_pivotcache c, int v) { PC(c)->setRefreshOnLoad(v != 0); }
XLPP_API void xlpp_pivotcache_set_save_data(xlpp_pivotcache c, int v) { PC(c)->setSaveData(v != 0); }
XLPP_API void xlpp_pivotcache_set_enable_refresh(xlpp_pivotcache c, int v) { PC(c)->setEnableRefresh(v != 0); }
XLPP_API void xlpp_pivotcache_set_missing_items_limit(xlpp_pivotcache c, int v) { try { PC(c)->setMissingItemsLimit(v); } catch (...) {} }
XLPP_API void xlpp_pivotcache_set_advanced_flags(xlpp_pivotcache c, int background_query, int optimize_memory, int upgrade_on_refresh, int support_subquery, int support_advanced_drill) { PC(c)->setBackgroundQuery(background_query != 0); PC(c)->setOptimizeMemory(optimize_memory != 0); PC(c)->setUpgradeOnRefresh(upgrade_on_refresh != 0); PC(c)->setSupportSubquery(support_subquery != 0); PC(c)->setSupportAdvancedDrill(support_advanced_drill != 0); }
XLPP_API void xlpp_pivotcache_set_refreshed_by(xlpp_pivotcache c, const char* v) { PC(c)->setRefreshedBy(v ? v : ""); }

#define PF(h) reinterpret_cast<xlpp::PivotField*>(h)
XLPP_API void xlpp_pivotfield_set_repeat_item_labels(xlpp_pivotfield f, int v) { PF(f)->setRepeatItemLabels(v != 0); }
XLPP_API void xlpp_pivotfield_set_compact(xlpp_pivotfield f, int v) { PF(f)->setCompact(v != 0); }
XLPP_API void xlpp_pivotfield_set_outline(xlpp_pivotfield f, int v) { PF(f)->setOutline(v != 0); }
XLPP_API void xlpp_pivotfield_set_show_dropdowns(xlpp_pivotfield f, int v) { PF(f)->setShowDropDowns(v != 0); }
XLPP_API void xlpp_pivotfield_set_behavior(xlpp_pivotfield f, int show_all, int sort_type, int subtotal_top, int insert_blank_row, int include_new_items_in_filter, int multiple_item_selection_allowed, int selected_item_index, int insert_page_break, int default_subtotal) { PF(f)->setShowAll(show_all != 0); PF(f)->setSortType(sort_type); PF(f)->setSubtotalTop(subtotal_top != 0); PF(f)->setInsertBlankRow(insert_blank_row != 0); PF(f)->setIncludeNewItemsInFilter(include_new_items_in_filter != 0); PF(f)->setMultipleItemSelectionAllowed(multiple_item_selection_allowed != 0); PF(f)->setSelectedItemIndex(selected_item_index); PF(f)->setInsertPageBreak(insert_page_break != 0); PF(f)->setDefaultSubtotal(default_subtotal != 0); }
XLPP_API void xlpp_pivotfield_add_subtotal(xlpp_pivotfield f, const char* subtotal) { try { PF(f)->addSubtotal(subtotal ? subtotal : "sum"); } catch (...) {} }
XLPP_API void xlpp_pivotfield_set_item_hidden(xlpp_pivotfield f, int item_index, int hidden) { try { PF(f)->setItemHidden(item_index, hidden != 0); } catch (...) {} }
XLPP_API void xlpp_pivotfield_set_numeric_grouping(xlpp_pivotfield f, int auto_start, int auto_end, double start, double end, double interval) { xlpp::PivotGrouping g; g.kind = xlpp::PivotGrouping::Kind::Numeric; g.autoStart = auto_start != 0; g.autoEnd = auto_end != 0; g.start = start; g.end = end; g.interval = interval; PF(f)->setGrouping(std::move(g)); }
XLPP_API void xlpp_pivotfield_set_date_grouping(xlpp_pivotfield f, int date_part, int auto_start, int auto_end, const char* start_date, const char* end_date) { xlpp::PivotGrouping g; g.kind = xlpp::PivotGrouping::Kind::Date; g.datePart = static_cast<xlpp::PivotGrouping::DatePart>(date_part); g.autoStart = auto_start != 0; g.autoEnd = auto_end != 0; g.startDate = start_date ? start_date : ""; g.endDate = end_date ? end_date : ""; PF(f)->setGrouping(std::move(g)); }
XLPP_API void xlpp_pivotfield_clear_grouping(xlpp_pivotfield f) { PF(f)->setGrouping(xlpp::PivotGrouping{}); }

// ============================================================
// Conditional formatting
// ============================================================
#define CF(h) reinterpret_cast<xlpp::ConditionalFormattingCollection*>(h)
#define CFE(h) reinterpret_cast<xlpp::ConditionalFormattingEntry*>(h)
#define CFR(h) reinterpret_cast<xlpp::ConditionalRule*>(h)
XLPP_API int xlpp_cfcollection_entry_count(xlpp_cfcollection c) { return static_cast<int>(CF(c)->entries().size()); }
XLPP_API xlpp_cfentry xlpp_cfcollection_add_entry(xlpp_cfcollection c, const char* reference) {
    try { return reinterpret_cast<xlpp_cfentry>(&CF(c)->add(reference)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_cfentry xlpp_cfcollection_entry_at(xlpp_cfcollection c, int index) {
    try { return reinterpret_cast<xlpp_cfentry>(&CF(c)->entries()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API void xlpp_cfcollection_clear(xlpp_cfcollection c) { CF(c)->clear(); }
XLPP_API int xlpp_cfcollection_empty(xlpp_cfcollection c) { return CF(c)->empty() ? 1 : 0; }

XLPP_API void xlpp_cfentry_reference(xlpp_cfentry e, char* out, int outSize) { copyStr(CFE(e)->reference(), out, outSize); }
XLPP_API void xlpp_cfentry_set_reference(xlpp_cfentry e, const char* v) { try { CFE(e)->setReference(v); } catch (...) {} }
XLPP_API int xlpp_cfentry_rule_count(xlpp_cfentry e) { return static_cast<int>(CFE(e)->rules().size()); }
XLPP_API xlpp_cfrule xlpp_cfentry_add_rule(xlpp_cfentry e, int type) {
    xlpp::ConditionalRule rule;
    switch (type) {
    case 0: rule = xlpp::ConditionalRule::formula(""); break;
    case 1: rule = xlpp::ConditionalRule::cellIs(xlpp::ConditionalOperator::Equal, ""); break;
    case 2: rule = xlpp::ConditionalRule::dataBar(); break;
    case 3: rule = xlpp::ConditionalRule::colorScale(); break;
    case 4: rule = xlpp::ConditionalRule::iconSet(); break;
    default: rule = xlpp::ConditionalRule::formula(""); break;
    }
    return reinterpret_cast<xlpp_cfrule>(&CFE(e)->addRule(std::move(rule)));
}
XLPP_API xlpp_cfrule xlpp_cfentry_rule_at(xlpp_cfentry e, int index) {
    try { return reinterpret_cast<xlpp_cfrule>(&CFE(e)->rules()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}

XLPP_API void xlpp_cfrule_add_formula(xlpp_cfrule r, const char* f) { CFR(r)->addFormula(f); }
XLPP_API int xlpp_cfrule_formula_count(xlpp_cfrule r) { return static_cast<int>(CFR(r)->formulas().size()); }
XLPP_API void xlpp_cfrule_formula_at(xlpp_cfrule r, int index, char* out, int outSize) {
    try { copyStr(CFR(r)->formulas()[static_cast<std::size_t>(index)], out, outSize); } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}
XLPP_API void xlpp_cfrule_set_operator(xlpp_cfrule r, int op) { CFR(r)->setOperator(static_cast<xlpp::ConditionalOperator>(op)); }
XLPP_API int xlpp_cfrule_operator(xlpp_cfrule r) { return static_cast<int>(CFR(r)->op()); }
XLPP_API int xlpp_cfrule_type(xlpp_cfrule r) { return static_cast<int>(CFR(r)->type()); }
XLPP_API void xlpp_cfrule_set_priority(xlpp_cfrule r, uint64_t v) { CFR(r)->setPriority(static_cast<std::size_t>(v)); }
XLPP_API uint64_t xlpp_cfrule_priority(xlpp_cfrule r) { return CFR(r)->priority(); }
XLPP_API void xlpp_cfrule_set_stop_if_true(xlpp_cfrule r, int v) { CFR(r)->setStopIfTrue(v != 0); }
XLPP_API int xlpp_cfrule_stop_if_true(xlpp_cfrule r) { return CFR(r)->stopIfTrue() ? 1 : 0; }
XLPP_API void xlpp_cfrule_set_differential_style(xlpp_cfrule r, xlpp_style s) { CFR(r)->setDifferentialStyle(*STY(s)); }

// ============================================================
// Data validation
// ============================================================
#define DV(h) reinterpret_cast<xlpp::DataValidationCollection*>(h)
#define DVX(h) reinterpret_cast<xlpp::DataValidation*>(h)
XLPP_API int xlpp_dvcollection_count(xlpp_dvcollection c) { return static_cast<int>(DV(c)->items().size()); }
XLPP_API xlpp_datavalidation xlpp_dvcollection_add(xlpp_dvcollection c, int type, const char* reference) {
    try { return reinterpret_cast<xlpp_datavalidation>(&DV(c)->add(static_cast<xlpp::DataValidationType>(type), reference)); } catch (...) { return nullptr; }
}
XLPP_API xlpp_datavalidation xlpp_dvcollection_at(xlpp_dvcollection c, int index) {
    try { return reinterpret_cast<xlpp_datavalidation>(&DV(c)->items()[static_cast<std::size_t>(index)]); } catch (...) { return nullptr; }
}
XLPP_API void xlpp_dvcollection_clear(xlpp_dvcollection c) { DV(c)->clear(); }

XLPP_API int xlpp_datavalidation_type(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->type()); }
XLPP_API void xlpp_datavalidation_set_type(xlpp_datavalidation d, int v) { DVX(d)->setType(static_cast<xlpp::DataValidationType>(v)); }
XLPP_API int xlpp_datavalidation_operator(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->op()); }
XLPP_API void xlpp_datavalidation_set_operator(xlpp_datavalidation d, int v) { DVX(d)->setOperator(static_cast<xlpp::DataValidationOperator>(v)); }
XLPP_API int xlpp_datavalidation_error_style(xlpp_datavalidation d) { return static_cast<int>(DVX(d)->errorStyle()); }
XLPP_API void xlpp_datavalidation_set_error_style(xlpp_datavalidation d, int v) { DVX(d)->setErrorStyle(static_cast<xlpp::DataValidationErrorStyle>(v)); }
XLPP_API void xlpp_datavalidation_set_formula1(xlpp_datavalidation d, const char* v) { DVX(d)->setFormula1(v); }
XLPP_API void xlpp_datavalidation_formula1(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->formula1(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_formula2(xlpp_datavalidation d, const char* v) { DVX(d)->setFormula2(v); }
XLPP_API void xlpp_datavalidation_formula2(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->formula2(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_reference(xlpp_datavalidation d, const char* v) { try { DVX(d)->setReference(v); } catch (...) {} }
XLPP_API void xlpp_datavalidation_reference(xlpp_datavalidation d, char* out, int outSize) { copyStr(DVX(d)->reference(), out, outSize); }
XLPP_API void xlpp_datavalidation_set_allow_blank(xlpp_datavalidation d, int v) { DVX(d)->setAllowBlank(v != 0); }
XLPP_API int xlpp_datavalidation_allow_blank(xlpp_datavalidation d) { return DVX(d)->allowBlank() ? 1 : 0; }
XLPP_API void xlpp_datavalidation_set_show_drop_down(xlpp_datavalidation d, int v) { DVX(d)->setShowDropDown(v != 0); }
XLPP_API int xlpp_datavalidation_show_drop_down(xlpp_datavalidation d) { return DVX(d)->showDropDown() ? 1 : 0; }
XLPP_API void xlpp_datavalidation_set_show_input_message(xlpp_datavalidation d, int v) { DVX(d)->setShowInputMessage(v != 0); }
XLPP_API void xlpp_datavalidation_set_show_error_message(xlpp_datavalidation d, int v) { DVX(d)->setShowErrorMessage(v != 0); }
XLPP_API void xlpp_datavalidation_set_prompt_title(xlpp_datavalidation d, const char* v) { DVX(d)->setPromptTitle(v); }
XLPP_API void xlpp_datavalidation_set_prompt(xlpp_datavalidation d, const char* v) { DVX(d)->setPrompt(v); }
XLPP_API void xlpp_datavalidation_set_error_title(xlpp_datavalidation d, const char* v) { DVX(d)->setErrorTitle(v); }
XLPP_API void xlpp_datavalidation_set_error(xlpp_datavalidation d, const char* v) { DVX(d)->setError(v); }

// ============================================================
// Font
// ============================================================
XLPP_API void xlpp_font_set_name(xlpp_font f, const char* v)       { FONT(f)->setName(v); }
XLPP_API void xlpp_font_set_size(xlpp_font f, double v)            { FONT(f)->setSize(v); }
XLPP_API void xlpp_font_set_bold(xlpp_font f, int v)               { FONT(f)->setBold(v != 0); }
XLPP_API void xlpp_font_set_italic(xlpp_font f, int v)             { FONT(f)->setItalic(v != 0); }
XLPP_API void xlpp_font_set_underline(xlpp_font f, int v)          { FONT(f)->setUnderline(v != 0); }
XLPP_API void xlpp_font_set_strike(xlpp_font f, int v)             { FONT(f)->setStrike(v != 0); }
XLPP_API void xlpp_font_set_color(xlpp_font f, const char* argb)   { FONT(f)->color().setArgb(argb); }
XLPP_API const char* xlpp_font_get_name(xlpp_font f)               { return FONT(f)->name().c_str(); }
XLPP_API double xlpp_font_get_size(xlpp_font f)                    { return FONT(f)->size(); }
XLPP_API int xlpp_font_get_bold(xlpp_font f)                       { return FONT(f)->bold() ? 1 : 0; }
XLPP_API int xlpp_font_get_italic(xlpp_font f)                     { return FONT(f)->italic() ? 1 : 0; }
XLPP_API int xlpp_font_get_underline(xlpp_font f)                  { return FONT(f)->underline() ? 1 : 0; }
XLPP_API int xlpp_font_get_strike(xlpp_font f)                     { return FONT(f)->strike() ? 1 : 0; }
XLPP_API void xlpp_font_get_color(xlpp_font f, char* out, int outSize) { copyStr(FONT(f)->color().argb(), out, outSize); }

// ============================================================
// Fill
// ============================================================
XLPP_API void xlpp_fill_set_pattern(xlpp_fill f, const char* v)     { FILL(f)->setPatternType(v); }
XLPP_API void xlpp_fill_set_fg_color(xlpp_fill f, const char* argb) { FILL(f)->foregroundColor().setArgb(argb); }
XLPP_API void xlpp_fill_set_bg_color(xlpp_fill f, const char* argb) { FILL(f)->backgroundColor().setArgb(argb); }
XLPP_API void xlpp_fill_get_pattern(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->patternType(), out, outSize); }
XLPP_API void xlpp_fill_get_fg_color(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->foregroundColor().argb(), out, outSize); }
XLPP_API void xlpp_fill_get_bg_color(xlpp_fill f, char* out, int outSize) { copyStr(FILL(f)->backgroundColor().argb(), out, outSize); }

// ============================================================
// Border
// ============================================================
XLPP_API xlpp_borderside xlpp_border_left(xlpp_border b)    { return reinterpret_cast<xlpp_borderside>(&BDR(b)->left()); }
XLPP_API xlpp_borderside xlpp_border_right(xlpp_border b)   { return reinterpret_cast<xlpp_borderside>(&BDR(b)->right()); }
XLPP_API xlpp_borderside xlpp_border_top(xlpp_border b)     { return reinterpret_cast<xlpp_borderside>(&BDR(b)->top()); }
XLPP_API xlpp_borderside xlpp_border_bottom(xlpp_border b)  { return reinterpret_cast<xlpp_borderside>(&BDR(b)->bottom()); }
XLPP_API xlpp_borderside xlpp_border_diagonal(xlpp_border b){ return reinterpret_cast<xlpp_borderside>(&BDR(b)->diagonal()); }
XLPP_API void xlpp_borderside_set_style(xlpp_borderside s, const char* v) { BS(s)->setStyle(v); }
XLPP_API void xlpp_borderside_set_color(xlpp_borderside s, const char* argb) { BS(s)->color().setArgb(argb); }
XLPP_API void xlpp_borderside_get_style(xlpp_borderside s, char* out, int outSize) { copyStr(BS(s)->style(), out, outSize); }
XLPP_API void xlpp_borderside_get_color(xlpp_borderside s, char* out, int outSize) { copyStr(BS(s)->color().argb(), out, outSize); }

// ============================================================
// Alignment
// ============================================================
XLPP_API void xlpp_alignment_set_horizontal(xlpp_alignment a, const char* v) { ALN(a)->setHorizontal(v); }
XLPP_API void xlpp_alignment_set_vertical(xlpp_alignment a, const char* v)   { ALN(a)->setVertical(v); }
XLPP_API void xlpp_alignment_set_wrap_text(xlpp_alignment a, int v)           { ALN(a)->setWrapText(v != 0); }
XLPP_API void xlpp_alignment_set_shrink_to_fit(xlpp_alignment a, int v)       { ALN(a)->setShrinkToFit(v != 0); }
XLPP_API void xlpp_alignment_set_text_rotation(xlpp_alignment a, int v)       { ALN(a)->setTextRotation(v); }
XLPP_API void xlpp_alignment_set_indent(xlpp_alignment a, int v)              { ALN(a)->setIndent(v); }
XLPP_API void xlpp_alignment_get_horizontal(xlpp_alignment a, char* out, int outSize) { copyStr(ALN(a)->horizontal(), out, outSize); }
XLPP_API void xlpp_alignment_get_vertical(xlpp_alignment a, char* out, int outSize) { copyStr(ALN(a)->vertical(), out, outSize); }
XLPP_API int xlpp_alignment_get_wrap_text(xlpp_alignment a) { return ALN(a)->wrapText() ? 1 : 0; }
XLPP_API int xlpp_alignment_get_shrink_to_fit(xlpp_alignment a) { return ALN(a)->shrinkToFit() ? 1 : 0; }
XLPP_API int xlpp_alignment_get_text_rotation(xlpp_alignment a) { return ALN(a)->textRotation(); }
XLPP_API int xlpp_alignment_get_indent(xlpp_alignment a) { return ALN(a)->indent(); }

// ============================================================
// Style
// ============================================================
XLPP_API xlpp_font xlpp_style_font(xlpp_style s)            { return reinterpret_cast<xlpp_font>(&STY(s)->font()); }
XLPP_API xlpp_fill xlpp_style_fill(xlpp_style s)            { return reinterpret_cast<xlpp_fill>(&STY(s)->fill()); }
XLPP_API xlpp_border xlpp_style_border(xlpp_style s)        { return reinterpret_cast<xlpp_border>(&STY(s)->border()); }
XLPP_API xlpp_alignment xlpp_style_alignment(xlpp_style s)   { return reinterpret_cast<xlpp_alignment>(&STY(s)->alignment()); }
XLPP_API void xlpp_style_set_number_format(xlpp_style s, const char* v) { STY(s)->setNumberFormat(v); }
XLPP_API void xlpp_style_number_format(xlpp_style s, char* out, int outSize) { copyStr(STY(s)->numberFormat(), out, outSize); }
XLPP_API void xlpp_style_set_num_fmt_id(xlpp_style s, int v) { STY(s)->setNumFmtId(v); }
XLPP_API int xlpp_style_num_fmt_id(xlpp_style s) { return STY(s)->numFmtId(); }
XLPP_API void xlpp_style_set_locked(xlpp_style s, int v) { STY(s)->setLocked(v != 0); }
XLPP_API int xlpp_style_locked(xlpp_style s) { return STY(s)->locked() ? 1 : 0; }
XLPP_API void xlpp_style_set_hidden(xlpp_style s, int v) { STY(s)->setHidden(v != 0); }
XLPP_API int xlpp_style_hidden(xlpp_style s) { return STY(s)->hidden() ? 1 : 0; }
XLPP_API int xlpp_style_is_default(xlpp_style s) { return STY(s)->isDefault() ? 1 : 0; }

// ============================================================
// Streaming writer
// ============================================================

struct xlpp_stream_writer_t {
    std::unique_ptr<xlpp::StreamingWorkbookWriter> writer;
};

XLPP_API xlpp_stream_writer xlpp_stream_create(const char* path) {
    try {
        auto* handle = new xlpp_stream_writer_t;
        handle->writer = std::make_unique<xlpp::StreamingWorkbookWriter>(
            std::filesystem::path(path), xlpp::SharedStringMode::Disabled);
        return handle;
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_stream_destroy(xlpp_stream_writer w) {
    delete w;
}

XLPP_API uint64_t xlpp_stream_add_sheet(xlpp_stream_writer w, const char* name) {
    try {
        w->writer->addWorksheet(name);
        return static_cast<uint64_t>(w->writer->sheetCount() - 1);
    } catch (...) { return static_cast<uint64_t>(-1); }
}

XLPP_API void xlpp_stream_append_row(xlpp_stream_writer w, uint64_t sheetIndex, const char** values, int count) {
    try {
        auto& sheet = w->writer->worksheet(static_cast<std::size_t>(sheetIndex));
        std::vector<xlpp::CellValue> row;
        row.reserve(static_cast<std::size_t>(count));
        for (int i = 0; i < count; ++i) {
            if (values[i] && values[i][0]) row.push_back(std::string(values[i]));
            else row.push_back(xlpp::CellValue{});
        }
        sheet.append(row);
    } catch (...) {}
}

XLPP_API void xlpp_stream_append_doubles(xlpp_stream_writer w, uint64_t sheetIndex, const double* values, int count) {
    try {
        auto& sheet = w->writer->worksheet(static_cast<std::size_t>(sheetIndex));
        std::vector<xlpp::CellValue> row;
        row.reserve(static_cast<std::size_t>(count));
        for (int i = 0; i < count; ++i) row.push_back(values[i]);
        sheet.append(row);
    } catch (...) {}
}

XLPP_API uint64_t xlpp_stream_row_count(xlpp_stream_writer w, uint64_t sheetIndex) {
    try { return static_cast<uint64_t>(w->writer->worksheet(static_cast<std::size_t>(sheetIndex)).rowCount()); }
    catch (...) { return 0; }
}

XLPP_API uint64_t xlpp_stream_sheet_count(xlpp_stream_writer w) {
    try { return static_cast<uint64_t>(w->writer->sheetCount()); } catch (...) { return 0; }
}

XLPP_API void xlpp_stream_set_date1904(xlpp_stream_writer w, int v) {
    try { w->writer->setDate1904(v != 0); } catch (...) {}
}

XLPP_API void xlpp_stream_set_compression_level(xlpp_stream_writer w, int level) {
    try { w->writer->setCompressionLevel(static_cast<xlpp::CompressionLevel>(level)); } catch (...) {}
}

XLPP_API void xlpp_stream_set_parallel_workers(xlpp_stream_writer w, uint64_t workers) {
    try { w->writer->setParallelWorkers(static_cast<std::size_t>(workers)); } catch (...) {}
}

XLPP_API void xlpp_stream_close(xlpp_stream_writer w) {
    try { if (w && w->writer) w->writer->close(); } catch (...) {}
}

// ============================================================
// Streaming reader
// ============================================================

struct xlpp_stream_reader_t {
    std::unique_ptr<xlpp::StreamingWorkbookReader> reader;
};

XLPP_API xlpp_stream_reader xlpp_stream_reader_open(const char* path) {
    try {
        auto* handle = new xlpp_stream_reader_t;
        handle->reader = std::make_unique<xlpp::StreamingWorkbookReader>(std::filesystem::path(path));
        return handle;
    } catch (...) { return nullptr; }
}

XLPP_API void xlpp_stream_reader_destroy(xlpp_stream_reader r) {
    delete r;
}

XLPP_API int xlpp_stream_reader_sheet_count(xlpp_stream_reader r) {
    try { return static_cast<int>(r->reader->worksheetNames().size()); } catch (...) { return 0; }
}

XLPP_API void xlpp_stream_reader_sheet_name(xlpp_stream_reader r, int index, char* out, int outSize) {
    try {
        const auto& names = r->reader->worksheetNames();
        if (index >= 0 && static_cast<std::size_t>(index) < names.size())
            copyStr(names[static_cast<std::size_t>(index)], out, outSize);
        else if (out && outSize > 0) out[0] = '\0';
    } catch (...) { if (out && outSize > 0) out[0] = '\0'; }
}

XLPP_API int xlpp_stream_reader_read_sheet(xlpp_stream_reader r, int index,
                                           xlpp_stream_row_callback callback, void* user) {
    try {
        const auto& names = r->reader->worksheetNames();
        if (index < 0 || static_cast<std::size_t>(index) >= names.size()) return 0;
        const auto& name = names[static_cast<std::size_t>(index)];
        bool keepGoing = true;
        r->reader->forEachRow(name, [&](std::size_t rowNumber, const xlpp::StreamingRow& row) {
            if (!keepGoing) return false;
            std::vector<const char*> addresses, strings, formulas;
            std::vector<double> numbers;
            std::vector<int> valueTypes, styleIndexes;
            addresses.reserve(row.size());
            strings.reserve(row.size());
            formulas.reserve(row.size());
            numbers.reserve(row.size());
            valueTypes.reserve(row.size());
            styleIndexes.reserve(row.size());
            for (const auto& cell : row) {
                addresses.push_back(cell.address.c_str());
                int vt = XLPP_VALUE_EMPTY;
                if (std::holds_alternative<bool>(cell.value)) vt = XLPP_VALUE_BOOL;
                else if (std::holds_alternative<double>(cell.value)) vt = XLPP_VALUE_NUMBER;
                else if (std::holds_alternative<std::string>(cell.value)) vt = XLPP_VALUE_STRING;
                else if (std::holds_alternative<xlpp::CellError>(cell.value)) vt = XLPP_VALUE_ERROR;
                else if (std::holds_alternative<xlpp::DateTime>(cell.value)) vt = XLPP_VALUE_DATE;
                valueTypes.push_back(vt);
                if (auto* v = std::get_if<double>(&cell.value)) numbers.push_back(*v);
                else if (auto* v = std::get_if<xlpp::DateTime>(&cell.value)) numbers.push_back(xlpp::toExcelSerial(*v, false));
                else numbers.push_back(0.0);
                if (auto* v = std::get_if<std::string>(&cell.value)) strings.push_back(v->c_str());
                else strings.push_back("");
                formulas.push_back(cell.formula.c_str());
                styleIndexes.push_back(cell.styleIndex ? static_cast<int>(*cell.styleIndex) : -1);
            }
            keepGoing = callback(user, rowNumber, static_cast<int>(row.size()),
                                 addresses.data(), numbers.data(), valueTypes.data(),
                                 strings.data(), formulas.data(), styleIndexes.data()) != 0;
            return keepGoing;
        });
        return 1;
    } catch (...) { return 0; }
}

XLPP_API void xlpp_free_string(const char* str) {
    (void)str;
}

XLPP_API const char* xlpp_last_error(void) {
    return g_lastError.c_str();
}

XLPP_API void xlpp_clear_error(void) {
    clearError();
}

} // extern "C"
