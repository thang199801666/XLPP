#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <pybind11/stl/filesystem.h>
#include <pybind11/functional.h>
#include <pybind11/numpy.h>
#include <XLPP/XLPP.h>
#include <XLPP/Cell/RichText.h>
#include <sstream>
#include <datetime.h>

namespace py = pybind11;
using namespace xlpp;

// StableVector-backed model collections deliberately keep native element addresses
// stable.  Python collection properties preserve their historical list-like value
// semantics by returning copies, while add/get-by-name APIs return live handles.
template <class T>
static std::vector<T> stable_vector_copy(const StableVector<T>& values) {
    return std::vector<T>(values.begin(), values.end());
}

template <class T>
static void assign_stable_vector(StableVector<T>& destination, const std::vector<T>& values) {
    destination.clear();
    destination.reserve(values.size());
    for (const auto& value : values) destination.push_back(value);
}

// --- DateTime converters ---
static DateTime py_to_datetime(const py::object& obj) {
    if (py::isinstance<py::none>(obj))
        throw std::invalid_argument("Expected datetime, got None");
    if (PyDateTime_Check(obj.ptr())) {
        PyDateTime_DateTime* dt = reinterpret_cast<PyDateTime_DateTime*>(obj.ptr());
        return DateTime{
            PyDateTime_GET_YEAR(dt),
            PyDateTime_GET_MONTH(dt),
            PyDateTime_GET_DAY(dt),
            PyDateTime_DATE_GET_HOUR(dt),
            PyDateTime_DATE_GET_MINUTE(dt),
            static_cast<double>(PyDateTime_DATE_GET_SECOND(dt)) +
                PyDateTime_DATE_GET_MICROSECOND(dt) / 1000000.0
        };
    }
    if (PyDate_Check(obj.ptr())) {
        PyDateTime_Date* d = reinterpret_cast<PyDateTime_Date*>(obj.ptr());
        return DateTime{
            PyDateTime_GET_YEAR(d),
            PyDateTime_GET_MONTH(d),
            PyDateTime_GET_DAY(d)
        };
    }
    throw std::invalid_argument("Expected datetime.date or datetime.datetime");
}

static py::object datetime_to_py(const DateTime& dt) {
    int s = static_cast<int>(dt.second);
    int us = static_cast<int>((dt.second - s) * 1000000.0);
    auto m = py::module_::import("datetime");
    auto cls = m.attr("datetime");
    return cls(dt.year, dt.month, dt.day, dt.hour, dt.minute, s, us);
}

// --- CellValue converters ---
static CellValue py_to_cellvalue(const py::object& obj) {
    if (PyUnicode_Check(obj.ptr())) return obj.cast<std::string>();
    if (py::isinstance<py::none>(obj)) return std::monostate{};
    if (py::isinstance<py::bool_>(obj)) return obj.cast<bool>();
    if (py::isinstance<py::int_>(obj)) return obj.cast<double>();
    if (py::isinstance<py::float_>(obj)) return obj.cast<double>();
    if (PyDateTime_Check(obj.ptr()) || PyDate_Check(obj.ptr()))
        return py_to_datetime(obj);
    return obj.cast<std::string>();
}

static py::object cellvalue_to_py(const CellValue& v) {
    return std::visit([](auto&& arg) -> py::object {
        using T = std::decay_t<decltype(arg)>;
        if constexpr (std::is_same_v<T, std::monostate>) return py::none{};
        else if constexpr (std::is_same_v<T, bool>) return py::bool_(arg);
        else if constexpr (std::is_same_v<T, double>) return py::float_(arg);
        else if constexpr (std::is_same_v<T, std::string>) return py::str(arg);
        else if constexpr (std::is_same_v<T, CellError>) return py::str(toString(arg));
        else if constexpr (std::is_same_v<T, DateTime>) return datetime_to_py(arg);
        return py::none{};
    }, v);
}

static CellError py_to_cellerror(const py::object& obj) {
    if (py::isinstance<CellError>(obj)) return obj.cast<CellError>();
    if (PyUnicode_Check(obj.ptr())) return cellErrorFromString(obj.cast<std::string>());
    throw std::invalid_argument("Expected a CellError or an Excel error string like '#VALUE!'");
}

// --- NumPy bulk write: write a 2D numeric array to cells (row1, col1)+ ---
static void write_numpy_array(Worksheet& ws, const py::array_t<double>& arr,
                              std::size_t row0, std::size_t col0, bool transpose) {
    auto buf = arr.request();
    if (buf.ndim < 1 || buf.ndim > 2)
        throw std::invalid_argument("array must be 1D or 2D");
    const std::size_t rows = buf.ndim == 1 ? 1 : static_cast<std::size_t>(buf.shape[0]);
    const std::size_t cols = buf.ndim == 1 ? static_cast<std::size_t>(buf.shape[0])
                                           : static_cast<std::size_t>(buf.shape[1]);
    if (row0 == 0) row0 = 1;
    if (col0 == 0) col0 = 1;
    const double* data = static_cast<const double*>(buf.ptr);
    const py::ssize_t stride0 = buf.strides[0] / static_cast<py::ssize_t>(sizeof(double));
    const py::ssize_t stride1 = buf.ndim == 1 ? 1 : buf.strides[1] / static_cast<py::ssize_t>(sizeof(double));
    if (transpose) {
        for (std::size_t r = 0; r < rows; ++r)
            for (std::size_t c = 0; c < cols; ++c)
                ws.cell(row0 + c, col0 + r).setValue(data[r * stride0 + c * stride1]);
    } else {
        for (std::size_t r = 0; r < rows; ++r)
            for (std::size_t c = 0; c < cols; ++c)
                ws.cell(row0 + r, col0 + c).setValue(data[r * stride0 + c * stride1]);
    }
}

// --- NumPy bulk read: read a 2D numeric region into a numpy array ---
static py::array_t<double> read_numpy_array(const Worksheet& ws,
                                            std::size_t minRow, std::size_t minCol,
                                            std::size_t maxRow, std::size_t maxCol) {
    if (maxRow == 0) maxRow = ws.maxRow();
    if (maxCol == 0) maxCol = ws.maxColumn();
    if (minRow == 0) minRow = 1;
    if (minCol == 0) minCol = 1;
    const std::size_t rows = maxRow - minRow + 1;
    const std::size_t cols = maxCol - minCol + 1;
    py::array_t<double> result({rows, cols});
    auto buf = result.request();
    double* data = static_cast<double*>(buf.ptr);
    for (std::size_t r = 0; r < rows; ++r)
        for (std::size_t c = 0; c < cols; ++c) {
            const auto* cell = ws.tryCell(minRow + r, minCol + c);
            data[r * cols + c] = cell ? cell->numericValueOr(0.0) : 0.0;
        }
    return result;
}

// --- Module ---
PYBIND11_MODULE(xlpp, m) {
    PyDateTime_IMPORT;
    m.doc() = "XL++ - high-performance C++20 Excel xlsx library for Python";

    py::register_exception_translator([](std::exception_ptr p) {
        try { if (p) std::rethrow_exception(p); }
        catch (const std::invalid_argument& e) { PyErr_SetString(PyExc_ValueError, e.what()); }
        catch (const std::out_of_range& e)   { PyErr_SetString(PyExc_IndexError, e.what()); }
        catch (const std::runtime_error& e)  { PyErr_SetString(PyExc_RuntimeError, e.what()); }
    });

    py::enum_<ErrorCode>(m, "ErrorCode")
        .value("UNKNOWN", ErrorCode::Unknown)
        .value("INVALID_ARGUMENT", ErrorCode::InvalidArgument)
        .value("INVALID_STATE", ErrorCode::InvalidState)
        .value("IO", ErrorCode::Io)
        .value("PACKAGE", ErrorCode::Package)
        .value("XML", ErrorCode::Xml)
        .value("VALIDATION", ErrorCode::Validation)
        .value("ENCRYPTION", ErrorCode::Encryption)
        .value("FORMULA", ErrorCode::Formula)
        .value("UNSUPPORTED", ErrorCode::Unsupported)
        .value("CANCELLED", ErrorCode::Cancelled)
        .value("RESOURCE_LIMIT", ErrorCode::ResourceLimit);
    py::register_exception<Exception>(m, "XLPPError");

    // === Compression ===
    py::enum_<CompressionLevel>(m, "CompressionLevel")
        .value("STORE", CompressionLevel::Store)
        .value("FASTEST", CompressionLevel::Fastest)
        .value("DEFAULT", CompressionLevel::Default)
        .value("BEST", CompressionLevel::Best);

    py::enum_<CompressionStrategy>(m, "CompressionStrategy")
        .value("DEFAULT", CompressionStrategy::Default)
        .value("FILTERED", CompressionStrategy::Filtered)
        .value("HUFFMAN_ONLY", CompressionStrategy::HuffmanOnly)
        .value("RLE", CompressionStrategy::Rle)
        .value("FIXED", CompressionStrategy::Fixed);

    py::enum_<OfficeEncryptionMode>(m, "OfficeEncryptionMode")
        .value("NONE", OfficeEncryptionMode::None)
        .value("AGILE_AES256_SHA512", OfficeEncryptionMode::AgileAes256Sha512)
        .value("STANDARD_AES_SHA1", OfficeEncryptionMode::StandardAesSha1)
        .value("UNSUPPORTED", OfficeEncryptionMode::Unsupported);

    py::class_<SaveOptions>(m, "SaveOptions")
        .def(py::init<>())
        .def_readwrite("compression_level", &SaveOptions::compressionLevel)
        .def_readwrite("compression_strategy", &SaveOptions::compressionStrategy)
        .def_readwrite("parallel_workers", &SaveOptions::parallelWorkers)
        .def_readwrite("parallel_sheets", &SaveOptions::parallelSheets)
        .def_readwrite("parallel_rows", &SaveOptions::parallelRows)
        .def_readwrite("strict_namespace", &SaveOptions::strictNamespace)
        .def_readwrite("synchronize_chart_caches", &SaveOptions::synchronizeChartCaches)
        .def_readwrite("synchronize_changed_chart_caches_only", &SaveOptions::synchronizeChangedChartCachesOnly)
        .def_readwrite("calculate_formulas_before_save", &SaveOptions::calculateFormulasBeforeSave)
        .def_readwrite("atomic_write", &SaveOptions::atomicWrite)
        .def_readwrite("durable_write", &SaveOptions::durableWrite)
        .def_readwrite("validate_before_save", &SaveOptions::validateBeforeSave)
        .def_readwrite("encryption_password", &SaveOptions::encryptionPassword)
        .def_readwrite("encryption_mode", &SaveOptions::encryptionMode)
        .def_readwrite("encryption_spin_count", &SaveOptions::encryptionSpinCount)
        .def_readwrite("encryption_key_bits", &SaveOptions::encryptionKeyBits);

    py::class_<LoadOptions>(m, "LoadOptions")
        .def(py::init<>())
        .def_readwrite("lenient", &LoadOptions::lenient)
        .def_readwrite("max_entries", &LoadOptions::maxEntries)
        .def_readwrite("max_entry_bytes", &LoadOptions::maxEntryBytes)
        .def_readwrite("max_total_bytes", &LoadOptions::maxTotalBytes)
        .def_readwrite("max_file_bytes", &LoadOptions::maxFileBytes)
        .def_readwrite("cancel", &LoadOptions::cancel)
        .def_readwrite("progress", &LoadOptions::progress)
        .def_readwrite("password", &LoadOptions::password)
        .def_readwrite("verify_encryption_integrity", &LoadOptions::verifyEncryptionIntegrity);

    // === Formula calculation / structural editing / encryption ===
    py::class_<CalculationCell>(m, "CalculationCell")
        .def(py::init<>())
        .def_readwrite("sheet", &CalculationCell::sheet)
        .def_readwrite("cell", &CalculationCell::cell);

    py::class_<CalculationOptions>(m, "CalculationOptions")
        .def(py::init<>())
        .def_readwrite("recursive_dependencies", &CalculationOptions::recursiveDependencies)
        .def_readwrite("update_cached_values", &CalculationOptions::updateCachedValues)
        .def_readwrite("evaluate_volatile_functions", &CalculationOptions::evaluateVolatileFunctions)
        .def_readwrite("spill_dynamic_arrays", &CalculationOptions::spillDynamicArrays)
        .def_readwrite("iterative_calculation", &CalculationOptions::iterativeCalculation)
        .def_readwrite("max_iterations", &CalculationOptions::maxIterations)
        .def_readwrite("max_change", &CalculationOptions::maxChange)
        .def_readwrite("external_reference_resolver", &CalculationOptions::externalReferenceResolver)
        .def_readwrite("max_depth", &CalculationOptions::maxDepth)
        .def_readwrite("changed_cells", &CalculationOptions::changedCells);

    py::class_<CalculationReport>(m, "CalculationReport")
        .def(py::init<>())
        .def_readwrite("formula_cells_visited", &CalculationReport::formulaCellsVisited)
        .def_readwrite("formula_cells_evaluated", &CalculationReport::formulaCellsEvaluated)
        .def_readwrite("cached_values_updated", &CalculationReport::cachedValuesUpdated)
        .def_readwrite("dependency_evaluations", &CalculationReport::dependencyEvaluations)
        .def_readwrite("defined_names_resolved", &CalculationReport::definedNamesResolved)
        .def_readwrite("circular_references", &CalculationReport::circularReferences)
        .def_readwrite("unsupported_formulas", &CalculationReport::unsupportedFormulas)
        .def_readwrite("evaluation_errors", &CalculationReport::evaluationErrors)
        .def_readwrite("dynamic_arrays_spilled", &CalculationReport::dynamicArraysSpilled)
        .def_readwrite("spill_cells_updated", &CalculationReport::spillCellsUpdated)
        .def_readwrite("spill_conflicts", &CalculationReport::spillConflicts)
        .def_readwrite("structured_references_resolved", &CalculationReport::structuredReferencesResolved)
        .def_readwrite("iterative_iterations", &CalculationReport::iterativeIterations)
        .def_readwrite("iterative_convergence_failures", &CalculationReport::iterativeConvergenceFailures)
        .def_readwrite("external_references_resolved", &CalculationReport::externalReferencesResolved)
        .def_readwrite("unresolved_external_references", &CalculationReport::unresolvedExternalReferences)
        .def_readonly("dirty_roots", &CalculationReport::dirtyRoots)
        .def_readonly("dirty_formula_cells_selected", &CalculationReport::dirtyFormulaCellsSelected)
        .def_readonly("warnings", &CalculationReport::warnings)
        .def_property_readonly("success", &CalculationReport::success);

    py::enum_<StructuralEditKind>(m, "StructuralEditKind")
        .value("INSERT_ROWS", StructuralEditKind::InsertRows)
        .value("DELETE_ROWS", StructuralEditKind::DeleteRows)
        .value("INSERT_COLUMNS", StructuralEditKind::InsertColumns)
        .value("DELETE_COLUMNS", StructuralEditKind::DeleteColumns);

    py::class_<StructuralEdit>(m, "StructuralEdit")
        .def(py::init<>())
        .def(py::init([](std::string sheetName, StructuralEditKind kind, std::size_t index, std::size_t amount) {
                return StructuralEdit{std::move(sheetName), kind, index, amount};
            }), py::arg("sheet_name"), py::arg("kind"), py::arg("index"), py::arg("amount") = 1)
        .def_readwrite("sheet_name", &StructuralEdit::sheetName)
        .def_readwrite("kind", &StructuralEdit::kind)
        .def_readwrite("index", &StructuralEdit::index)
        .def_readwrite("amount", &StructuralEdit::amount);

    py::class_<ReferenceTranslationResult>(m, "ReferenceTranslationResult")
        .def(py::init<>())
        .def_readwrite("value", &ReferenceTranslationResult::value)
        .def_readwrite("references_visited", &ReferenceTranslationResult::referencesVisited)
        .def_readwrite("references_changed", &ReferenceTranslationResult::referencesChanged)
        .def_readwrite("references_invalidated", &ReferenceTranslationResult::referencesInvalidated)
        .def_property_readonly("changed", &ReferenceTranslationResult::changed);

    py::class_<StructuralEditOptions>(m, "StructuralEditOptions")
        .def(py::init<>())
        .def_readwrite("transactional", &StructuralEditOptions::transactional)
        .def_readwrite("update_defined_names", &StructuralEditOptions::updateDefinedNames)
        .def_readwrite("recalculate_formulas", &StructuralEditOptions::recalculateFormulas)
        .def_readwrite("synchronize_chart_caches", &StructuralEditOptions::synchronizeChartCaches)
        .def_readwrite("changed_chart_caches_only", &StructuralEditOptions::changedChartCachesOnly)
        .def_readwrite("fail_on_invalid_reference", &StructuralEditOptions::failOnInvalidReference);

    py::class_<StructuralEditReport>(m, "StructuralEditReport")
        .def(py::init<>())
        .def_readwrite("worksheets_visited", &StructuralEditReport::worksheetsVisited)
        .def_readwrite("cells_moved", &StructuralEditReport::cellsMoved)
        .def_readwrite("cells_removed", &StructuralEditReport::cellsRemoved)
        .def_readwrite("formulas_updated", &StructuralEditReport::formulasUpdated)
        .def_readwrite("formula_metadata_updated", &StructuralEditReport::formulaMetadataUpdated)
        .def_readwrite("worksheet_references_updated", &StructuralEditReport::worksheetReferencesUpdated)
        .def_readwrite("defined_names_updated", &StructuralEditReport::definedNamesUpdated)
        .def_readwrite("chart_references_updated", &StructuralEditReport::chartReferencesUpdated)
        .def_readwrite("pivot_references_updated", &StructuralEditReport::pivotReferencesUpdated)
        .def_readwrite("drawing_anchors_updated", &StructuralEditReport::drawingAnchorsUpdated)
        .def_readwrite("hyperlinks_updated", &StructuralEditReport::hyperlinksUpdated)
        .def_readwrite("references_invalidated", &StructuralEditReport::referencesInvalidated)
        .def_readwrite("formulas_calculated", &StructuralEditReport::formulasCalculated)
        .def_readwrite("chart_caches_updated", &StructuralEditReport::chartCachesUpdated)
        .def_readwrite("warnings", &StructuralEditReport::warnings)
        .def_property_readonly("success", &StructuralEditReport::success);

    py::class_<WorksheetStructuralEditReport>(m, "WorksheetStructuralEditReport")
        .def(py::init<>())
        .def_readwrite("cells_moved", &WorksheetStructuralEditReport::cellsMoved)
        .def_readwrite("cells_removed", &WorksheetStructuralEditReport::cellsRemoved)
        .def_readwrite("formulas_updated", &WorksheetStructuralEditReport::formulasUpdated)
        .def_readwrite("formula_metadata_updated", &WorksheetStructuralEditReport::formulaMetadataUpdated)
        .def_readwrite("worksheet_references_updated", &WorksheetStructuralEditReport::worksheetReferencesUpdated)
        .def_readwrite("references_invalidated", &WorksheetStructuralEditReport::referencesInvalidated)
        .def_readwrite("drawing_anchors_updated", &WorksheetStructuralEditReport::drawingAnchorsUpdated)
        .def_readwrite("chart_references_updated", &WorksheetStructuralEditReport::chartReferencesUpdated)
        .def_readwrite("pivot_references_updated", &WorksheetStructuralEditReport::pivotReferencesUpdated)
        .def_readwrite("hyperlinks_updated", &WorksheetStructuralEditReport::hyperlinksUpdated);

    py::class_<WorksheetRenameOptions>(m, "WorksheetRenameOptions")
        .def(py::init<>())
        .def_readwrite("recalculate_formulas", &WorksheetRenameOptions::recalculateFormulas)
        .def_readwrite("synchronize_chart_caches", &WorksheetRenameOptions::synchronizeChartCaches)
        .def_readwrite("changed_chart_caches_only", &WorksheetRenameOptions::changedChartCachesOnly);

    py::class_<WorksheetRenameReport>(m, "WorksheetRenameReport")
        .def(py::init<>())
        .def_readwrite("worksheets_visited", &WorksheetRenameReport::worksheetsVisited)
        .def_readwrite("formulas_updated", &WorksheetRenameReport::formulasUpdated)
        .def_readwrite("formula_metadata_updated", &WorksheetRenameReport::formulaMetadataUpdated)
        .def_readwrite("defined_names_updated", &WorksheetRenameReport::definedNamesUpdated)
        .def_readwrite("chart_references_updated", &WorksheetRenameReport::chartReferencesUpdated)
        .def_readwrite("pivot_references_updated", &WorksheetRenameReport::pivotReferencesUpdated)
        .def_readwrite("hyperlinks_updated", &WorksheetRenameReport::hyperlinksUpdated)
        .def_readwrite("references_updated", &WorksheetRenameReport::referencesUpdated)
        .def_readwrite("formulas_calculated", &WorksheetRenameReport::formulasCalculated)
        .def_readwrite("chart_caches_updated", &WorksheetRenameReport::chartCachesUpdated)
        .def_readwrite("warnings", &WorksheetRenameReport::warnings)
        .def_property_readonly("success", &WorksheetRenameReport::success);

    py::class_<ChartCacheSyncOptions>(m, "ChartCacheSyncOptions")
        .def(py::init<>())
        .def_readwrite("synchronize_titles", &ChartCacheSyncOptions::synchronizeTitles)
        .def_readwrite("synchronize_categories", &ChartCacheSyncOptions::synchronizeCategories)
        .def_readwrite("synchronize_values", &ChartCacheSyncOptions::synchronizeValues)
        .def_readwrite("changed_references_only", &ChartCacheSyncOptions::changedReferencesOnly)
        .def_readwrite("clear_unsupported_references", &ChartCacheSyncOptions::clearUnsupportedReferences);

    py::class_<ChartCacheSyncReport>(m, "ChartCacheSyncReport")
        .def(py::init<>())
        .def_readwrite("charts_visited", &ChartCacheSyncReport::chartsVisited)
        .def_readwrite("series_visited", &ChartCacheSyncReport::seriesVisited)
        .def_readwrite("references_checked", &ChartCacheSyncReport::referencesChecked)
        .def_readwrite("references_unchanged", &ChartCacheSyncReport::referencesUnchanged)
        .def_readwrite("dependencies_registered", &ChartCacheSyncReport::dependenciesRegistered)
        .def_readwrite("dependencies_changed", &ChartCacheSyncReport::dependenciesChanged)
        .def_readwrite("caches_updated", &ChartCacheSyncReport::cachesUpdated)
        .def_readwrite("caches_cleared", &ChartCacheSyncReport::cachesCleared)
        .def_readwrite("references_skipped", &ChartCacheSyncReport::referencesSkipped)
        .def_readwrite("warnings", &ChartCacheSyncReport::warnings)
        .def_property_readonly("success", &ChartCacheSyncReport::success);

    py::enum_<WorkbookValidationSeverity>(m, "WorkbookValidationSeverity")
        .value("WARNING", WorkbookValidationSeverity::Warning)
        .value("ERROR", WorkbookValidationSeverity::Error);

    py::class_<WorkbookValidationIssue>(m, "WorkbookValidationIssue")
        .def(py::init<>())
        .def_readwrite("severity", &WorkbookValidationIssue::severity)
        .def_readwrite("code", &WorkbookValidationIssue::code)
        .def_readwrite("message", &WorkbookValidationIssue::message)
        .def_readwrite("worksheet", &WorkbookValidationIssue::worksheet);

    py::class_<WorkbookValidationOptions>(m, "WorkbookValidationOptions")
        .def(py::init<>())
        .def_readwrite("validate_worksheet_names", &WorkbookValidationOptions::validateWorksheetNames)
        .def_readwrite("validate_defined_names", &WorkbookValidationOptions::validateDefinedNames)
        .def_readwrite("validate_tables", &WorkbookValidationOptions::validateTables)
        .def_readwrite("validate_pivots", &WorkbookValidationOptions::validatePivots)
        .def_readwrite("validate_charts", &WorkbookValidationOptions::validateCharts)
        .def_readwrite("validate_vba", &WorkbookValidationOptions::validateVba);

    py::class_<WorkbookValidationReport>(m, "WorkbookValidationReport")
        .def(py::init<>())
        .def_readwrite("issues", &WorkbookValidationReport::issues)
        .def_readwrite("error_count", &WorkbookValidationReport::errorCount)
        .def_readwrite("warning_count", &WorkbookValidationReport::warningCount)
        .def_property_readonly("ok", &WorkbookValidationReport::ok)
        .def("__bool__", &WorkbookValidationReport::ok);

    py::enum_<VbaModuleType>(m, "VbaModuleType")
        .value("STANDARD", VbaModuleType::Standard)
        .value("DOCUMENT", VbaModuleType::Document)
        .value("CLASS", VbaModuleType::Class);

    py::class_<VbaModule>(m, "VbaModule")
        .def(py::init<>())
        .def_readwrite("name", &VbaModule::name)
        .def_readwrite("source", &VbaModule::source)
        .def_readwrite("type", &VbaModule::type)
        .def_readwrite("read_only", &VbaModule::readOnly)
        .def_readwrite("private_module", &VbaModule::privateModule);

    py::class_<VbaProjectProperties>(m, "VbaProjectProperties")
        .def(py::init<>())
        .def_readwrite("name", &VbaProjectProperties::name)
        .def_readwrite("description", &VbaProjectProperties::description)
        .def_readwrite("help_file", &VbaProjectProperties::helpFile)
        .def_readwrite("help_context_id", &VbaProjectProperties::helpContextId)
        .def_readwrite("constants", &VbaProjectProperties::constants);

    py::class_<OfficeEncryptionInfo>(m, "OfficeEncryptionInfo")
        .def(py::init<>())
        .def_readwrite("encrypted", &OfficeEncryptionInfo::encrypted)
        .def_readwrite("supported", &OfficeEncryptionInfo::supported)
        .def_readwrite("mode", &OfficeEncryptionInfo::mode)
        .def_readwrite("spin_count", &OfficeEncryptionInfo::spinCount)
        .def_readwrite("key_bits", &OfficeEncryptionInfo::keyBits)
        .def_readwrite("cipher_algorithm", &OfficeEncryptionInfo::cipherAlgorithm)
        .def_readwrite("hash_algorithm", &OfficeEncryptionInfo::hashAlgorithm);

    py::enum_<FormulaDependencyKind>(m, "FormulaDependencyKind")
        .value("CELL_OR_RANGE", FormulaDependencyKind::CellOrRange)
        .value("DEFINED_NAME", FormulaDependencyKind::DefinedName)
        .value("TABLE", FormulaDependencyKind::Table)
        .value("EXTERNAL_REFERENCE", FormulaDependencyKind::ExternalReference)
        .value("VOLATILE_REFERENCE", FormulaDependencyKind::VolatileReference);

    py::class_<FormulaDependency>(m, "FormulaDependency")
        .def(py::init<>())
        .def_readwrite("dependent_sheet", &FormulaDependency::dependentSheet)
        .def_readwrite("dependent_cell", &FormulaDependency::dependentCell)
        .def_readwrite("kind", &FormulaDependency::kind)
        .def_readwrite("precedent_sheet", &FormulaDependency::precedentSheet)
        .def_readwrite("precedent_reference", &FormulaDependency::precedentReference)
        .def_readwrite("symbol", &FormulaDependency::symbol);

    py::class_<FormulaDependencyReport>(m, "FormulaDependencyReport")
        .def(py::init<>())
        .def_readwrite("formula_cells", &FormulaDependencyReport::formulaCells)
        .def_readwrite("edges", &FormulaDependencyReport::edges)
        .def_readwrite("cell_or_range_edges", &FormulaDependencyReport::cellOrRangeEdges)
        .def_readwrite("defined_name_edges", &FormulaDependencyReport::definedNameEdges)
        .def_readwrite("table_edges", &FormulaDependencyReport::tableEdges)
        .def_readwrite("external_edges", &FormulaDependencyReport::externalEdges)
        .def_readwrite("volatile_references", &FormulaDependencyReport::volatileReferences)
        .def_readwrite("unresolved_symbols", &FormulaDependencyReport::unresolvedSymbols)
        .def_readwrite("warnings", &FormulaDependencyReport::warnings);

    py::class_<FormulaDependencyGraph>(m, "FormulaDependencyGraph")
        .def(py::init<>())
        .def_property_readonly("edges", [](const FormulaDependencyGraph& graph) { return graph.edges(); })
        .def_property_readonly("report", &FormulaDependencyGraph::report, py::return_value_policy::reference_internal)
        .def("precedents_of", &FormulaDependencyGraph::precedentsOf)
        .def("dependents_of", &FormulaDependencyGraph::dependentsOf)
        .def("depends_on", &FormulaDependencyGraph::dependsOn);

    // === CellReference ===
    py::class_<CellReference>(m, "CellReference")
        .def(py::init<>())
        .def(py::init<std::size_t, std::size_t>(), py::arg("row"), py::arg("column"))
        .def_readwrite("row", &CellReference::row)
        .def_readwrite("column", &CellReference::column)
        .def_static("parse", &CellReference::parse, py::arg("address"))
        .def_static("column_name", &CellReference::columnName, py::arg("column"))
        .def_static("column_index", [](const std::string& name) { return CellReference::columnIndex(name); },
                    py::arg("name"))
        .def("address", &CellReference::address)
        .def("__eq__", [](const CellReference& a, const CellReference& b) {
            return a.row == b.row && a.column == b.column;
        })
        .def("__repr__", [](const CellReference& r) { return "<CellRef " + r.address() + ">"; });

    m.attr("MAX_EXCEL_ROWS") = py::int_(MaxExcelRows);
    m.attr("MAX_EXCEL_COLUMNS") = py::int_(MaxExcelColumns);
    m.attr("MAX_WORKSHEET_NAME_CHARACTERS") = py::int_(MaxWorksheetNameCharacters);
    m.def("is_valid_cell_coordinate", &isValidCellCoordinate, py::arg("row"), py::arg("column"));
    m.def("make_cell_key", &makeCellKey, py::arg("row"), py::arg("column"));
    m.def("is_valid_worksheet_name", &isValidWorksheetName, py::arg("name"));
    m.def("validate_worksheet_name", &validateWorksheetName, py::arg("name"));
    m.def("worksheet_names_equal", &worksheetNamesEqual, py::arg("left"), py::arg("right"));

    // === DateTime (aggregate) ===
    py::class_<DateTime>(m, "DateTime")
        .def(py::init<>())
        .def(py::init<int, int, int, int, int, double>(),
             py::arg("year"), py::arg("month"), py::arg("day"),
             py::arg("hour") = 0, py::arg("minute") = 0, py::arg("second") = 0.0)
        .def_readwrite("year", &DateTime::year)
        .def_readwrite("month", &DateTime::month)
        .def_readwrite("day", &DateTime::day)
        .def_readwrite("hour", &DateTime::hour)
        .def_readwrite("minute", &DateTime::minute)
        .def_readwrite("second", &DateTime::second)
        .def_property_readonly("second_int", [](const DateTime& d) { return static_cast<int>(d.second); })
        .def_property_readonly("millisecond", [](const DateTime& d) {
            return static_cast<int>((d.second - static_cast<int>(d.second)) * 1000.0);
        })
        .def("to_iso8601_date", [](const DateTime& d) { return toIso8601Date(d); })
        .def("__eq__", [](const DateTime& a, const DateTime& b) { return a == b; })
        .def("__ne__", [](const DateTime& a, const DateTime& b) { return a != b; })
        .def("__str__", [](const DateTime& d) { return toIso8601(d); })
        .def("__repr__", [](const DateTime& d) { return "<DateTime " + toIso8601(d) + ">"; });

    m.def("from_excel_serial", [](double serial, bool date1904) { return fromExcelSerial(serial, date1904); },
          py::arg("serial"), py::arg("date1904") = false,
          "Convert an Excel serial day number to a DateTime.");
    m.def("to_excel_serial", [](const DateTime& d, bool date1904) { return toExcelSerial(d, date1904); },
          py::arg("value"), py::arg("date1904") = false);
    m.def("parse_iso8601", [](const std::string& text) -> std::optional<DateTime> {
            return parseIso8601(text);
          }, py::arg("text"));
    m.def("is_date_format_code", [](const std::string& format, int numFmtId) {
            return isDateFormatCode(format, numFmtId);
          }, py::arg("format"), py::arg("num_fmt_id") = -1);

    // === Formula metadata / CellError ===
    py::enum_<FormulaType>(m, "FormulaType")
        .value("NORMAL", FormulaType::Normal)
        .value("SHARED", FormulaType::Shared)
        .value("ARRAY", FormulaType::Array)
        .value("DATA_TABLE", FormulaType::DataTable)
        .value("DYNAMIC_ARRAY", FormulaType::DynamicArray);

    py::class_<FormulaMetadata>(m, "FormulaMetadata")
        .def(py::init<>())
        .def_property("type", &FormulaMetadata::type, &FormulaMetadata::setType)
        .def_property("reference", &FormulaMetadata::reference, &FormulaMetadata::setReference)
        .def("clear_reference", &FormulaMetadata::clearReference)
        .def_property("shared_index",
            [](const FormulaMetadata& f) -> std::optional<unsigned> { return f.sharedIndex(); },
            [](FormulaMetadata& f, std::optional<unsigned> v) {
                if (v) f.setSharedIndex(*v); else f.clearSharedIndex();
            })
        .def("clear_shared_index", &FormulaMetadata::clearSharedIndex)
        .def_property("always_calculate_array", &FormulaMetadata::alwaysCalculateArray, &FormulaMetadata::setAlwaysCalculateArray)
        .def_property("calculate_on_load", &FormulaMetadata::calculateOnLoad, &FormulaMetadata::setCalculateOnLoad)
        .def("empty", &FormulaMetadata::empty);

    py::enum_<CellError>(m, "CellError")
        .value("NULL", CellError::Null)
        .value("DIVISION_BY_ZERO", CellError::DivisionByZero)
        .value("VALUE", CellError::Value)
        .value("REFERENCE", CellError::Reference)
        .value("NAME", CellError::Name)
        .value("NUMBER", CellError::Number)
        .value("NOT_AVAILABLE", CellError::NotAvailable)
        .value("GETTING_DATA", CellError::GettingData)
        .value("SPILL", CellError::Spill)
        .value("CALCULATION", CellError::Calculation);

    m.def("cell_error_to_string", [](CellError e) { return toString(e); }, py::arg("error"));
    m.def("cell_error_from_string", [](const std::string& s) { return cellErrorFromString(s); }, py::arg("value"));

    // === Color ===
    py::class_<Color>(m, "Color")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("argb"))
        .def("set_argb", &Color::setArgb, py::return_value_policy::reference_internal)
        .def_property("argb", &Color::argb, &Color::setArgb)
        .def("empty", &Color::empty)
        .def("hash", &Color::hash)
        .def("__eq__", [](const Color& a, const Color& b) { return a == b; });

    // === Font ===
    py::class_<Font>(m, "Font")
        .def(py::init<>())
        .def_property("name", &Font::name, [](Font& f, std::string v) -> Font& { f.setName(std::move(v)); return f; })
        .def_property("size", &Font::size, [](Font& f, double v) -> Font& { f.setSize(v); return f; })
        .def_property("bold", &Font::bold, [](Font& f, bool v) -> Font& { f.setBold(v); return f; })
        .def_property("italic", &Font::italic, [](Font& f, bool v) -> Font& { f.setItalic(v); return f; })
        .def_property("underline", &Font::underline, [](Font& f, bool v) -> Font& { f.setUnderline(v); return f; })
        .def_property("strike", &Font::strike, [](Font& f, bool v) -> Font& { f.setStrike(v); return f; })
        .def("color", static_cast<Color& (Font::*)()>(&Font::color), py::return_value_policy::reference_internal)
        .def("hash", &Font::hash)
        .def("__eq__", [](const Font& a, const Font& b) { return a == b; });

    // === Fill ===
    py::class_<Fill>(m, "Fill")
        .def(py::init<>())
        .def_property("pattern_type", &Fill::patternType,
            [](Fill& f, std::string v) -> Fill& { f.setPatternType(std::move(v)); return f; })
        .def("foreground", static_cast<Color& (Fill::*)()>(&Fill::foregroundColor), py::return_value_policy::reference_internal)
        .def("background", static_cast<Color& (Fill::*)()>(&Fill::backgroundColor), py::return_value_policy::reference_internal)
        .def("hash", &Fill::hash)
        .def("__eq__", [](const Fill& a, const Fill& b) { return a == b; });

    // === BorderSide ===
    py::class_<BorderSide>(m, "BorderSide")
        .def(py::init<>())
        .def_property("style", &BorderSide::style,
            [](BorderSide& b, std::string v) -> BorderSide& { b.setStyle(std::move(v)); return b; })
        .def("color", static_cast<Color& (BorderSide::*)()>(&BorderSide::color), py::return_value_policy::reference_internal)
        .def("hash", &BorderSide::hash)
        .def("__eq__", [](const BorderSide& a, const BorderSide& b) { return a == b; });

    // === Border ===
    py::class_<Border>(m, "Border")
        .def(py::init<>())
        .def("left", static_cast<BorderSide& (Border::*)()>(&Border::left), py::return_value_policy::reference_internal)
        .def("right", static_cast<BorderSide& (Border::*)()>(&Border::right), py::return_value_policy::reference_internal)
        .def("top", static_cast<BorderSide& (Border::*)()>(&Border::top), py::return_value_policy::reference_internal)
        .def("bottom", static_cast<BorderSide& (Border::*)()>(&Border::bottom), py::return_value_policy::reference_internal)
        .def("diagonal", static_cast<BorderSide& (Border::*)()>(&Border::diagonal), py::return_value_policy::reference_internal)
        .def("hash", &Border::hash)
        .def("__eq__", [](const Border& a, const Border& b) { return a == b; });

    // === Alignment ===
    py::class_<Alignment>(m, "Alignment")
        .def(py::init<>())
        .def_property("horizontal", &Alignment::horizontal,
            [](Alignment& a, std::string v) -> Alignment& { a.setHorizontal(std::move(v)); return a; })
        .def_property("vertical", &Alignment::vertical,
            [](Alignment& a, std::string v) -> Alignment& { a.setVertical(std::move(v)); return a; })
        .def_property("wrap_text", &Alignment::wrapText,
            [](Alignment& a, bool v) -> Alignment& { a.setWrapText(v); return a; })
        .def_property("shrink_to_fit", &Alignment::shrinkToFit,
            [](Alignment& a, bool v) -> Alignment& { a.setShrinkToFit(v); return a; })
        .def_property("text_rotation", &Alignment::textRotation,
            [](Alignment& a, int v) -> Alignment& { a.setTextRotation(v); return a; })
        .def_property("indent", &Alignment::indent,
            [](Alignment& a, int v) -> Alignment& { a.setIndent(v); return a; })
        .def("hash", &Alignment::hash)
        .def("__eq__", [](const Alignment& a, const Alignment& b) { return a == b; });

    // === Style ===
    py::class_<Style>(m, "Style")
        .def(py::init<>())
        .def("font", static_cast<Font& (Style::*)()>(&Style::font), py::return_value_policy::reference_internal)
        .def("fill", static_cast<Fill& (Style::*)()>(&Style::fill), py::return_value_policy::reference_internal)
        .def("border", static_cast<Border& (Style::*)()>(&Style::border), py::return_value_policy::reference_internal)
        .def("alignment", static_cast<Alignment& (Style::*)()>(&Style::alignment), py::return_value_policy::reference_internal)
        .def_property("number_format", &Style::numberFormat, &Style::setNumberFormat)
        .def_property("num_fmt_id", &Style::numFmtId, &Style::setNumFmtId)
        .def_property("locked", &Style::locked, &Style::setLocked)
        .def_property("hidden", &Style::hidden, &Style::setHidden)
        .def("is_default", &Style::isDefault)
        .def("hash", &Style::hash)
        .def("__eq__", [](const Style& a, const Style& b) { return a == b; });

    // === Hyperlink ===
    py::class_<Hyperlink>(m, "Hyperlink")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("target"))
        .def_property("target", &Hyperlink::target, &Hyperlink::setTarget)
        .def_property("display", &Hyperlink::display, &Hyperlink::setDisplay)
        .def_property("tooltip", &Hyperlink::tooltip, &Hyperlink::setTooltip)
        .def_property("external", &Hyperlink::external, &Hyperlink::setExternal);

    // === Comment ===
    py::class_<Comment>(m, "Comment")
        .def(py::init<>())
        .def(py::init<std::string, std::string>(), py::arg("text"), py::arg("author") = "XL++")
        .def_property("text", &Comment::text, &Comment::setText)
        .def_property("author", &Comment::author, &Comment::setAuthor);

    // === RichText ===
    py::class_<RichTextRun>(m, "RichTextRun")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("text"))
        .def_property("text", &RichTextRun::text, &RichTextRun::setText)
        .def_property("bold", &RichTextRun::bold, &RichTextRun::setBold)
        .def_property("italic", &RichTextRun::italic, &RichTextRun::setItalic)
        .def_property("underline", &RichTextRun::underline, &RichTextRun::setUnderline)
        .def_property("strike", &RichTextRun::strike, &RichTextRun::setStrike)
        .def_property("color", &RichTextRun::color, &RichTextRun::setColor)
        .def_property("size", &RichTextRun::size, &RichTextRun::setSize)
        .def_property("font_name", &RichTextRun::fontName, &RichTextRun::setFontName);

    py::class_<RichText>(m, "RichText")
        .def(py::init<>())
        .def("add_run", &RichText::addRun)
        .def_property("runs",
            [](const RichText& rt) { return stable_vector_copy(rt.runs()); },
            [](RichText& rt, const std::vector<RichTextRun>& v) { assign_stable_vector(rt.runs(), v); },
            py::return_value_policy::reference_internal)
        .def("empty", &RichText::empty)
        .def("plain_text", &RichText::plainText)
        .def_static("from_plain", &RichText::fromPlain, py::arg("text"));

    // === DocumentProperties ===
    py::class_<DocumentProperties>(m, "DocumentProperties")
        .def(py::init<>())
        .def_property("title", &DocumentProperties::title, &DocumentProperties::setTitle)
        .def_property("subject", &DocumentProperties::subject, &DocumentProperties::setSubject)
        .def_property("creator", &DocumentProperties::creator, &DocumentProperties::setCreator)
        .def_property("description", &DocumentProperties::description, &DocumentProperties::setDescription)
        .def_property("keywords", &DocumentProperties::keywords, &DocumentProperties::setKeywords)
        .def_property("category", &DocumentProperties::category, &DocumentProperties::setCategory)
        .def_property("last_modified_by", &DocumentProperties::lastModifiedBy, &DocumentProperties::setLastModifiedBy);

    // === AutoFilter ===
    py::enum_<FilterOperator>(m, "FilterOperator")
        .value("EQUAL", FilterOperator::Equal)
        .value("NOT_EQUAL", FilterOperator::NotEqual)
        .value("LESS_THAN", FilterOperator::LessThan)
        .value("LESS_THAN_OR_EQUAL", FilterOperator::LessThanOrEqual)
        .value("GREATER_THAN", FilterOperator::GreaterThan)
        .value("GREATER_THAN_OR_EQUAL", FilterOperator::GreaterThanOrEqual);

    py::class_<CustomFilter>(m, "CustomFilter")
        .def(py::init<>())
        .def_readwrite("op", &CustomFilter::op)
        .def_readwrite("value", &CustomFilter::value);

    py::enum_<DynamicFilterType>(m, "DynamicFilterType")
        .value("ABOVE_AVERAGE", DynamicFilterType::AboveAverage).value("BELOW_AVERAGE", DynamicFilterType::BelowAverage)
        .value("TODAY", DynamicFilterType::Today).value("YESTERDAY", DynamicFilterType::Yesterday).value("TOMORROW", DynamicFilterType::Tomorrow)
        .value("THIS_WEEK", DynamicFilterType::ThisWeek).value("LAST_WEEK", DynamicFilterType::LastWeek).value("NEXT_WEEK", DynamicFilterType::NextWeek)
        .value("THIS_MONTH", DynamicFilterType::ThisMonth).value("LAST_MONTH", DynamicFilterType::LastMonth).value("NEXT_MONTH", DynamicFilterType::NextMonth)
        .value("THIS_YEAR", DynamicFilterType::ThisYear).value("LAST_YEAR", DynamicFilterType::LastYear).value("NEXT_YEAR", DynamicFilterType::NextYear)
        .value("YEAR_TO_DATE", DynamicFilterType::YearToDate);

    py::enum_<DateTimeGrouping>(m, "DateTimeGrouping")
        .value("YEAR", DateTimeGrouping::Year).value("MONTH", DateTimeGrouping::Month).value("DAY", DateTimeGrouping::Day)
        .value("HOUR", DateTimeGrouping::Hour).value("MINUTE", DateTimeGrouping::Minute).value("SECOND", DateTimeGrouping::Second);

    py::class_<DynamicFilter>(m, "DynamicFilter")
        .def(py::init<>()).def_readwrite("type", &DynamicFilter::type).def_readwrite("value", &DynamicFilter::value).def_readwrite("max_value", &DynamicFilter::maxValue);
    py::class_<Top10Filter>(m, "Top10Filter")
        .def(py::init<>()).def_readwrite("top", &Top10Filter::top).def_readwrite("percent", &Top10Filter::percent)
        .def_readwrite("value", &Top10Filter::value).def_readwrite("filter_value", &Top10Filter::filterValue);
    py::class_<ColorFilter>(m, "ColorFilter")
        .def(py::init<>()).def_readwrite("dxf_id", &ColorFilter::dxfId).def_readwrite("cell_color", &ColorFilter::cellColor);
    py::class_<IconFilter>(m, "IconFilter")
        .def(py::init<>()).def_readwrite("icon_set", &IconFilter::iconSet).def_readwrite("icon_id", &IconFilter::iconId);
    py::class_<DateGroupItem>(m, "DateGroupItem")
        .def(py::init<>()).def_readwrite("year", &DateGroupItem::year).def_readwrite("month", &DateGroupItem::month)
        .def_readwrite("day", &DateGroupItem::day).def_readwrite("hour", &DateGroupItem::hour).def_readwrite("minute", &DateGroupItem::minute)
        .def_readwrite("second", &DateGroupItem::second).def_readwrite("grouping", &DateGroupItem::grouping);

    py::class_<FilterColumn>(m, "FilterColumn")
        .def(py::init<>())
        .def(py::init<std::size_t>(), py::arg("column_id"))
        .def_property("column_id", &FilterColumn::columnId, &FilterColumn::setColumnId)
        .def("add_value", &FilterColumn::addValue)
        .def("clear_values", &FilterColumn::clearValues)
        .def_property_readonly("values", &FilterColumn::values)
        .def("add_custom_filter", &FilterColumn::addCustomFilter)
        .def("clear_custom_filters", &FilterColumn::clearCustomFilters)
        .def_property_readonly("custom_filters", &FilterColumn::customFilters)
        .def("add_date_group", &FilterColumn::addDateGroup)
        .def("clear_date_groups", &FilterColumn::clearDateGroups)
        .def_property_readonly("date_groups", &FilterColumn::dateGroups)
        .def_property("and_mode", &FilterColumn::andMode, &FilterColumn::setAndMode)
        .def_property("include_blank", &FilterColumn::includeBlank, &FilterColumn::setIncludeBlank)
        .def_property("dynamic_filter", &FilterColumn::dynamicFilter, [](FilterColumn& c, const std::optional<DynamicFilter>& v){ if(v)c.setDynamicFilter(*v); else c.clearDynamicFilter(); })
        .def("set_dynamic_filter", &FilterColumn::setDynamicFilter)
        .def("clear_dynamic_filter", &FilterColumn::clearDynamicFilter)
        .def_property("top10_filter", &FilterColumn::top10Filter, [](FilterColumn& c, const std::optional<Top10Filter>& v){ if(v)c.setTop10Filter(*v); else c.clearTop10Filter(); })
        .def("set_top10_filter", &FilterColumn::setTop10Filter)
        .def("clear_top10_filter", &FilterColumn::clearTop10Filter)
        .def_property("color_filter", &FilterColumn::colorFilter, [](FilterColumn& c, const std::optional<ColorFilter>& v){ if(v)c.setColorFilter(*v); else c.clearColorFilter(); })
        .def("set_color_filter", &FilterColumn::setColorFilter)
        .def("clear_color_filter", &FilterColumn::clearColorFilter)
        .def_property("icon_filter", &FilterColumn::iconFilter, [](FilterColumn& c, const std::optional<IconFilter>& v){ if(v)c.setIconFilter(*v); else c.clearIconFilter(); })
        .def("set_icon_filter", &FilterColumn::setIconFilter)
        .def("clear_icon_filter", &FilterColumn::clearIconFilter);

    py::class_<SortCondition>(m, "SortCondition")
        .def(py::init<>())
        .def_readwrite("reference", &SortCondition::reference)
        .def_readwrite("descending", &SortCondition::descending);

    py::class_<SortState>(m, "SortState")
        .def(py::init<>())
        .def("set_reference", &SortState::setReference)
        .def_property("reference", &SortState::reference, &SortState::setReference)
        .def_property("case_sensitive", &SortState::caseSensitive, &SortState::setCaseSensitive)
        .def("add_condition", &SortState::addCondition, py::arg("ref"), py::arg("descending") = false)
        .def_property_readonly("conditions", &SortState::conditions)
        .def("clear", &SortState::clear);

    py::class_<AutoFilter>(m, "AutoFilter")
        .def(py::init<>())
        .def("set_reference", &AutoFilter::setReference)
        .def_property("reference", &AutoFilter::reference, &AutoFilter::setReference)
        .def_property_readonly("enabled", &AutoFilter::enabled)
        .def("clear", &AutoFilter::clear)
        .def("column", &AutoFilter::column, py::return_value_policy::reference_internal)
        .def("try_column", [](const AutoFilter& f, std::size_t id) -> py::object {
            const auto* c = f.tryColumn(id);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def_property_readonly("columns", &AutoFilter::columns)
        .def("sort_state", [](AutoFilter& f) -> SortState& { return f.sortState(); },
             py::return_value_policy::reference_internal)
        .def_property_readonly("sort_state_value", &AutoFilter::sortStateValue);

    // === PageSetup ===
    py::enum_<PageOrientation>(m, "PageOrientation")
        .value("DEFAULT", PageOrientation::Default)
        .value("PORTRAIT", PageOrientation::Portrait)
        .value("LANDSCAPE", PageOrientation::Landscape);

    py::enum_<PaperSize>(m, "PaperSize")
        .value("DEFAULT", PaperSize::Default)
        .value("LETTER", PaperSize::Letter)
        .value("LEGAL", PaperSize::Legal)
        .value("A4", PaperSize::A4)
        .value("A3", PaperSize::A3);

    py::class_<PageSetup>(m, "PageSetup")
        .def(py::init<>())
        .def_property("orientation", &PageSetup::orientation, &PageSetup::setOrientation)
        .def_property("paper_size", &PageSetup::paperSize, &PageSetup::setPaperSize)
        .def_property("scale", &PageSetup::scale, &PageSetup::setScale)
        .def_property("fit_to_width", &PageSetup::fitToWidth, &PageSetup::setFitToWidth)
        .def_property("fit_to_height", &PageSetup::fitToHeight, &PageSetup::setFitToHeight)
        .def_property("fit_to_page", &PageSetup::fitToPage, &PageSetup::setFitToPage)
        .def_property("black_and_white", &PageSetup::blackAndWhite, &PageSetup::setBlackAndWhite)
        .def_property("draft", &PageSetup::draft, &PageSetup::setDraft)
        .def_property("first_page_number", &PageSetup::firstPageNumber, &PageSetup::setFirstPageNumber)
        .def_property("use_first_page_number", &PageSetup::useFirstPageNumber, &PageSetup::setUseFirstPageNumber);

    py::class_<PageMargins>(m, "PageMargins")
        .def(py::init<>())
        .def_property("left", &PageMargins::left, &PageMargins::setLeft)
        .def_property("right", &PageMargins::right, &PageMargins::setRight)
        .def_property("top", &PageMargins::top, &PageMargins::setTop)
        .def_property("bottom", &PageMargins::bottom, &PageMargins::setBottom)
        .def_property("header", &PageMargins::header, &PageMargins::setHeader)
        .def_property("footer", &PageMargins::footer, &PageMargins::setFooter);

    py::class_<PrintOptions>(m, "PrintOptions")
        .def(py::init<>())
        .def_property("horizontal_centered", &PrintOptions::horizontalCentered, &PrintOptions::setHorizontalCentered)
        .def_property("vertical_centered", &PrintOptions::verticalCentered, &PrintOptions::setVerticalCentered)
        .def_property("headings", &PrintOptions::headings, &PrintOptions::setHeadings)
        .def_property("grid_lines", &PrintOptions::gridLines, &PrintOptions::setGridLines);

    py::class_<HeaderFooter>(m, "HeaderFooter")
        .def(py::init<>())
        .def_property("odd_header", &HeaderFooter::oddHeader, &HeaderFooter::setOddHeader)
        .def_property("odd_footer", &HeaderFooter::oddFooter, &HeaderFooter::setOddFooter)
        .def_property("even_header", &HeaderFooter::evenHeader, &HeaderFooter::setEvenHeader)
        .def_property("even_footer", &HeaderFooter::evenFooter, &HeaderFooter::setEvenFooter)
        .def_property("different_odd_even", &HeaderFooter::differentOddEven, &HeaderFooter::setDifferentOddEven)
        .def_property("different_first", &HeaderFooter::differentFirst, &HeaderFooter::setDifferentFirst);

    // === Drawing metadata / Image ===
    py::enum_<DrawingAnchorType>(m, "DrawingAnchorType")
        .value("ONE_CELL", DrawingAnchorType::OneCell).value("TWO_CELL", DrawingAnchorType::TwoCell)
        .value("ABSOLUTE", DrawingAnchorType::Absolute);
    py::class_<DrawingMarker>(m, "DrawingMarker")
        .def(py::init<>()).def_readwrite("row", &DrawingMarker::row).def_readwrite("column", &DrawingMarker::column)
        .def_readwrite("row_offset_emu", &DrawingMarker::rowOffsetEmu).def_readwrite("column_offset_emu", &DrawingMarker::columnOffsetEmu);
    py::class_<DrawingAnchorInfo>(m, "DrawingAnchorInfo")
        .def(py::init<>()).def_readwrite("type", &DrawingAnchorInfo::type).def_readwrite("from_marker", &DrawingAnchorInfo::from)
        .def_readwrite("to_marker", &DrawingAnchorInfo::to).def_readwrite("x_emu", &DrawingAnchorInfo::xEmu)
        .def_readwrite("y_emu", &DrawingAnchorInfo::yEmu).def_readwrite("width_emu", &DrawingAnchorInfo::widthEmu)
        .def_readwrite("height_emu", &DrawingAnchorInfo::heightEmu).def_readwrite("edit_as", &DrawingAnchorInfo::editAs);
    py::class_<Image>(m, "Image")
        .def(py::init<>())
        .def(py::init<std::string, std::vector<unsigned char>, std::string>(),
             py::arg("anchor"), py::arg("bytes"), py::arg("extension"))
        .def_static("from_file", &Image::fromFile)
        .def_property("anchor", &Image::anchor, &Image::setAnchor)
        .def_property_readonly("bytes", &Image::bytes)
        .def_property_readonly("extension", &Image::extension)
        .def_property("width_pixels", &Image::widthPixels, &Image::setWidthPixels)
        .def_property("height_pixels", &Image::heightPixels, &Image::setHeightPixels)
        .def_property("name", &Image::name, &Image::setName)
        .def_property("anchor_info", &Image::anchorInfo, &Image::setAnchorInfo)
        .def_property("stable_id", &Image::stableId, &Image::setStableId)
        .def_property("source_drawing_part", &Image::sourceDrawingPart, &Image::setSourceDrawingPart)
        .def_property("source_media_part", &Image::sourceMediaPart, &Image::setSourceMediaPart)
        .def_property("source_relationship_id", &Image::sourceRelationshipId, &Image::setSourceRelationshipId)
        .def_property("imported", &Image::imported, &Image::setImported);

    // === Chart ===
    py::enum_<Chart::Type>(m, "ChartType")
        .value("BAR", Chart::Type::Bar).value("LINE", Chart::Type::Line).value("PIE", Chart::Type::Pie)
        .value("SCATTER", Chart::Type::Scatter).value("DOUGHNUT", Chart::Type::Doughnut)
        .value("RADAR", Chart::Type::Radar).value("AREA", Chart::Type::Area).value("BUBBLE", Chart::Type::Bubble)
        .value("STOCK", Chart::Type::Stock).value("BAR_3D", Chart::Type::Bar3D).value("LINE_3D", Chart::Type::Line3D)
        .value("AREA_3D", Chart::Type::Area3D).value("PIE_3D", Chart::Type::Pie3D).value("SURFACE", Chart::Type::Surface)
        .value("SURFACE_3D", Chart::Type::Surface3D).value("PIE_OF_PIE", Chart::Type::PieOfPie)
        .value("BAR_OF_PIE", Chart::Type::BarOfPie).value("HORIZONTAL_BAR", Chart::Type::HorizontalBar)
        .value("HISTOGRAM", Chart::Type::Histogram).value("PARETO", Chart::Type::Pareto)
        .value("BOX_WHISKER", Chart::Type::BoxWhisker).value("WATERFALL", Chart::Type::Waterfall)
        .value("FUNNEL", Chart::Type::Funnel).value("TREEMAP", Chart::Type::Treemap)
        .value("SUNBURST", Chart::Type::Sunburst).value("FILLED_MAP", Chart::Type::FilledMap);
    py::enum_<Chart::Grouping>(m, "ChartGrouping")
        .value("STANDARD", Chart::Grouping::Standard).value("STACKED", Chart::Grouping::Stacked)
        .value("PERCENT_STACKED", Chart::Grouping::PercentStacked).value("CLUSTERED", Chart::Grouping::Clustered);
    py::enum_<Chart::BarDirection>(m, "ChartBarDirection")
        .value("COLUMN", Chart::BarDirection::Column).value("BAR", Chart::BarDirection::Bar);
    py::enum_<Chart::ScatterStyle>(m, "ChartScatterStyle")
        .value("NONE", Chart::ScatterStyle::None).value("LINE", Chart::ScatterStyle::Line)
        .value("LINE_MARKER", Chart::ScatterStyle::LineMarker).value("MARKER", Chart::ScatterStyle::Marker)
        .value("SMOOTH", Chart::ScatterStyle::Smooth).value("SMOOTH_MARKER", Chart::ScatterStyle::SmoothMarker);
    py::enum_<Chart::BubbleSizeRepresents>(m, "ChartBubbleSizeRepresents")
        .value("AREA", Chart::BubbleSizeRepresents::Area).value("WIDTH", Chart::BubbleSizeRepresents::Width);

    py::enum_<ChartColorTransform::Kind>(m, "ChartColorTransformKind")
        .value("ALPHA", ChartColorTransform::Kind::Alpha).value("ALPHA_MOD", ChartColorTransform::Kind::AlphaMod)
        .value("ALPHA_OFF", ChartColorTransform::Kind::AlphaOff).value("TINT", ChartColorTransform::Kind::Tint)
        .value("SHADE", ChartColorTransform::Kind::Shade).value("LUM_MOD", ChartColorTransform::Kind::LumMod)
        .value("LUM_OFF", ChartColorTransform::Kind::LumOff).value("SAT_MOD", ChartColorTransform::Kind::SatMod)
        .value("SAT_OFF", ChartColorTransform::Kind::SatOff);
    py::class_<ChartColorTransform>(m, "ChartColorTransform")
        .def(py::init<>()).def_readwrite("kind", &ChartColorTransform::kind).def_readwrite("value", &ChartColorTransform::value);
    py::enum_<ChartColor::Kind>(m, "ChartColorKind")
        .value("NONE", ChartColor::Kind::None).value("SRGB", ChartColor::Kind::SRgb)
        .value("SCHEME", ChartColor::Kind::Scheme).value("SYSTEM", ChartColor::Kind::System)
        .value("PRESET", ChartColor::Kind::Preset).value("UNKNOWN", ChartColor::Kind::Unknown);
    py::class_<ChartColor>(m, "ChartColor")
        .def(py::init<>()).def_readwrite("kind", &ChartColor::kind).def_readwrite("value", &ChartColor::value)
        .def_readwrite("transforms", &ChartColor::transforms).def_property_readonly("present", &ChartColor::present);
    py::class_<ChartCustomDashStop>(m, "ChartCustomDashStop")
        .def(py::init<>()).def_readwrite("dash", &ChartCustomDashStop::dash).def_readwrite("space", &ChartCustomDashStop::space);
    py::class_<ChartLineFormat>(m, "ChartLineFormat")
        .def(py::init<>()).def_readwrite("present", &ChartLineFormat::present).def_readwrite("no_fill", &ChartLineFormat::noFill)
        .def_readwrite("color", &ChartLineFormat::color).def_readwrite("width_points", &ChartLineFormat::widthPoints)
        .def_readwrite("dash", &ChartLineFormat::dash).def_readwrite("cap", &ChartLineFormat::cap)
        .def_readwrite("compound", &ChartLineFormat::compound).def_readwrite("join", &ChartLineFormat::join)
        .def_readwrite("custom_dash", &ChartLineFormat::customDash);
    py::class_<ChartGradientStop>(m, "ChartGradientStop")
        .def(py::init<>()).def_readwrite("position", &ChartGradientStop::position).def_readwrite("color", &ChartGradientStop::color);
    py::enum_<ChartFillFormat::Kind>(m, "ChartFillKind")
        .value("NONE", ChartFillFormat::Kind::None).value("NO_FILL", ChartFillFormat::Kind::NoFill)
        .value("SOLID", ChartFillFormat::Kind::Solid).value("GRADIENT", ChartFillFormat::Kind::Gradient)
        .value("PATTERN", ChartFillFormat::Kind::Pattern);
    py::class_<ChartFillFormat>(m, "ChartFillFormat")
        .def(py::init<>()).def_readwrite("present", &ChartFillFormat::present).def_readwrite("no_fill", &ChartFillFormat::noFill)
        .def_readwrite("color", &ChartFillFormat::color).def_readwrite("kind", &ChartFillFormat::kind)
        .def_readwrite("gradient_stops", &ChartFillFormat::gradientStops).def_readwrite("gradient_angle_degrees", &ChartFillFormat::gradientAngleDegrees)
        .def_readwrite("pattern", &ChartFillFormat::pattern).def_readwrite("foreground_color", &ChartFillFormat::foregroundColor)
        .def_readwrite("background_color", &ChartFillFormat::backgroundColor);
    py::class_<ChartTextRun>(m, "ChartTextRun")
        .def(py::init<>()).def_readwrite("text", &ChartTextRun::text).def_readwrite("bold", &ChartTextRun::bold)
        .def_readwrite("italic", &ChartTextRun::italic).def_readwrite("font_size_points", &ChartTextRun::fontSizePoints)
        .def_readwrite("typeface", &ChartTextRun::typeface).def_readwrite("color", &ChartTextRun::color);
    py::class_<ChartTextStyle>(m, "ChartTextStyle")
        .def(py::init<>()).def_readwrite("present", &ChartTextStyle::present).def_readwrite("bold", &ChartTextStyle::bold)
        .def_readwrite("italic", &ChartTextStyle::italic).def_readwrite("font_size_points", &ChartTextStyle::fontSizePoints)
        .def_readwrite("typeface", &ChartTextStyle::typeface).def_readwrite("color", &ChartTextStyle::color);
    py::class_<ChartRichText>(m, "ChartRichText")
        .def(py::init<>()).def_readwrite("present", &ChartRichText::present).def_readwrite("runs", &ChartRichText::runs)
        .def_property_readonly("plain_text", &ChartRichText::plainText);
    py::class_<ChartCachePoint>(m, "ChartCachePoint")
        .def(py::init<>()).def_readwrite("index", &ChartCachePoint::index).def_readwrite("value", &ChartCachePoint::value);
    py::class_<ChartSeriesCache>(m, "ChartSeriesCache")
        .def(py::init<>()).def_readwrite("present", &ChartSeriesCache::present).def_readwrite("numeric", &ChartSeriesCache::numeric)
        .def_readwrite("format_code", &ChartSeriesCache::formatCode).def_readwrite("point_count", &ChartSeriesCache::pointCount)
        .def_readwrite("points", &ChartSeriesCache::points).def_property_readonly("effective_point_count", &ChartSeriesCache::effectivePointCount)
        .def("valid", &ChartSeriesCache::valid, py::arg("allow_sparse") = true);
    py::class_<ChartManualLayout>(m, "ChartManualLayout")
        .def(py::init<>()).def_readwrite("present", &ChartManualLayout::present).def_readwrite("target", &ChartManualLayout::target)
        .def_readwrite("x_mode", &ChartManualLayout::xMode).def_readwrite("y_mode", &ChartManualLayout::yMode)
        .def_readwrite("width_mode", &ChartManualLayout::widthMode).def_readwrite("height_mode", &ChartManualLayout::heightMode)
        .def_readwrite("has_x", &ChartManualLayout::hasX).def_readwrite("has_y", &ChartManualLayout::hasY)
        .def_readwrite("has_width", &ChartManualLayout::hasWidth).def_readwrite("has_height", &ChartManualLayout::hasHeight)
        .def_readwrite("x", &ChartManualLayout::x).def_readwrite("y", &ChartManualLayout::y)
        .def_readwrite("width", &ChartManualLayout::width).def_readwrite("height", &ChartManualLayout::height);
    py::class_<ChartAxisScaling>(m, "ChartAxisScaling")
        .def(py::init<>()).def_readwrite("has_minimum", &ChartAxisScaling::hasMinimum).def_readwrite("has_maximum", &ChartAxisScaling::hasMaximum)
        .def_readwrite("has_log_base", &ChartAxisScaling::hasLogBase).def_readwrite("minimum", &ChartAxisScaling::minimum)
        .def_readwrite("maximum", &ChartAxisScaling::maximum).def_readwrite("log_base", &ChartAxisScaling::logBase)
        .def_readwrite("reverse_order", &ChartAxisScaling::reverseOrder);
    py::class_<ChartDataLabelPoint>(m, "ChartDataLabelPoint")
        .def(py::init<>()).def_readwrite("index", &ChartDataLabelPoint::index).def_readwrite("deleted", &ChartDataLabelPoint::deleted)
        .def_readwrite("show_value", &ChartDataLabelPoint::showValue).def_readwrite("show_category_name", &ChartDataLabelPoint::showCategoryName)
        .def_readwrite("show_series_name", &ChartDataLabelPoint::showSeriesName).def_readwrite("position", &ChartDataLabelPoint::position)
        .def_readwrite("separator", &ChartDataLabelPoint::separator).def_readwrite("rich_text", &ChartDataLabelPoint::richText);
    py::class_<ChartDisplayUnits>(m, "ChartDisplayUnits")
        .def(py::init<>()).def_readwrite("present", &ChartDisplayUnits::present).def_readwrite("built_in_unit", &ChartDisplayUnits::builtInUnit)
        .def_readwrite("has_custom_unit", &ChartDisplayUnits::hasCustomUnit).def_readwrite("custom_unit", &ChartDisplayUnits::customUnit)
        .def_readwrite("show_label", &ChartDisplayUnits::showLabel).def_readwrite("label_rich_text", &ChartDisplayUnits::labelRichText);
    py::class_<ChartDataPointFormat>(m, "ChartDataPointFormat")
        .def(py::init<>()).def_readwrite("index", &ChartDataPointFormat::index).def_readwrite("fill", &ChartDataPointFormat::fill)
        .def_readwrite("line", &ChartDataPointFormat::line).def_readwrite("marker", &ChartDataPointFormat::marker);
    py::class_<ChartMarkerFormat>(m, "ChartMarkerFormat")
        .def(py::init<>()).def_readwrite("present", &ChartMarkerFormat::present).def_readwrite("symbol", &ChartMarkerFormat::symbol)
        .def_readwrite("size", &ChartMarkerFormat::size).def_readwrite("fill", &ChartMarkerFormat::fill).def_readwrite("line", &ChartMarkerFormat::line);
    py::class_<ChartView3D>(m, "ChartView3D")
        .def(py::init<>()).def_readwrite("present", &ChartView3D::present).def_readwrite("has_rotation_x", &ChartView3D::hasRotationX)
        .def_readwrite("has_rotation_y", &ChartView3D::hasRotationY).def_readwrite("rotation_x", &ChartView3D::rotationX)
        .def_readwrite("rotation_y", &ChartView3D::rotationY).def_readwrite("height_percent", &ChartView3D::heightPercent)
        .def_readwrite("depth_percent", &ChartView3D::depthPercent).def_readwrite("right_angle_axes", &ChartView3D::rightAngleAxes)
        .def_readwrite("perspective", &ChartView3D::perspective);
    py::class_<ChartWallFormat>(m, "ChartWallFormat")
        .def(py::init<>()).def_readwrite("present", &ChartWallFormat::present).def_readwrite("has_thickness", &ChartWallFormat::hasThickness)
        .def_readwrite("thickness", &ChartWallFormat::thickness).def_readwrite("fill", &ChartWallFormat::fill).def_readwrite("line", &ChartWallFormat::line);
    py::class_<ChartUpDownBars>(m, "ChartUpDownBars")
        .def(py::init<>()).def_readwrite("present", &ChartUpDownBars::present).def_readwrite("gap_width", &ChartUpDownBars::gapWidth)
        .def_readwrite("up_fill", &ChartUpDownBars::upFill).def_readwrite("up_line", &ChartUpDownBars::upLine)
        .def_readwrite("down_fill", &ChartUpDownBars::downFill).def_readwrite("down_line", &ChartUpDownBars::downLine);
    py::class_<ChartProjectedPieOptions>(m, "ChartProjectedPieOptions")
        .def(py::init<>()).def_readwrite("present", &ChartProjectedPieOptions::present).def_readwrite("of_pie_type", &ChartProjectedPieOptions::ofPieType)
        .def_readwrite("gap_width", &ChartProjectedPieOptions::gapWidth).def_readwrite("split_type", &ChartProjectedPieOptions::splitType)
        .def_readwrite("has_split_position", &ChartProjectedPieOptions::hasSplitPosition).def_readwrite("split_position", &ChartProjectedPieOptions::splitPosition)
        .def_readwrite("custom_split_points", &ChartProjectedPieOptions::customSplitPoints).def_readwrite("second_plot_size", &ChartProjectedPieOptions::secondPlotSize)
        .def_readwrite("has_series_lines", &ChartProjectedPieOptions::hasSeriesLines).def_readwrite("series_lines_format", &ChartProjectedPieOptions::seriesLinesFormat);
    py::class_<ChartDataLabels>(m, "ChartDataLabels")
        .def(py::init<>()).def_readwrite("present", &ChartDataLabels::present).def_readwrite("show_value", &ChartDataLabels::showValue)
        .def_readwrite("show_category_name", &ChartDataLabels::showCategoryName).def_readwrite("show_series_name", &ChartDataLabels::showSeriesName)
        .def_readwrite("show_percent", &ChartDataLabels::showPercent).def_readwrite("show_bubble_size", &ChartDataLabels::showBubbleSize)
        .def_readwrite("show_leader_lines", &ChartDataLabels::showLeaderLines).def_readwrite("has_leader_lines", &ChartDataLabels::hasLeaderLines)
        .def_readwrite("leader_line_format", &ChartDataLabels::leaderLineFormat).def_readwrite("position", &ChartDataLabels::position)
        .def_readwrite("separator", &ChartDataLabels::separator).def_readwrite("points", &ChartDataLabels::points);
    py::class_<ChartDataTable>(m, "ChartDataTable")
        .def(py::init<>()).def_readwrite("present", &ChartDataTable::present).def_readwrite("show_horizontal_border", &ChartDataTable::showHorizontalBorder)
        .def_readwrite("show_vertical_border", &ChartDataTable::showVerticalBorder).def_readwrite("show_outline", &ChartDataTable::showOutline)
        .def_readwrite("show_legend_keys", &ChartDataTable::showLegendKeys).def_readwrite("fill", &ChartDataTable::fill)
        .def_readwrite("line", &ChartDataTable::line).def_readwrite("text_style", &ChartDataTable::textStyle);

    py::enum_<Chart::AxisKind>(m, "ChartAxisKind")
        .value("CATEGORY", Chart::AxisKind::Category).value("VALUE", Chart::AxisKind::Value)
        .value("DATE", Chart::AxisKind::Date).value("SERIES", Chart::AxisKind::Series);
    py::class_<Chart::Axis>(m, "ChartAxis")
        .def(py::init<>()).def_readwrite("kind", &Chart::Axis::kind).def_readwrite("id", &Chart::Axis::id)
        .def_readwrite("cross_axis_id", &Chart::Axis::crossAxisId).def_readwrite("position", &Chart::Axis::position)
        .def_readwrite("title", &Chart::Axis::title).def_readwrite("title_rich_text", &Chart::Axis::titleRichText)
        .def_readwrite("secondary", &Chart::Axis::secondary).def_readwrite("number_format", &Chart::Axis::numberFormat)
        .def_readwrite("number_format_source_linked", &Chart::Axis::numberFormatSourceLinked)
        .def_readwrite("major_tick_mark", &Chart::Axis::majorTickMark).def_readwrite("minor_tick_mark", &Chart::Axis::minorTickMark)
        .def_readwrite("tick_label_position", &Chart::Axis::tickLabelPosition).def_readwrite("has_major_unit", &Chart::Axis::hasMajorUnit)
        .def_readwrite("has_minor_unit", &Chart::Axis::hasMinorUnit).def_readwrite("major_unit", &Chart::Axis::majorUnit)
        .def_readwrite("minor_unit", &Chart::Axis::minorUnit).def_readwrite("crosses", &Chart::Axis::crosses)
        .def_readwrite("cross_between", &Chart::Axis::crossBetween).def_readwrite("has_crosses_at", &Chart::Axis::hasCrossesAt)
        .def_readwrite("crosses_at", &Chart::Axis::crossesAt).def_readwrite("scaling", &Chart::Axis::scaling)
        .def_readwrite("display_units", &Chart::Axis::displayUnits).def_readwrite("has_major_gridlines", &Chart::Axis::hasMajorGridlines)
        .def_readwrite("has_minor_gridlines", &Chart::Axis::hasMinorGridlines).def_readwrite("line_format", &Chart::Axis::lineFormat)
        .def_readwrite("major_gridline_format", &Chart::Axis::majorGridlineFormat).def_readwrite("minor_gridline_format", &Chart::Axis::minorGridlineFormat);
    py::class_<Chart::Plot>(m, "ChartPlot")
        .def(py::init<>())
        .def_readwrite("type", &Chart::Plot::type)
        .def_readwrite("grouping", &Chart::Plot::grouping)
        .def_readwrite("first_series", &Chart::Plot::firstSeries)
        .def_readwrite("series_count", &Chart::Plot::seriesCount)
        .def_readwrite("uses_secondary_axes", &Chart::Plot::usesSecondaryAxes)
        .def_readwrite("axis_ids", &Chart::Plot::axisIds)
        .def_readwrite("data_labels", &Chart::Plot::dataLabels)
        .def_readwrite("has_drop_lines", &Chart::Plot::hasDropLines).def_readwrite("drop_lines_format", &Chart::Plot::dropLinesFormat)
        .def_readwrite("has_high_low_lines", &Chart::Plot::hasHighLowLines).def_readwrite("high_low_lines_format", &Chart::Plot::highLowLinesFormat)
        .def_readwrite("up_down_bars", &Chart::Plot::upDownBars).def_readwrite("has_gap_depth", &Chart::Plot::hasGapDepth)
        .def_readwrite("gap_depth", &Chart::Plot::gapDepth).def_readwrite("has_wireframe", &Chart::Plot::hasWireframe)
        .def_readwrite("wireframe", &Chart::Plot::wireframe).def_readwrite("shape", &Chart::Plot::shape)
        .def_readwrite("has_first_slice_angle", &Chart::Plot::hasFirstSliceAngle).def_readwrite("first_slice_angle", &Chart::Plot::firstSliceAngle)
        .def_readwrite("has_hole_size", &Chart::Plot::hasHoleSize).def_readwrite("hole_size", &Chart::Plot::holeSize)
        .def_readwrite("radar_style", &Chart::Plot::radarStyle).def_readwrite("projected_pie", &Chart::Plot::projectedPie)
        .def_readwrite("bar_direction", &Chart::Plot::barDirection)
        .def_readwrite("scatter_style", &Chart::Plot::scatterStyle)
        .def_readwrite("has_bubble_scale", &Chart::Plot::hasBubbleScale)
        .def_readwrite("bubble_scale", &Chart::Plot::bubbleScale)
        .def_readwrite("show_negative_bubbles", &Chart::Plot::showNegativeBubbles)
        .def_readwrite("bubble_size_represents", &Chart::Plot::bubbleSizeRepresents)
        .def_readwrite("bubble_3d", &Chart::Plot::bubble3D)
        .def_readwrite("histogram_bin_width", &Chart::Plot::histogramBinWidth)
        .def_readwrite("histogram_bin_count", &Chart::Plot::histogramBinCount)
        .def_readwrite("histogram_automatic_bins", &Chart::Plot::histogramAutomaticBins)
        .def_readwrite("histogram_has_underflow", &Chart::Plot::histogramHasUnderflow)
        .def_readwrite("histogram_underflow", &Chart::Plot::histogramUnderflow)
        .def_readwrite("histogram_has_overflow", &Chart::Plot::histogramHasOverflow)
        .def_readwrite("histogram_overflow", &Chart::Plot::histogramOverflow)
        .def_readwrite("box_whisker_show_inner_points", &Chart::Plot::boxWhiskerShowInnerPoints)
        .def_readwrite("box_whisker_show_outlier_points", &Chart::Plot::boxWhiskerShowOutlierPoints)
        .def_readwrite("box_whisker_show_mean_line", &Chart::Plot::boxWhiskerShowMeanLine)
        .def_readwrite("box_whisker_show_mean_marker", &Chart::Plot::boxWhiskerShowMeanMarker)
        .def_readwrite("box_whisker_quartile_inclusive", &Chart::Plot::boxWhiskerQuartileInclusive)
        .def_readwrite("waterfall_show_connector_lines", &Chart::Plot::waterfallShowConnectorLines)
        .def_readwrite("map_projection", &Chart::Plot::mapProjection)
        .def_readwrite("map_area", &Chart::Plot::mapArea)
        .def_readwrite("map_labels", &Chart::Plot::mapLabels);

    py::enum_<ChartSeries::TrendlineType>(m, "ChartTrendlineType")
        .value("LINEAR", ChartSeries::TrendlineType::Linear).value("EXPONENTIAL", ChartSeries::TrendlineType::Exponential)
        .value("LOGARITHMIC", ChartSeries::TrendlineType::Logarithmic).value("POLYNOMIAL", ChartSeries::TrendlineType::Polynomial)
        .value("POWER", ChartSeries::TrendlineType::Power).value("MOVING_AVERAGE", ChartSeries::TrendlineType::MovingAverage);
    py::enum_<ChartSeries::ErrorBarDirection>(m, "ChartErrorBarDirection")
        .value("X", ChartSeries::ErrorBarDirection::X).value("Y", ChartSeries::ErrorBarDirection::Y);
    py::enum_<ChartSeries::ErrorBarType>(m, "ChartErrorBarType")
        .value("BOTH", ChartSeries::ErrorBarType::Both).value("PLUS", ChartSeries::ErrorBarType::Plus).value("MINUS", ChartSeries::ErrorBarType::Minus);
    py::enum_<ChartSeries::ErrorValueType>(m, "ChartErrorValueType")
        .value("FIXED_VALUE", ChartSeries::ErrorValueType::FixedValue).value("PERCENTAGE", ChartSeries::ErrorValueType::Percentage)
        .value("STANDARD_DEVIATION", ChartSeries::ErrorValueType::StandardDeviation).value("STANDARD_ERROR", ChartSeries::ErrorValueType::StandardError)
        .value("CUSTOM", ChartSeries::ErrorValueType::Custom);
    py::class_<ChartSeries::Trendline>(m, "ChartTrendline")
        .def(py::init<>()).def_readwrite("type", &ChartSeries::Trendline::type).def_readwrite("order", &ChartSeries::Trendline::order)
        .def_readwrite("period", &ChartSeries::Trendline::period).def_readwrite("forward", &ChartSeries::Trendline::forward)
        .def_readwrite("backward", &ChartSeries::Trendline::backward).def_readwrite("display_equation", &ChartSeries::Trendline::displayEquation)
        .def_readwrite("display_r_squared", &ChartSeries::Trendline::displayRSquared).def_readwrite("line_format", &ChartSeries::Trendline::lineFormat);
    py::class_<ChartSeries::ErrorBars>(m, "ChartErrorBars")
        .def(py::init<>()).def_readwrite("direction", &ChartSeries::ErrorBars::direction).def_readwrite("bar_type", &ChartSeries::ErrorBars::barType)
        .def_readwrite("value_type", &ChartSeries::ErrorBars::valueType).def_readwrite("value", &ChartSeries::ErrorBars::value)
        .def_readwrite("no_end_cap", &ChartSeries::ErrorBars::noEndCap).def_readwrite("plus_reference", &ChartSeries::ErrorBars::plusReference)
        .def_readwrite("minus_reference", &ChartSeries::ErrorBars::minusReference).def_readwrite("line_format", &ChartSeries::ErrorBars::lineFormat);
    py::class_<ChartSeries>(m, "ChartSeries")
        .def(py::init<>())
        .def(py::init<std::string>())
        .def_property("title", &ChartSeries::title, &ChartSeries::setTitle)
        .def_property("values_reference", &ChartSeries::valuesReference, &ChartSeries::setValuesReference)
        .def_property("categories_reference", &ChartSeries::categoriesReference, &ChartSeries::setCategoriesReference)
        .def_property("bubble_size_reference", &ChartSeries::bubbleSizeReference, &ChartSeries::setBubbleSizeReference)
        .def_property("title_reference", &ChartSeries::titleReference, &ChartSeries::setTitleReference)
        .def_property("title_cache", &ChartSeries::titleCache, &ChartSeries::setTitleCache)
        .def_property("categories_cache", &ChartSeries::categoriesCache, &ChartSeries::setCategoriesCache)
        .def_property("values_cache", &ChartSeries::valuesCache, &ChartSeries::setValuesCache)
        .def_property("bubble_size_cache", &ChartSeries::bubbleSizeCache, &ChartSeries::setBubbleSizeCache)
        .def_property_readonly("has_smooth", &ChartSeries::hasSmooth)
        .def_property("smooth", &ChartSeries::smooth, &ChartSeries::setSmooth)
        .def("clear_smooth", &ChartSeries::clearSmooth)
        .def_property("trendlines", &ChartSeries::trendlines, &ChartSeries::setTrendlines)
        .def_property("error_bars", &ChartSeries::errorBars, &ChartSeries::setErrorBars)
        .def_property("data_labels", &ChartSeries::dataLabels, &ChartSeries::setDataLabels)
        .def_property("line_format", &ChartSeries::lineFormat, &ChartSeries::setLineFormat)
        .def_property("fill_format", &ChartSeries::fillFormat, &ChartSeries::setFillFormat)
        .def_property("marker_format", &ChartSeries::markerFormat, &ChartSeries::setMarkerFormat)
        .def_property("data_points", &ChartSeries::dataPoints, &ChartSeries::setDataPoints)
        .def("data_point", [](const ChartSeries& s, std::size_t index) -> py::object {
            const auto* point = s.dataPoint(index);
            return point ? py::cast(*point, py::return_value_policy::reference) : py::none{};
        })
        .def("reference", &ChartSeries::reference, py::arg("sheet_name"), py::arg("range_ref"))
        .def("categories", &ChartSeries::categories, py::arg("sheet_name"), py::arg("range_ref"));
    py::class_<ChartThemeFontScheme>(m, "ChartThemeFontScheme")
        .def(py::init<>()).def_readwrite("present", &ChartThemeFontScheme::present).def_readwrite("name", &ChartThemeFontScheme::name)
        .def_readwrite("major_latin_typeface", &ChartThemeFontScheme::majorLatinTypeface).def_readwrite("minor_latin_typeface", &ChartThemeFontScheme::minorLatinTypeface);
    py::class_<ChartThemeEffectScheme>(m, "ChartThemeEffectScheme")
        .def(py::init<>()).def_readwrite("present", &ChartThemeEffectScheme::present).def_readwrite("name", &ChartThemeEffectScheme::name)
        .def_readwrite("fill_style_count", &ChartThemeEffectScheme::fillStyleCount).def_readwrite("line_style_count", &ChartThemeEffectScheme::lineStyleCount)
        .def_readwrite("effect_style_count", &ChartThemeEffectScheme::effectStyleCount).def_readwrite("background_fill_style_count", &ChartThemeEffectScheme::backgroundFillStyleCount);
    py::class_<ChartThemeColor>(m, "ChartThemeColor")
        .def(py::init<>()).def_readwrite("name", &ChartThemeColor::name).def_readwrite("srgb", &ChartThemeColor::srgb);
    py::class_<ChartResolvedColor>(m, "ChartResolvedColor")
        .def(py::init<>()).def_readwrite("present", &ChartResolvedColor::present).def_readwrite("red", &ChartResolvedColor::red)
        .def_readwrite("green", &ChartResolvedColor::green).def_readwrite("blue", &ChartResolvedColor::blue)
        .def_readwrite("alpha", &ChartResolvedColor::alpha).def("srgb", &ChartResolvedColor::srgb);
    py::class_<ChartThemePalette>(m, "ChartThemePalette")
        .def(py::init<>()).def_readwrite("present", &ChartThemePalette::present).def_readwrite("colors", &ChartThemePalette::colors)
        .def_readwrite("font_scheme", &ChartThemePalette::fontScheme).def_readwrite("effect_scheme", &ChartThemePalette::effectScheme)
        .def("base_color", &ChartThemePalette::baseColor).def("resolve_base", &ChartThemePalette::resolveBase)
        .def("resolve", &ChartThemePalette::resolve).def("resolve_final_rgb", &ChartThemePalette::resolveFinalRgb);
    py::class_<ChartStyleResources>(m, "ChartStyleResources")
        .def(py::init<>()).def_readwrite("chart_style_present", &ChartStyleResources::chartStylePresent)
        .def_readwrite("color_style_present", &ChartStyleResources::colorStylePresent)
        .def_readwrite("chart_style_part", &ChartStyleResources::chartStylePart).def_readwrite("color_style_part", &ChartStyleResources::colorStylePart);
    py::class_<ChartLegendFormat>(m, "ChartLegendFormat")
        .def(py::init<>()).def_readwrite("present", &ChartLegendFormat::present).def_readwrite("overlay", &ChartLegendFormat::overlay)
        .def_readwrite("layout", &ChartLegendFormat::layout).def_readwrite("fill", &ChartLegendFormat::fill).def_readwrite("line", &ChartLegendFormat::line);
    py::class_<Chart>(m, "Chart")
        .def(py::init<Chart::Type>(), py::arg("type") = Chart::Type::Bar)
        .def_property_readonly("type", &Chart::type)
        .def_property("grouping", &Chart::grouping, &Chart::setGrouping)
        .def_property("title", &Chart::title, &Chart::setTitle)
        .def_property("x_axis_title", &Chart::xAxisTitle, &Chart::setXAxisTitle)
        .def_property("y_axis_title", &Chart::yAxisTitle, &Chart::setYAxisTitle)
        .def_property("style", &Chart::style, &Chart::setStyle)
        .def_property("title_rich_text", &Chart::titleRichText, &Chart::setTitleRichText)
        .def_property("theme_palette", &Chart::themePalette, &Chart::setThemePalette)
        .def_property("style_resources", &Chart::styleResources, &Chart::setStyleResources)
        .def("resolve_theme_base_color", &Chart::resolveThemeBaseColor)
        .def("resolve_theme_color", &Chart::resolveThemeColor)
        .def("resolve_theme_final_rgb", &Chart::resolveThemeFinalRgb)
        .def_property("width", &Chart::width, &Chart::setWidth)
        .def_property("height", &Chart::height, &Chart::setHeight)
        .def_property("show_legend", &Chart::showLegend, &Chart::setShowLegend)
        .def_property("legend_position", &Chart::legendPosition, &Chart::setLegendPosition)
        .def_property("legend_format", &Chart::legendFormat, &Chart::setLegendFormat)
        .def_property("plot_area_layout", &Chart::plotAreaLayout, &Chart::setPlotAreaLayout)
        .def_property("chart_area_fill_format", &Chart::chartAreaFillFormat, &Chart::setChartAreaFillFormat)
        .def_property("chart_area_line_format", &Chart::chartAreaLineFormat, &Chart::setChartAreaLineFormat)
        .def_property("plot_area_fill_format", &Chart::plotAreaFillFormat, &Chart::setPlotAreaFillFormat)
        .def_property("plot_area_line_format", &Chart::plotAreaLineFormat, &Chart::setPlotAreaLineFormat)
        .def_property("data_table", &Chart::dataTable, &Chart::setDataTable)
        .def_property("view_3d", &Chart::view3D, &Chart::setView3D)
        .def_property("floor_format", &Chart::floorFormat, &Chart::setFloorFormat)
        .def_property("side_wall_format", &Chart::sideWallFormat, &Chart::setSideWallFormat)
        .def_property("back_wall_format", &Chart::backWallFormat, &Chart::setBackWallFormat)
        .def_property_readonly("modern", &Chart::modern)
        .def_property_readonly("combined", &Chart::combined)
        .def("primary_plot", [](Chart& c) -> Chart::Plot& { return c.primaryPlot(); }, py::return_value_policy::reference_internal)
        .def("primary_plot_or_none", [](const Chart& c) -> py::object {
            const auto* plot = c.primaryPlotOrNull();
            return plot ? py::cast(*plot, py::return_value_policy::reference) : py::none{};
        })
        .def("primary_plot_or_null", [](const Chart& c) -> py::object {
            const auto* plot = c.primaryPlotOrNull();
            return plot ? py::cast(*plot, py::return_value_policy::reference) : py::none{};
        })
        .def("add_plot", &Chart::addPlot, py::arg("type"), py::arg("first_series"), py::arg("series_count"), py::arg("secondary_axes") = false,
             py::return_value_policy::reference_internal)
        .def("add_series", &Chart::addSeries, py::return_value_policy::reference_internal)
        .def_property_readonly("series", [](const Chart& c) { return stable_vector_copy(c.series()); })
        .def_property_readonly("plots", [](Chart& c) -> std::vector<Chart::Plot>& { return c.plots(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("axes", &Chart::axes)
        .def_property("anchor_info", &Chart::anchorInfo, &Chart::setAnchorInfo)
        .def_property_readonly("stable_id", &Chart::stableId)
        .def_property_readonly("source_drawing_part", &Chart::sourceDrawingPart)
        .def_property_readonly("source_chart_part", &Chart::sourceChartPart)
        .def_property_readonly("source_relationship_id", &Chart::sourceRelationshipId)
        .def_property_readonly("drawing_object_name", &Chart::drawingObjectName)
        .def_property_readonly("imported", &Chart::imported)
        .def("set_stable_id", &Chart::setStableId)
        .def("set_source_drawing_part", &Chart::setSourceDrawingPart)
        .def("set_source_chart_part", &Chart::setSourceChartPart)
        .def("set_source_relationship_id", &Chart::setSourceRelationshipId)
        .def("set_drawing_object_name", &Chart::setDrawingObjectName)
        .def("set_imported", &Chart::setImported)
        .def("set_axes", &Chart::setAxes)
        .def("set_plots", &Chart::setPlots)
        .def("set_primary_axis_ids", &Chart::setPrimaryAxisIds, py::arg("x_axis_id"), py::arg("y_axis_id"))
        .def_property_readonly("has_secondary_axes", &Chart::hasSecondaryAxes)
        .def_property_readonly("primary_x_axis_id", &Chart::primaryXAxisId)
        .def_property_readonly("primary_y_axis_id", &Chart::primaryYAxisId)
        .def("axis_by_id", [](const Chart& c, std::uint64_t id) -> py::object {
            const auto* axis = c.axisById(id);
            return axis ? py::cast(*axis, py::return_value_policy::reference) : py::none{};
        })
        .def_static("type_name", &Chart::typeName, py::arg("type"), py::arg("grouping") = Chart::Grouping::Standard)
        .def_static("is_modern_type", &Chart::isModernType, py::arg("type"));

    // === Pivot tables ===
    py::enum_<PivotGrouping::Kind>(m, "PivotGroupingKind")
        .value("NONE", PivotGrouping::Kind::None).value("NUMERIC", PivotGrouping::Kind::Numeric).value("DATE", PivotGrouping::Kind::Date);
    py::enum_<PivotGrouping::DatePart>(m, "PivotDatePart")
        .value("SECONDS", PivotGrouping::DatePart::Seconds).value("MINUTES", PivotGrouping::DatePart::Minutes)
        .value("HOURS", PivotGrouping::DatePart::Hours).value("DAYS", PivotGrouping::DatePart::Days)
        .value("MONTHS", PivotGrouping::DatePart::Months).value("QUARTERS", PivotGrouping::DatePart::Quarters)
        .value("YEARS", PivotGrouping::DatePart::Years);
    py::class_<PivotGrouping>(m, "PivotGrouping")
        .def(py::init<>())
        .def_readwrite("kind", &PivotGrouping::kind)
        .def_readwrite("auto_start", &PivotGrouping::autoStart)
        .def_readwrite("auto_end", &PivotGrouping::autoEnd)
        .def_readwrite("start", &PivotGrouping::start)
        .def_readwrite("end", &PivotGrouping::end)
        .def_readwrite("interval", &PivotGrouping::interval)
        .def_readwrite("date_part", &PivotGrouping::datePart)
        .def_readwrite("start_date", &PivotGrouping::startDate)
        .def_readwrite("end_date", &PivotGrouping::endDate)
        .def_property_readonly("active", &PivotGrouping::active);

    py::class_<PivotFilter>(m, "PivotFilter")
        .def(py::init<>())
        .def_readwrite("type", &PivotFilter::type)
        .def_readwrite("field_index", &PivotFilter::fieldIndex)
        .def_readwrite("measure_field_index", &PivotFilter::measureFieldIndex)
        .def_readwrite("value1", &PivotFilter::value1)
        .def_readwrite("value2", &PivotFilter::value2)
        .def_readwrite("top10_value", &PivotFilter::top10Value)
        .def_readwrite("top10_percent", &PivotFilter::top10Percent)
        .def_readwrite("top10_top", &PivotFilter::top10Top);

    py::class_<PivotCache>(m, "PivotCache")
        .def(py::init<>())
        .def_property("cache_id", &PivotCache::cacheId, &PivotCache::setCacheId)
        .def_property("source_data", &PivotCache::sourceData, &PivotCache::setSourceData)
        .def_property("refresh_on_load", &PivotCache::refreshOnLoad, &PivotCache::setRefreshOnLoad)
        .def_property("save_data", &PivotCache::saveData, &PivotCache::setSaveData)
        .def_property("enable_refresh", &PivotCache::enableRefresh, &PivotCache::setEnableRefresh)
        .def_property("missing_items_limit", &PivotCache::missingItemsLimit, &PivotCache::setMissingItemsLimit)
        .def_property("background_query", &PivotCache::backgroundQuery, &PivotCache::setBackgroundQuery)
        .def_property("optimize_memory", &PivotCache::optimizeMemory, &PivotCache::setOptimizeMemory)
        .def_property("upgrade_on_refresh", &PivotCache::upgradeOnRefresh, &PivotCache::setUpgradeOnRefresh)
        .def_property("support_subquery", &PivotCache::supportSubquery, &PivotCache::setSupportSubquery)
        .def_property("support_advanced_drill", &PivotCache::supportAdvancedDrill, &PivotCache::setSupportAdvancedDrill)
        .def_property("refreshed_by", &PivotCache::refreshedBy, &PivotCache::setRefreshedBy)
        .def_property("fields", [](PivotCache& c) -> std::vector<std::string>& { return c.fields(); }, &PivotCache::setFields, py::return_value_policy::reference_internal)
        .def_property("records", [](PivotCache& c) -> std::vector<std::vector<std::string>>& { return c.records(); }, &PivotCache::setRecords, py::return_value_policy::reference_internal)
        .def("add_field", &PivotCache::addField)
        .def("add_record", &PivotCache::addRecord)
        .def("clear_records", &PivotCache::clearRecords)
        .def("field_index", &PivotCache::fieldIndex);

    py::class_<PivotField>(m, "PivotField")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("name"))
        .def_property("name", &PivotField::name, &PivotField::setName)
        .def_property("axis", &PivotField::axis, &PivotField::setAxis)
        .def_property("show_all", &PivotField::showAll, &PivotField::setShowAll)
        .def_property("sort_type", &PivotField::sortType, &PivotField::setSortType)
        .def_property("subtotal_top", &PivotField::subtotalTop, &PivotField::setSubtotalTop)
        .def_property("insert_blank_row", &PivotField::insertBlankRow, &PivotField::setInsertBlankRow)
        .def_property("repeat_item_labels", &PivotField::repeatItemLabels, &PivotField::setRepeatItemLabels)
        .def_property("include_new_items_in_filter", &PivotField::includeNewItemsInFilter, &PivotField::setIncludeNewItemsInFilter)
        .def_property("multiple_item_selection_allowed", &PivotField::multipleItemSelectionAllowed, &PivotField::setMultipleItemSelectionAllowed)
        .def_property("selected_item_index", &PivotField::selectedItemIndex, &PivotField::setSelectedItemIndex)
        .def_property("field_index", &PivotField::fieldIndex, &PivotField::setFieldIndex)
        .def_property("compact", &PivotField::compact, &PivotField::setCompact)
        .def_property("outline", &PivotField::outline, &PivotField::setOutline)
        .def_property("insert_page_break", &PivotField::insertPageBreak, &PivotField::setInsertPageBreak)
        .def_property("show_drop_downs", &PivotField::showDropDowns, &PivotField::setShowDropDowns)
        .def_property("default_subtotal", &PivotField::defaultSubtotal, &PivotField::setDefaultSubtotal)
        .def_property("subtotals", &PivotField::subtotals, &PivotField::setSubtotals)
        .def_property("hidden_item_indexes", &PivotField::hiddenItemIndexes, &PivotField::setHiddenItemIndexes)
        .def_property("grouping", [](PivotField& f) -> PivotGrouping& { return f.grouping(); }, &PivotField::setGrouping, py::return_value_policy::reference_internal)
        .def("add_subtotal", &PivotField::addSubtotal)
        .def("item_hidden", &PivotField::itemHidden)
        .def("set_item_hidden", &PivotField::setItemHidden, py::arg("index"), py::arg("hidden") = true);

    py::class_<PivotFieldReference>(m, "PivotFieldReference")
        .def(py::init<>())
        .def_property("field_index", &PivotFieldReference::fieldIndex, &PivotFieldReference::setFieldIndex)
        .def_property("name", &PivotFieldReference::name, &PivotFieldReference::setName)
        .def_property("subtotal", &PivotFieldReference::subtotal, &PivotFieldReference::setSubtotal)
        .def_property("caption", &PivotFieldReference::caption, &PivotFieldReference::setCaption)
        .def_property("number_format_id", &PivotFieldReference::numberFormatId, &PivotFieldReference::setNumberFormatId)
        .def_property("show_data_as", &PivotFieldReference::showDataAs, &PivotFieldReference::setShowDataAs)
        .def_property("base_field", &PivotFieldReference::baseField, &PivotFieldReference::setBaseField)
        .def_property("base_item", &PivotFieldReference::baseItem, &PivotFieldReference::setBaseItem);

    py::enum_<PivotLayout>(m, "PivotLayout")
        .value("COMPACT", PivotLayout::Compact)
        .value("OUTLINE", PivotLayout::Outline)
        .value("TABULAR", PivotLayout::Tabular);

    py::class_<PivotTable>(m, "PivotTable")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("name"))
        .def_property("name", &PivotTable::name, &PivotTable::setName)
        .def_property("location", &PivotTable::location, &PivotTable::setLocation)
        .def_property("layout", &PivotTable::layout, &PivotTable::setLayout)
        .def_property("row_grand_totals", &PivotTable::rowGrandTotals, &PivotTable::setRowGrandTotals)
        .def_property("column_grand_totals", &PivotTable::columnGrandTotals, &PivotTable::setColumnGrandTotals)
        .def_property("preserve_formatting", &PivotTable::preserveFormatting, &PivotTable::setPreserveFormatting)
        .def_property("use_auto_formatting", &PivotTable::useAutoFormatting, &PivotTable::setUseAutoFormatting)
        .def_property("data_caption", &PivotTable::dataCaption, &PivotTable::setDataCaption)
        .def_property("style_name", &PivotTable::styleName, &PivotTable::setStyleName)
        .def_property("show_row_headers", &PivotTable::showRowHeaders, &PivotTable::setShowRowHeaders)
        .def_property("show_column_headers", &PivotTable::showColumnHeaders, &PivotTable::setShowColumnHeaders)
        .def_property("show_row_stripes", &PivotTable::showRowStripes, &PivotTable::setShowRowStripes)
        .def_property("show_column_stripes", &PivotTable::showColumnStripes, &PivotTable::setShowColumnStripes)
        .def_property("show_last_column", &PivotTable::showLastColumn, &PivotTable::setShowLastColumn)
        .def_property("show_empty_row", &PivotTable::showEmptyRow, &PivotTable::setShowEmptyRow)
        .def_property("show_empty_column", &PivotTable::showEmptyColumn, &PivotTable::setShowEmptyColumn)
        .def_property("show_drill", &PivotTable::showDrill, &PivotTable::setShowDrill)
        .def_property("enable_drill", &PivotTable::enableDrill, &PivotTable::setEnableDrill)
        .def_property("show_data_tips", &PivotTable::showDataTips, &PivotTable::setShowDataTips)
        .def_property("show_member_property_tips", &PivotTable::showMemberPropertyTips, &PivotTable::setShowMemberPropertyTips)
        .def_property("show_headers", &PivotTable::showHeaders, &PivotTable::setShowHeaders)
        .def_property("multiple_field_filters", &PivotTable::multipleFieldFilters, &PivotTable::setMultipleFieldFilters)
        .def_property("show_values_row", &PivotTable::showValuesRow, &PivotTable::setShowValuesRow)
        .def_property("subtotal_hidden_items", &PivotTable::subtotalHiddenItems, &PivotTable::setSubtotalHiddenItems)
        .def_property("page_wrap", &PivotTable::pageWrap, &PivotTable::setPageWrap)
        .def_property("page_over_then_down", &PivotTable::pageOverThenDown, &PivotTable::setPageOverThenDown)
        .def_property_readonly("cache", [](PivotTable& p) -> PivotCache& { return p.cache(); }, py::return_value_policy::reference_internal)
        .def("add_row_field", &PivotTable::addRowField, py::return_value_policy::reference_internal)
        .def("add_column_field", &PivotTable::addColumnField, py::return_value_policy::reference_internal)
        .def("add_page_field", &PivotTable::addPageField, py::return_value_policy::reference_internal)
        .def("add_data_field", [](PivotTable& p, const std::string& name, const std::string& subtotal) -> PivotFieldReference& { return p.addDataField(name, subtotal); },
             py::arg("name"), py::arg("subtotal") = "sum", py::return_value_policy::reference_internal)
        .def("add_data_field_by_index", [](PivotTable& p, int index) -> PivotFieldReference& { return p.addDataField(index); }, py::arg("field_index") = 0,
             py::return_value_policy::reference_internal)
        .def("add_filter", &PivotTable::addFilter, py::return_value_policy::reference_internal)
        .def_property_readonly("row_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.rowFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("column_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.columnFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("page_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.pageFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("data_fields", [](PivotTable& p) -> std::vector<PivotFieldReference>& { return p.dataFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("filters", [](PivotTable& p) -> std::vector<PivotFilter>& { return p.filters(); }, py::return_value_policy::reference_internal);

    // === Table ===
    py::class_<TableColumn>(m, "TableColumn")
        .def(py::init<>())
        .def(py::init<std::size_t, std::string>(), py::arg("id"), py::arg("name"))
        .def_property_readonly("id", &TableColumn::id)
        .def_property("name", &TableColumn::name, &TableColumn::setName);

    py::class_<TableStyleInfo>(m, "TableStyleInfo")
        .def(py::init<>())
        .def_property("name", &TableStyleInfo::name, &TableStyleInfo::setName)
        .def_property("show_first_column", &TableStyleInfo::showFirstColumn, &TableStyleInfo::setShowFirstColumn)
        .def_property("show_last_column", &TableStyleInfo::showLastColumn, &TableStyleInfo::setShowLastColumn)
        .def_property("show_row_stripes", &TableStyleInfo::showRowStripes, &TableStyleInfo::setShowRowStripes)
        .def_property("show_column_stripes", &TableStyleInfo::showColumnStripes, &TableStyleInfo::setShowColumnStripes);

    py::class_<Table>(m, "Table")
        .def(py::init<>())
        .def(py::init<std::string, std::string>())
        .def_property_readonly("name", &Table::name)
        .def_property("display_name", &Table::displayName, &Table::setDisplayName)
        .def_property("reference", &Table::reference, &Table::setReference)
        .def_property("show_header_row", &Table::showHeaderRow, &Table::setShowHeaderRow)
        .def_property("show_totals_row", &Table::showTotalsRow, &Table::setShowTotalsRow)
        .def_property_readonly("columns", [](const Table& t) { return stable_vector_copy(t.columns()); }, py::return_value_policy::reference_internal)
        .def_property_readonly("style_info", [](Table& t) -> TableStyleInfo& { return t.styleInfo(); }, py::return_value_policy::reference_internal)
        .def("add_column", &Table::addColumn, py::return_value_policy::reference_internal)
        .def("__repr__", [](const Table& t) { return "<Table " + t.name() + ">"; });

    // === NamedStyle ===
    py::class_<NamedStyle>(m, "NamedStyle")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("name"))
        .def(py::init<std::string, Style>(), py::arg("name"), py::arg("style"))
        .def_property("name", &NamedStyle::name, &NamedStyle::setName)
        .def("style", static_cast<Style& (NamedStyle::*)()>(&NamedStyle::style),
             py::return_value_policy::reference_internal);

    // === DefinedName ===
    py::class_<DefinedName>(m, "DefinedName")
        .def(py::init<>())
        .def(py::init<std::string, std::string>())
        .def_property_readonly("name", &DefinedName::name)
        .def_property("value", &DefinedName::value, &DefinedName::setValue)
        .def_property("local_sheet_id",
            [](const DefinedName& d) -> std::optional<std::size_t> { return d.localSheetId(); },
            [](DefinedName& d, std::optional<std::size_t> v) {
                if (v) d.setLocalSheetId(*v); else d.clearLocalSheetId();
            })
        .def("clear_local_sheet_id", &DefinedName::clearLocalSheetId)
        .def_property("hidden", &DefinedName::hidden, &DefinedName::setHidden)
        .def_property("comment", &DefinedName::comment, &DefinedName::setComment);

    // === Conditional formatting ===
    py::enum_<ConditionalRuleType>(m, "ConditionalRuleType")
        .value("FORMULA", ConditionalRuleType::Formula)
        .value("CELL_IS", ConditionalRuleType::CellIs)
        .value("DATA_BAR", ConditionalRuleType::DataBar)
        .value("COLOR_SCALE", ConditionalRuleType::ColorScale)
        .value("ICON_SET", ConditionalRuleType::IconSet);

    py::enum_<ConditionalOperator>(m, "ConditionalOperator")
        .value("EQUAL", ConditionalOperator::Equal)
        .value("NOT_EQUAL", ConditionalOperator::NotEqual)
        .value("LESS_THAN", ConditionalOperator::LessThan)
        .value("LESS_THAN_OR_EQUAL", ConditionalOperator::LessThanOrEqual)
        .value("GREATER_THAN", ConditionalOperator::GreaterThan)
        .value("GREATER_THAN_OR_EQUAL", ConditionalOperator::GreaterThanOrEqual)
        .value("BETWEEN", ConditionalOperator::Between)
        .value("NOT_BETWEEN", ConditionalOperator::NotBetween);

    py::enum_<IconSetStyle>(m, "IconSetStyle")
        .value("THREE_ARROWS", IconSetStyle::ThreeArrows)
        .value("THREE_ARROWS_GRAY", IconSetStyle::ThreeArrowsGray)
        .value("THREE_FLAGS", IconSetStyle::ThreeFlags)
        .value("THREE_TRAFFIC_LIGHTS", IconSetStyle::ThreeTrafficLights)
        .value("THREE_SIGNS", IconSetStyle::ThreeSigns)
        .value("THREE_SYMBOLS", IconSetStyle::ThreeSymbols)
        .value("THREE_STARS", IconSetStyle::ThreeStars)
        .value("THREE_TRIANGLES", IconSetStyle::ThreeTriangles)
        .value("FOUR_ARROWS", IconSetStyle::FourArrows)
        .value("FOUR_ARROWS_GRAY", IconSetStyle::FourArrowsGray)
        .value("FOUR_RED_TO_BLACK", IconSetStyle::FourRedToBlack)
        .value("FOUR_RATING", IconSetStyle::FourRating)
        .value("FOUR_TRAFFIC_LIGHTS", IconSetStyle::FourTrafficLights)
        .value("FIVE_ARROWS", IconSetStyle::FiveArrows)
        .value("FIVE_ARROWS_GRAY", IconSetStyle::FiveArrowsGray)
        .value("FIVE_RATING", IconSetStyle::FiveRating)
        .value("FIVE_QUARTERS", IconSetStyle::FiveQuarters)
        .value("FIVE_BOXES", IconSetStyle::FiveBoxes);

    py::class_<Cfvo>(m, "Cfvo")
        .def(py::init<>())
        .def(py::init<std::string, double>(), py::arg("type"), py::arg("value"))
        .def(py::init<std::string, std::string>(), py::arg("type"), py::arg("formula"))
        .def_readwrite("type", &Cfvo::type)
        .def_readwrite("value", &Cfvo::value)
        .def_readwrite("has_value", &Cfvo::hasValue)
        .def_readwrite("formula", &Cfvo::formula)
        .def_readwrite("color", &Cfvo::color);

    py::class_<DataBar>(m, "DataBar")
        .def(py::init<>())
        .def_readwrite("color", &DataBar::color)
        .def_readwrite("min", &DataBar::min)
        .def_readwrite("max", &DataBar::max)
        .def_readwrite("show_value", &DataBar::showValue)
        .def_readwrite("direction", &DataBar::direction)
        .def_readwrite("axis_position", &DataBar::axisPosition);

    py::class_<ColorScale>(m, "ColorScale")
        .def(py::init<>())
        .def("add_stop", &ColorScale::addStop)
        .def_readwrite("stops", &ColorScale::stops);

    py::class_<IconSet>(m, "IconSet")
        .def(py::init<>())
        .def_readwrite("icons", &IconSet::icons)
        .def("add_threshold", &IconSet::addThreshold)
        .def_readwrite("thresholds", &IconSet::thresholds)
        .def_readwrite("reverse", &IconSet::reverse)
        .def_readwrite("show_value", &IconSet::showValue)
        .def_readwrite("style", &IconSet::style);

    py::class_<ConditionalRule>(m, "ConditionalRule")
        .def(py::init<>())
        .def_static("formula", &ConditionalRule::formula, py::arg("expression"))
        .def_static("cell_is", &ConditionalRule::cellIs, py::arg("op"), py::arg("value"))
        .def_static("cell_is_between", &ConditionalRule::cellIsBetween,
                    py::arg("lower"), py::arg("upper"), py::arg("negate") = false)
        .def_static("data_bar", &ConditionalRule::dataBar, py::arg("color") = "FF638EC6")
        .def_static("color_scale", &ConditionalRule::colorScale)
        .def_static("icon_set", &ConditionalRule::iconSet, py::arg("icons") = "3Arrows")
        .def_property_readonly("type", &ConditionalRule::type)
        .def_property("op", &ConditionalRule::op, &ConditionalRule::setOperator)
        .def_property("formulas",
            [](const ConditionalRule& r) -> const std::vector<std::string>& { return r.formulas(); },
            [](ConditionalRule& r, const std::vector<std::string>& v) { r.setFormulas(v); })
        .def("add_formula", &ConditionalRule::addFormula)
        .def_property("priority", &ConditionalRule::priority, &ConditionalRule::setPriority)
        .def_property("stop_if_true", &ConditionalRule::stopIfTrue, &ConditionalRule::setStopIfTrue)
        .def_property_readonly("has_differential_style", &ConditionalRule::hasDifferentialStyle)
        .def("differential_style", static_cast<Style& (ConditionalRule::*)()>(&ConditionalRule::differentialStyle), py::return_value_policy::reference_internal)
        .def("set_differential_style", &ConditionalRule::setDifferentialStyle)
        .def("clear_differential_style", &ConditionalRule::clearDifferentialStyle)
        .def("get_data_bar", static_cast<DataBar& (ConditionalRule::*)()>(&ConditionalRule::getDataBar), py::return_value_policy::reference_internal)
        .def("get_color_scale", static_cast<ColorScale& (ConditionalRule::*)()>(&ConditionalRule::getColorScale), py::return_value_policy::reference_internal)
        .def("get_icon_set", static_cast<IconSet& (ConditionalRule::*)()>(&ConditionalRule::getIconSet), py::return_value_policy::reference_internal)
        .def("__repr__", [](const ConditionalRule& r) { return "<ConditionalRule>"; });

    py::class_<ConditionalFormattingEntry>(m, "ConditionalFormattingEntry")
        .def(py::init<std::string>(), py::arg("reference"))
        .def_property("reference", &ConditionalFormattingEntry::reference, &ConditionalFormattingEntry::setReference)
        .def("add_rule", &ConditionalFormattingEntry::addRule, py::return_value_policy::reference_internal)
        .def_property_readonly("rules", [](const ConditionalFormattingEntry& e) { return stable_vector_copy(e.rules()); }, py::return_value_policy::reference_internal)
        .def("empty", &ConditionalFormattingEntry::empty);

    py::class_<ConditionalFormattingCollection>(m, "ConditionalFormattingCollection")
        .def(py::init<>())
        .def("__call__", [](ConditionalFormattingCollection& c) -> ConditionalFormattingCollection& { return c; }, py::return_value_policy::reference_internal)
        .def("add", &ConditionalFormattingCollection::add, py::return_value_policy::reference_internal)
        .def("add_rule", [](ConditionalFormattingCollection& c, const std::string& ref,
                            ConditionalRule rule) -> ConditionalRule& {
                return c.addRule(ref, std::move(rule));
            }, py::arg("reference"), py::arg("rule"),
            py::return_value_policy::reference_internal)
        .def("add_rule", [](ConditionalFormattingCollection& c, const std::string& ref,
                            ConditionalRuleType type, const std::string& formula) -> ConditionalRule& {
                return c.addRule(ref, ConditionalRule::formula(formula));
            }, py::arg("reference"), py::arg("type"), py::arg("formula"),
            py::return_value_policy::reference_internal)
        .def_property_readonly("entries", [](const ConditionalFormattingCollection& c) { return stable_vector_copy(c.entries()); }, py::return_value_policy::reference_internal)
        .def("empty", &ConditionalFormattingCollection::empty)
        .def("clear", &ConditionalFormattingCollection::clear);

    // === Data validation ===
    py::enum_<DataValidationType>(m, "DataValidationType")
        .value("NONE", DataValidationType::None)
        .value("WHOLE", DataValidationType::Whole)
        .value("DECIMAL", DataValidationType::Decimal)
        .value("LIST", DataValidationType::List)
        .value("DATE", DataValidationType::Date)
        .value("TIME", DataValidationType::Time)
        .value("TEXT_LENGTH", DataValidationType::TextLength)
        .value("CUSTOM", DataValidationType::Custom);

    py::enum_<DataValidationOperator>(m, "DataValidationOperator")
        .value("BETWEEN", DataValidationOperator::Between)
        .value("NOT_BETWEEN", DataValidationOperator::NotBetween)
        .value("EQUAL", DataValidationOperator::Equal)
        .value("NOT_EQUAL", DataValidationOperator::NotEqual)
        .value("LESS_THAN", DataValidationOperator::LessThan)
        .value("LESS_THAN_OR_EQUAL", DataValidationOperator::LessThanOrEqual)
        .value("GREATER_THAN", DataValidationOperator::GreaterThan)
        .value("GREATER_THAN_OR_EQUAL", DataValidationOperator::GreaterThanOrEqual);

    py::enum_<DataValidationErrorStyle>(m, "DataValidationErrorStyle")
        .value("STOP", DataValidationErrorStyle::Stop)
        .value("WARNING", DataValidationErrorStyle::Warning)
        .value("INFORMATION", DataValidationErrorStyle::Information);

    py::class_<DataValidation>(m, "DataValidation")
        .def(py::init<>())
        .def(py::init<DataValidationType>(), py::arg("type"))
        .def_property("type", &DataValidation::type, &DataValidation::setType)
        .def_property("op", &DataValidation::op, &DataValidation::setOperator)
        .def_property("error_style", &DataValidation::errorStyle, &DataValidation::setErrorStyle)
        .def_property("formula1", &DataValidation::formula1, &DataValidation::setFormula1)
        .def_property("formula2", &DataValidation::formula2, &DataValidation::setFormula2)
        .def_property("reference", &DataValidation::reference, &DataValidation::setReference)
        .def_property("allow_blank", &DataValidation::allowBlank, &DataValidation::setAllowBlank)
        .def_property("show_drop_down", &DataValidation::showDropDown, &DataValidation::setShowDropDown)
        .def_property("show_input_message", &DataValidation::showInputMessage, &DataValidation::setShowInputMessage)
        .def_property("show_error_message", &DataValidation::showErrorMessage, &DataValidation::setShowErrorMessage)
        .def_property("prompt_title", &DataValidation::promptTitle, &DataValidation::setPromptTitle)
        .def_property("prompt", &DataValidation::prompt, &DataValidation::setPrompt)
        .def_property("error_title", &DataValidation::errorTitle, &DataValidation::setErrorTitle)
        .def_property("error", &DataValidation::error, &DataValidation::setError)
        .def_static("list", &DataValidation::list, py::arg("reference"), py::arg("formula"))
        .def("__repr__", [](const DataValidation& v) { return "<DataValidation>"; });

    py::class_<DataValidationCollection>(m, "DataValidationCollection")
        .def(py::init<>())
        .def("__call__", [](DataValidationCollection& c) -> DataValidationCollection& { return c; }, py::return_value_policy::reference_internal)
        .def("add", [](DataValidationCollection& c, DataValidationType type,
                       const std::string& reference) -> DataValidation& {
                return c.add(type, reference);
            }, py::arg("type"), py::arg("reference"),
            py::return_value_policy::reference_internal)
        .def("add_validation", [](DataValidationCollection& c, DataValidation v) -> DataValidation& {
                return c.add(std::move(v));
            }, py::arg("validation"), py::return_value_policy::reference_internal)
        .def_property_readonly("items", [](const DataValidationCollection& c) { return stable_vector_copy(c.items()); }, py::return_value_policy::reference_internal)
        .def("empty", &DataValidationCollection::empty)
        .def("clear", &DataValidationCollection::clear);

    // === Protection ===
    py::class_<WorksheetProtection>(m, "WorksheetProtection")
        .def(py::init<>())
        .def_property("enabled", &WorksheetProtection::enabled, &WorksheetProtection::setEnabled)
        .def_property("password_hash", &WorksheetProtection::passwordHash, &WorksheetProtection::setPasswordHash)
        .def("has_password", &WorksheetProtection::hasPassword)
        .def("set_password", &WorksheetProtection::setPassword)
        .def("clear_password", &WorksheetProtection::clearPassword)
        .def_property("select_locked_cells", &WorksheetProtection::selectLockedCells, &WorksheetProtection::setSelectLockedCells)
        .def_property("select_unlocked_cells", &WorksheetProtection::selectUnlockedCells, &WorksheetProtection::setSelectUnlockedCells)
        .def_property("format_cells", &WorksheetProtection::formatCells, &WorksheetProtection::setFormatCells)
        .def_property("format_columns", &WorksheetProtection::formatColumns, &WorksheetProtection::setFormatColumns)
        .def_property("format_rows", &WorksheetProtection::formatRows, &WorksheetProtection::setFormatRows)
        .def_property("insert_rows", &WorksheetProtection::insertRows, &WorksheetProtection::setInsertRows)
        .def_property("insert_columns", &WorksheetProtection::insertColumns, &WorksheetProtection::setInsertColumns)
        .def_property("delete_rows", &WorksheetProtection::deleteRows, &WorksheetProtection::setDeleteRows)
        .def_property("delete_columns", &WorksheetProtection::deleteColumns, &WorksheetProtection::setDeleteColumns)
        .def_property("sort", &WorksheetProtection::sort, &WorksheetProtection::setSort)
        .def_property("auto_filter", &WorksheetProtection::autoFilter, &WorksheetProtection::setAutoFilter);

    py::class_<WorkbookProtection>(m, "WorkbookProtection")
        .def(py::init<>())
        .def_property("lock_structure", &WorkbookProtection::lockStructure, &WorkbookProtection::setLockStructure)
        .def_property("lock_windows", &WorkbookProtection::lockWindows, &WorkbookProtection::setLockWindows)
        .def_property("lock_revision", &WorkbookProtection::lockRevision, &WorkbookProtection::setLockRevision)
        .def_property("workbook_password_hash", &WorkbookProtection::workbookPasswordHash, &WorkbookProtection::setWorkbookPasswordHash)
        .def("has_password", &WorkbookProtection::hasPassword)
        .def("set_password", &WorkbookProtection::setPassword)
        .def("clear_password", &WorkbookProtection::clearPassword);

    // === Dimensions ===
    py::class_<RowDimension>(m, "RowDimension")
        .def(py::init<>())
        .def_readwrite("height", &RowDimension::height)
        .def_readwrite("hidden", &RowDimension::hidden)
        .def_readwrite("outline_level", &RowDimension::outlineLevel)
        .def_readwrite("collapsed", &RowDimension::collapsed);

    py::class_<ColumnDimension>(m, "ColumnDimension")
        .def(py::init<>())
        .def_readwrite("width", &ColumnDimension::width)
        .def_readwrite("hidden", &ColumnDimension::hidden)
        .def_readwrite("best_fit", &ColumnDimension::bestFit)
        .def_readwrite("outline_level", &ColumnDimension::outlineLevel)
        .def_readwrite("collapsed", &ColumnDimension::collapsed);

    // === SheetView ===
    py::class_<SheetView>(m, "SheetView")
        .def(py::init<>())
        .def_property("workbook_view_id", &SheetView::workbookViewId, &SheetView::setWorkbookViewId)
        .def_property("tab_color",
            [](const SheetView& s) -> std::optional<std::string> { return s.tabColor(); },
            [](SheetView& s, std::optional<std::string> v) { if (v) s.setTabColor(*v); else s.clearTabColor(); })
        .def("clear_tab_color", &SheetView::clearTabColor)
        .def_property("zoom_scale", &SheetView::zoomScale, &SheetView::setZoomScale)
        .def_property("zoom_scale_normal", &SheetView::zoomScaleNormal, &SheetView::setZoomScaleNormal)
        .def_property("show_grid_lines", &SheetView::showGridLines, &SheetView::setShowGridLines)
        .def_property("tab_selected", &SheetView::tabSelected, &SheetView::setTabSelected)
        .def_property("right_to_left", &SheetView::rightToLeft, &SheetView::setRightToLeft)
        .def_property("show_outline_symbols", &SheetView::showOutlineSymbols, &SheetView::setShowOutlineSymbols)
        .def_property("pane", &SheetView::pane, &SheetView::setPane)
        .def_property("top_left_cell", &SheetView::topLeftCell, &SheetView::setTopLeftCell)
        .def_property("x_split", &SheetView::xSplit, &SheetView::setXSplit)
        .def_property("y_split", &SheetView::ySplit, &SheetView::setYSplit);

    // === CalcProperties / CustomProperties ===
    py::enum_<CalculationMode>(m, "CalculationMode")
        .value("AUTOMATIC", CalculationMode::Automatic)
        .value("AUTOMATIC_EXCEPT_DATA_TABLES", CalculationMode::AutomaticExceptDataTables)
        .value("MANUAL", CalculationMode::Manual);
    py::class_<CalcProperties>(m, "CalcProperties")
        .def(py::init<>())
        .def_property("calc_id", &CalcProperties::calcId, &CalcProperties::setCalcId)
        .def_property("calc_mode", &CalcProperties::calcMode, &CalcProperties::setCalcMode)
        .def_property("calculation_mode", &CalcProperties::calculationMode, &CalcProperties::setCalculationMode)
        .def_property("calc_on_save", &CalcProperties::calcOnSave, &CalcProperties::setCalcOnSave)
        .def_property("full_calc_on_load", &CalcProperties::fullCalcOnLoad, &CalcProperties::setFullCalcOnLoad)
        .def_property("full_precision", &CalcProperties::fullPrecision, &CalcProperties::setFullPrecision)
        .def_property("iterate", &CalcProperties::iterate, &CalcProperties::setIterate)
        .def_property("iterate_count", &CalcProperties::iterateCount, &CalcProperties::setIterateCount)
        .def_property("iterate_delta", &CalcProperties::iterateDelta, &CalcProperties::setIterateDelta);

    py::class_<CustomProperty>(m, "CustomProperty")
        .def(py::init<>())
        .def(py::init<std::string, std::string>(), py::arg("name"), py::arg("value"))
        .def(py::init<std::string, int>(), py::arg("name"), py::arg("value"))
        .def(py::init<std::string, double>(), py::arg("name"), py::arg("value"))
        .def(py::init<std::string, bool>(), py::arg("name"), py::arg("value"))
        .def_property("name", &CustomProperty::name, &CustomProperty::setName)
        .def_property("value", &CustomProperty::value, &CustomProperty::setValue)
        .def_property("type", &CustomProperty::type, &CustomProperty::setType);

    py::class_<CustomProperties>(m, "CustomProperties")
        .def(py::init<>())
        .def("add", &CustomProperties::add)
        .def_property_readonly("items", [](const CustomProperties& c) { return stable_vector_copy(c.items()); }, py::return_value_policy::reference_internal)
        .def("empty", &CustomProperties::empty);

    // === PreservedPart / LoadDiagnostics ===
    py::class_<PreservedPart>(m, "PreservedPart")
        .def(py::init<>())
        .def_readwrite("name", &PreservedPart::name)
        .def_readwrite("data", &PreservedPart::data)
        .def_readwrite("override_type", &PreservedPart::overrideType)
        .def_readwrite("extension", &PreservedPart::extension)
        .def_readwrite("default_type", &PreservedPart::defaultType)
        .def_readwrite("compress", &PreservedPart::compress);
    py::class_<PreservedRelationship>(m, "PreservedRelationship")
        .def(py::init<>())
        .def_readwrite("source_part", &PreservedRelationship::sourcePart)
        .def_readwrite("id", &PreservedRelationship::id)
        .def_readwrite("type", &PreservedRelationship::type)
        .def_readwrite("target", &PreservedRelationship::target)
        .def_readwrite("target_mode", &PreservedRelationship::targetMode);

    py::class_<LoadDiagnostics>(m, "LoadDiagnostics")
        .def(py::init<>())
        .def_readwrite("warnings", &LoadDiagnostics::warnings)
        .def_readwrite("errors", &LoadDiagnostics::errors)
        .def("had_errors", &LoadDiagnostics::hadErrors);

    // === CellRange ===
    py::class_<CellRange>(m, "CellRange")
        .def_property_readonly("min_row", &CellRange::minRow)
        .def_property_readonly("min_column", &CellRange::minColumn)
        .def_property_readonly("max_row", &CellRange::maxRow)
        .def_property_readonly("max_column", &CellRange::maxColumn)
        .def_property_readonly("row_count", &CellRange::rowCount)
        .def_property_readonly("column_count", &CellRange::columnCount)
        .def("address", &CellRange::address)
        .def("cell", &CellRange::cell, py::return_value_policy::reference_internal)
        .def("cells", [](CellRange& r) -> py::list {
            py::list out;
            for (auto* c : r.cells()) out.append(py::cast(c, py::return_value_policy::reference));
            return out;
        })
        .def("rows", [](CellRange& r) -> py::list {
            py::list out;
            for (auto& row : r.rows()) {
                py::list cells;
                for (auto* c : row) cells.append(py::cast(c, py::return_value_policy::reference));
                out.append(cells);
            }
            return out;
        })
        .def("set_value", [](CellRange& r, const py::object& v) { r.setValue(py_to_cellvalue(v)); })
        .def("clear", &CellRange::clear)
        .def("for_each", [](CellRange& r, const std::function<void(Cell&)>& cb) { r.forEach(cb); })
        .def("values", &CellRange::values)
        .def("formulas", &CellRange::formulas);

    // === Cell ===
    py::class_<Cell>(m, "Cell")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("address"))
        .def(py::init<std::size_t, std::size_t>(), py::arg("row"), py::arg("column"))
        .def_property_readonly("address", &Cell::address)
        .def_property_readonly("row", &Cell::row)
        .def_property_readonly("column", &Cell::column)
        .def_property("value",
            [](const Cell& c) { return cellvalue_to_py(c.value()); },
            [](Cell& c, const py::object& v) {
                auto cv = py_to_cellvalue(v);
                if (auto* dt = std::get_if<DateTime>(&cv)) {
                    if (dt->hour == 0 && dt->minute == 0 && dt->second == 0.0)
                        c.setDate(*dt);
                    else
                        c.setDateTime(*dt);
                } else {
                    c.setValue(cv);
                }
            })
        .def("set_string_value", [](Cell& c, std::string v) { c.setStringValue(std::move(v)); })
        .def("set_numeric_value", [](Cell& c, double v) { c.setNumericValue(v); })
        .def("set_bool_value", [](Cell& c, bool v) { c.setBoolValue(v); })
        .def("set_error", [](Cell& c, const py::object& e) { c.setError(py_to_cellerror(e)); })
        .def("set_date", [](Cell& c, const DateTime& d) { c.setDate(d); })
        .def("set_date", [](Cell& c, int year, int month, int day) { c.setDate(year, month, day); },
             py::arg("year"), py::arg("month"), py::arg("day"))
        .def("set_datetime", [](Cell& c, const DateTime& d) { c.setDateTime(d); })
        .def("date", [](const Cell& c) -> std::optional<DateTime> { return c.date(); })
        .def("numeric_value_or", &Cell::numericValueOr, py::arg("fallback") = 0.0)
        .def("string_value_or", &Cell::stringValueOr, py::arg("fallback") = std::string{})
        .def("is_error", &Cell::isError)
        .def("error", &Cell::error)
        .def("has_value", &Cell::hasValue)
        .def("set_formula", [](Cell& c, const std::string& f) { c.setFormula(f); })
        .def("set_shared_formula", &Cell::setSharedFormula, py::arg("formula"), py::arg("shared_index"), py::arg("reference") = std::string{})
        .def("set_array_formula", &Cell::setArrayFormula, py::arg("formula"), py::arg("reference"))
        .def("set_dynamic_array_formula", [](Cell& c, const std::string& f, const std::string& ref) {
            c.setDynamicArrayFormula(f, ref);
        })
        .def_property_readonly("formula", &Cell::formula)
        .def("has_formula", &Cell::hasFormula)
        .def_property("formula_metadata",
            [](Cell& c) -> FormulaMetadata& { return c.formulaMetadata(); },
            [](Cell& c, const FormulaMetadata& m) { c.formulaMetadata() = m; },
            py::return_value_policy::reference_internal)
        .def("clear_formula", &Cell::clearFormula)
        .def("clear", &Cell::clear)
        .def("empty", &Cell::empty)
        .def("is_numeric", &Cell::isNumeric)
        .def("is_string", &Cell::isString)
        .def("is_bool", &Cell::isBoolean)
        .def("is_date", &Cell::isDate)
        .def("value_type", &Cell::valueType)
        .def_property_readonly("has_rich_text", &Cell::hasRichText)
        .def("rich_text", &Cell::richText, py::return_value_policy::reference_internal)
        .def_property_readonly("rich_text_value", &Cell::richTextValue)
        .def("set_rich_text", &Cell::setRichText)
        .def("clear_rich_text", &Cell::clearRichText)
        .def("style", static_cast<Style& (Cell::*)()>(&Cell::style), py::return_value_policy::reference_internal)
        .def("font", static_cast<Font& (Cell::*)()>(&Cell::font), py::return_value_policy::reference_internal)
        .def("fill", static_cast<Fill& (Cell::*)()>(&Cell::fill), py::return_value_policy::reference_internal)
        .def("border", static_cast<Border& (Cell::*)()>(&Cell::border), py::return_value_policy::reference_internal)
        .def("alignment", static_cast<Alignment& (Cell::*)()>(&Cell::alignment), py::return_value_policy::reference_internal)
        .def("set_number_format", [](Cell& c, std::string v) { c.setNumberFormat(std::move(v)); })
        .def_property_readonly("number_format", &Cell::numberFormat)
        .def_property("style_index", &Cell::styleIndex, &Cell::setRawStyleIndex)
        .def("clear_raw_style_index", &Cell::clearRawStyleIndex)
        .def("has_hyperlink", &Cell::hasHyperlink)
        .def("hyperlink", &Cell::hyperlink, py::return_value_policy::reference_internal)
        .def("hyperlink_value", &Cell::hyperlinkValue)
        .def("set_hyperlink", &Cell::setHyperlink)
        .def("clear_hyperlink", &Cell::clearHyperlink)
        .def("has_comment", &Cell::hasComment)
        .def("comment", &Cell::comment, py::return_value_policy::reference_internal)
        .def("comment_value", &Cell::commentValue)
        .def("set_comment", [](Cell& c, const Comment& v) { c.setComment(v); })
        .def("clear_comment", &Cell::clearComment)
        .def("named_style", &Cell::namedStyle)
        .def("set_named_style", [](Cell& c, std::optional<std::string> name) { c.setNamedStyle(std::move(name)); })
        .def("offset", &Cell::offset)
        .def("__repr__", [](const Cell& c) { return "<Cell " + c.address() + ">"; });

    // === Worksheet ===
    py::class_<Worksheet>(m, "Worksheet")
        .def(py::init<>())
        .def(py::init<std::string>())
        .def_property("name", &Worksheet::name, &Worksheet::rename)
        .def_property("vba_code_name", &Worksheet::vbaCodeName, &Worksheet::setVbaCodeName)
        .def("cell",
            [](Worksheet& ws, const py::object& key) -> Cell& {
                if (py::isinstance<py::str>(key))
                    return ws.cell(key.cast<std::string>());
                if (py::isinstance<py::tuple>(key)) {
                    auto t = key.cast<py::tuple>();
                    return ws.cell(t[0].cast<std::size_t>(), t[1].cast<std::size_t>());
                }
                throw std::invalid_argument("key must be 'A1' or (row, col)");
            }, py::arg("key"), py::return_value_policy::reference_internal)
        .def("cell",
            [](Worksheet& ws, std::size_t row, std::size_t col) -> Cell& {
                return ws.cell(row, col);
            }, py::arg("row"), py::arg("col"), py::return_value_policy::reference_internal)
        .def("__getitem__",
            [](Worksheet& ws, const std::string& key) -> Cell& { return ws.cell(key); },
            py::return_value_policy::reference_internal)
        .def("try_cell", [](const Worksheet& ws, const std::string& addr) -> py::object {
            auto* c = ws.tryCell(addr);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def("try_cell", [](const Worksheet& ws, std::size_t row, std::size_t col) -> py::object {
            auto* c = ws.tryCell(row, col);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def("range", py::overload_cast<const std::string&>(&Worksheet::range))
        .def("range", py::overload_cast<std::size_t, std::size_t, std::size_t, std::size_t>(&Worksheet::range),
             py::arg("min_row"), py::arg("min_col"), py::arg("max_row"), py::arg("max_col"))
        .def("append", [](Worksheet& ws, const py::list& values) {
            std::vector<CellValue> cv;
            cv.reserve(values.size());
            for (auto item : values) cv.push_back(py_to_cellvalue(py::reinterpret_borrow<py::object>(item)));
            ws.append(cv);
        })
        .def("append_strings", [](Worksheet& ws, const py::list& values) {
            std::vector<CellValue> cv;
            cv.reserve(values.size());
            for (auto item : values) {
                if (!PyUnicode_Check(item.ptr())) cv.push_back(py_to_cellvalue(py::reinterpret_borrow<py::object>(item)));
                else cv.push_back(py::cast<std::string>(item));
            }
            ws.append(cv);
        })
        .def("write_array",
            [](Worksheet& ws, const py::array_t<double>& arr, std::size_t row0, std::size_t col0,
               bool transpose) {
                write_numpy_array(ws, arr, row0, col0, transpose);
            },
            py::arg("array"), py::arg("row") = 1, py::arg("col") = 1, py::arg("transpose") = false)
        .def("to_array",
            [](const Worksheet& ws, std::size_t minRow, std::size_t minCol,
               std::size_t maxRow, std::size_t maxCol) {
                return read_numpy_array(ws, minRow, minCol, maxRow, maxCol);
            },
            py::arg("min_row") = 0, py::arg("min_col") = 0,
            py::arg("max_row") = 0, py::arg("max_col") = 0)
        .def("from_records",
            [](Worksheet& ws, const py::list& rows, const py::object& columns, bool header) {
                const auto rowCount = rows.size();
                if (rowCount == 0) return;
                const auto& first = rows[0];
                const auto colCount = py::len(first);
                std::size_t r0 = 1;
                if (header && !columns.is_none()) {
                    for (std::size_t c = 0; c < colCount; ++c)
                        ws.cell(1, c + 1).setValue(py::cast<std::string>(columns.attr("__getitem__")(static_cast<py::ssize_t>(c))));
                    r0 = 2;
                }
                for (std::size_t r = 0; r < static_cast<std::size_t>(rowCount); ++r) {
                    const auto& row = rows[static_cast<py::ssize_t>(r)];
                    for (std::size_t c = 0; c < colCount; ++c)
                        ws.cell(r0 + r, c + 1).setValue(py_to_cellvalue(py::reinterpret_borrow<py::object>(row.attr("__getitem__")(static_cast<py::ssize_t>(c)))));
                }
            },
            py::arg("rows"), py::arg("columns") = py::none{}, py::arg("header") = true)
        .def("to_records",
            [](const Worksheet& ws, bool include_header) {
                const auto maxR = ws.maxRow();
                const auto maxC = ws.maxColumn();
                py::list result;
                if (include_header) {
                    py::list header;
                    for (std::size_t c = 1; c <= maxC; ++c) {
                        const auto* cell = ws.tryCell(1, c);
                        header.append(cell && cell->isString() ? cell->stringValueOr("") : "");
                    }
                    result.append(header);
                }
                for (std::size_t r = 2; r <= maxR; ++r) {
                    py::list row;
                    for (std::size_t c = 1; c <= maxC; ++c) {
                        const auto* cell = ws.tryCell(r, c);
                        row.append(cell ? cellvalue_to_py(cell->value()) : py::none{});
                    }
                    result.append(row);
                }
                return result;
            },
            py::arg("include_header") = true)
        .def("merge_cells", &Worksheet::mergeCells)
        .def("unmerge_cells", &Worksheet::unmergeCells)
        .def("is_merged", &Worksheet::isMerged)
        .def_property_readonly("merged_ranges", &Worksheet::mergedRanges,
                               py::return_value_policy::reference_internal)
        .def("freeze_panes", &Worksheet::freezePanes)
        .def("clear_freeze_panes", &Worksheet::clearFreezePanes)
        .def_property_readonly("frozen_pane", &Worksheet::frozenPane)
        .def("row_dimension", py::overload_cast<std::size_t>(&Worksheet::rowDimension), py::return_value_policy::reference_internal)
        .def("try_row_dimension", [](const Worksheet& ws, std::size_t row) -> py::object {
            const auto* d = ws.tryRowDimension(row);
            return d ? py::cast(*d, py::return_value_policy::reference) : py::none{};
        })
        .def("column_dimension", py::overload_cast<std::size_t>(&Worksheet::columnDimension), py::return_value_policy::reference_internal)
        .def("column_dimension", py::overload_cast<const std::string&>(&Worksheet::columnDimension), py::return_value_policy::reference_internal)
        .def("try_column_dimension", [](const Worksheet& ws, std::size_t col) -> py::object {
            const auto* d = ws.tryColumnDimension(col);
            return d ? py::cast(*d, py::return_value_policy::reference) : py::none{};
        })
        .def("set_print_area", &Worksheet::setPrintArea)
        .def_property_readonly("print_area", &Worksheet::printArea)
        .def_property("print_titles_rows", &Worksheet::printTitlesRows, &Worksheet::setPrintTitlesRows)
        .def_property("print_titles_cols", &Worksheet::printTitlesCols, &Worksheet::setPrintTitlesCols)
        .def_property_readonly("max_row", &Worksheet::maxRow)
        .def_property_readonly("max_column", &Worksheet::maxColumn)
        .def("dimensions", &Worksheet::dimensions)
        .def("extents", &Worksheet::extents)
        .def_property_readonly("empty", &Worksheet::empty)
        .def_property_readonly("row_count", &Worksheet::rowCount)
        .def_property_readonly("col_count", &Worksheet::columnCount)
        .def("insert_rows", &Worksheet::insertRows, py::arg("index"), py::arg("amount") = 1)
        .def("delete_rows", &Worksheet::deleteRows, py::arg("index"), py::arg("amount") = 1)
        .def("insert_columns", &Worksheet::insertColumns, py::arg("index"), py::arg("amount") = 1)
        .def("delete_columns", &Worksheet::deleteColumns, py::arg("index"), py::arg("amount") = 1)
        .def("apply_structural_edit", &Worksheet::applyStructuralEdit, py::arg("edit"))
        .def("auto_filter", py::overload_cast<>(&Worksheet::autoFilter),
             py::return_value_policy::reference_internal)
        .def_property_readonly("conditional_formatting", py::overload_cast<>(&Worksheet::conditionalFormatting),
             py::return_value_policy::reference_internal)
        .def_property_readonly("data_validations", py::overload_cast<>(&Worksheet::dataValidations),
             py::return_value_policy::reference_internal)
        .def("page_setup", py::overload_cast<>(&Worksheet::pageSetup),
             py::return_value_policy::reference_internal)
        .def("page_margins", py::overload_cast<>(&Worksheet::pageMargins),
             py::return_value_policy::reference_internal)
        .def("print_options", py::overload_cast<>(&Worksheet::printOptions),
             py::return_value_policy::reference_internal)
        .def("header_footer", py::overload_cast<>(&Worksheet::headerFooter),
             py::return_value_policy::reference_internal)
        .def("protection", py::overload_cast<>(&Worksheet::protection),
             py::return_value_policy::reference_internal)
        .def("add_table", &Worksheet::addTable, py::return_value_policy::reference_internal)
        .def("table", [](Worksheet& ws, const std::string& name) -> py::object {
            auto* t = ws.table(name);
            return t ? py::cast(*t, py::return_value_policy::reference) : py::none{};
        })
        .def_property_readonly("tables", [](const Worksheet& ws) { return stable_vector_copy(ws.tables()); }, py::return_value_policy::reference_internal)
        .def("add_image", [](Worksheet& ws, const std::string& path, const std::string& anchor) -> Image& {
            return ws.addImage(path, anchor);
        }, py::return_value_policy::reference_internal)
        .def("add_image", [](Worksheet& ws, const Image& image) -> Image& {
            return ws.addImage(image);
        }, py::return_value_policy::reference_internal)
        .def_property_readonly("image_count", [](const Worksheet& ws) { return ws.images().size(); })
        .def_property_readonly("images", [](const Worksheet& ws) { return stable_vector_copy(ws.images()); }, py::return_value_policy::reference_internal)
        .def("sheet_view", py::overload_cast<>(&Worksheet::sheetView),
             py::return_value_policy::reference_internal)
        .def("set_sheet_view", &Worksheet::setSheetView)
        .def("add_chart", &Worksheet::addChart)
        .def("chart", py::overload_cast<std::size_t>(&Worksheet::chart), py::return_value_policy::reference_internal)
        .def_property_readonly("chart_count", &Worksheet::chartCount)
        .def_property_readonly("charts", [](const Worksheet& ws) { return stable_vector_copy(ws.charts()); }, py::return_value_policy::reference_internal)
        .def("add_pivot_table", &Worksheet::addPivotTable)
        .def("add_loaded_pivot_table", &Worksheet::addLoadedPivotTable, py::return_value_policy::reference_internal)
        .def_property_readonly("loaded_pivot_count", &Worksheet::loadedPivotCount)
        .def_property_readonly("generated_pivot_start", &Worksheet::generatedPivotStart)
        .def_property_readonly("pivot_tables", [](const Worksheet& ws) { return stable_vector_copy(ws.pivotTables()); }, py::return_value_policy::reference_internal)
        .def("add_loaded_image", &Worksheet::addLoadedImage, py::return_value_policy::reference_internal)
        .def_property_readonly("loaded_image_count", &Worksheet::loadedImageCount)
        .def_property_readonly("appended_image_count", &Worksheet::appendedImageCount)
        .def("image_by_stable_id", [](const Worksheet& ws, const std::string& id) -> py::object {
            const auto* image = ws.imageByStableId(id);
            return image ? py::cast(*image, py::return_value_policy::reference) : py::none{};
        })
        .def("move_image", &Worksheet::moveImage)
        .def("move_image_absolute", &Worksheet::moveImageAbsolute)
        .def("resize_image", &Worksheet::resizeImage)
        .def("replace_image", py::overload_cast<const std::string&, Image>(&Worksheet::replaceImage))
        .def("replace_image", [](Worksheet& ws, const std::string& stableId, const std::string& path) {
            return ws.replaceImage(stableId, std::filesystem::path(path));
        }, py::arg("stable_id"), py::arg("path"))
        .def("remove_image", &Worksheet::removeImage)
        .def("add_loaded_chart", &Worksheet::addLoadedChart, py::return_value_policy::reference_internal)
        .def_property_readonly("loaded_chart_count", &Worksheet::loadedChartCount)
        .def_property_readonly("appended_chart_count", &Worksheet::appendedChartCount)
        .def("chart_by_stable_id", [](const Worksheet& ws, const std::string& id) -> py::object {
            const auto* chart = ws.chartByStableId(id);
            return chart ? py::cast(*chart, py::return_value_policy::reference) : py::none{};
        })
        .def("move_chart", &Worksheet::moveChart)
        .def("move_chart_absolute", &Worksheet::moveChartAbsolute)
        .def("resize_chart", &Worksheet::resizeChart)
        .def("set_chart_title", &Worksheet::setChartTitle)
        .def("set_chart_style", &Worksheet::setChartStyle)
        .def("set_chart_title_rich_text", &Worksheet::setChartTitleRichText)
        .def("set_chart_x_axis_title", &Worksheet::setChartXAxisTitle)
        .def("set_chart_y_axis_title", &Worksheet::setChartYAxisTitle)
        .def("set_chart_axis_title", &Worksheet::setChartAxisTitle)
        .def("set_chart_axis_title_rich_text", &Worksheet::setChartAxisTitleRichText)
        .def("set_chart_axis_number_format", &Worksheet::setChartAxisNumberFormat,
             py::arg("stable_id"), py::arg("axis_id"), py::arg("format_code"), py::arg("source_linked") = false)
        .def("set_chart_axis_ticks", &Worksheet::setChartAxisTicks)
        .def("set_chart_axis_units", &Worksheet::setChartAxisUnits,
             py::arg("stable_id"), py::arg("axis_id"), py::arg("major_unit"), py::arg("minor_unit") = 0.0)
        .def("set_chart_axis_crossing", &Worksheet::setChartAxisCrossing,
             py::arg("stable_id"), py::arg("axis_id"), py::arg("crosses"), py::arg("cross_between") = std::string{})
        .def("set_chart_axis_crosses_at", &Worksheet::setChartAxisCrossesAt)
        .def("clear_chart_axis_crosses_at", &Worksheet::clearChartAxisCrossesAt)
        .def("set_chart_axis_scaling", &Worksheet::setChartAxisScaling)
        .def("set_chart_axis_display_units", &Worksheet::setChartAxisDisplayUnits)
        .def("clear_chart_axis_display_units", &Worksheet::clearChartAxisDisplayUnits)
        .def("set_chart_axis_line_format", &Worksheet::setChartAxisLineFormat)
        .def("set_chart_axis_gridline_format", &Worksheet::setChartAxisGridlineFormat)
        .def("remove_chart_axis_gridlines", &Worksheet::removeChartAxisGridlines)
        .def("set_chart_area_line_format", &Worksheet::setChartAreaLineFormat)
        .def("set_chart_area_fill_format", &Worksheet::setChartAreaFillFormat)
        .def("set_chart_plot_area_line_format", &Worksheet::setChartPlotAreaLineFormat)
        .def("set_chart_plot_area_fill_format", &Worksheet::setChartPlotAreaFillFormat)
        .def("set_chart_plot_area_layout", &Worksheet::setChartPlotAreaLayout)
        .def("set_chart_view_3d", &Worksheet::setChartView3D)
        .def("set_chart_floor_format", &Worksheet::setChartFloorFormat)
        .def("set_chart_side_wall_format", &Worksheet::setChartSideWallFormat)
        .def("set_chart_back_wall_format", &Worksheet::setChartBackWallFormat)
        .def("set_chart_data_table", &Worksheet::setChartDataTable)
        .def("remove_chart_data_table", &Worksheet::removeChartDataTable)
        .def("set_chart_legend", &Worksheet::setChartLegend,
             py::arg("stable_id"), py::arg("show"), py::arg("position") = "r")
        .def("set_chart_legend_overlay", &Worksheet::setChartLegendOverlay)
        .def("set_chart_legend_layout", &Worksheet::setChartLegendLayout)
        .def("set_chart_legend_line_format", &Worksheet::setChartLegendLineFormat)
        .def("set_chart_legend_fill_format", &Worksheet::setChartLegendFillFormat)
        .def("set_chart_plot_drop_lines", &Worksheet::setChartPlotDropLines)
        .def("remove_chart_plot_drop_lines", &Worksheet::removeChartPlotDropLines)
        .def("set_chart_plot_high_low_lines", &Worksheet::setChartPlotHighLowLines)
        .def("remove_chart_plot_high_low_lines", &Worksheet::removeChartPlotHighLowLines)
        .def("remove_chart_plot_up_down_bars", &Worksheet::removeChartPlotUpDownBars)
        .def("set_chart_plot_up_down_bars", &Worksheet::setChartPlotUpDownBars)
        .def("set_chart_plot_first_slice_angle", &Worksheet::setChartPlotFirstSliceAngle)
        .def("set_chart_plot_doughnut_hole_size", &Worksheet::setChartPlotDoughnutHoleSize)
        .def("set_chart_plot_radar_style", &Worksheet::setChartPlotRadarStyle)
        .def("set_chart_plot_projected_pie_options", &Worksheet::setChartPlotProjectedPieOptions)
        .def("set_chart_plot_leader_line_format", &Worksheet::setChartPlotLeaderLineFormat)
        .def("remove_chart_plot_leader_lines", &Worksheet::removeChartPlotLeaderLines)
        .def("remove_chart_series_leader_lines", &Worksheet::removeChartSeriesLeaderLines)
        .def("set_chart_series_title", &Worksheet::setChartSeriesTitle)
        .def("set_chart_series_references", &Worksheet::setChartSeriesReferences)
        .def("set_chart_series_category_cache", &Worksheet::setChartSeriesCategoryCache)
        .def("set_chart_series_value_cache", &Worksheet::setChartSeriesValueCache)
        .def("set_chart_series_title_cache", &Worksheet::setChartSeriesTitleCache)
        .def("clear_chart_series_caches", &Worksheet::clearChartSeriesCaches)
        .def("remove_chart_plot_data_label_point", &Worksheet::removeChartPlotDataLabelPoint)
        .def("remove_chart_series_data_label_point", &Worksheet::removeChartSeriesDataLabelPoint)
        .def("set_chart_plot_data_labels", &Worksheet::setChartPlotDataLabels)
        .def("set_chart_series_data_labels", &Worksheet::setChartSeriesDataLabels)
        .def("set_chart_plot_data_label_point", &Worksheet::setChartPlotDataLabelPoint)
        .def("set_chart_series_data_label_point", &Worksheet::setChartSeriesDataLabelPoint)
        .def("set_chart_series_data_label_point_rich_text", &Worksheet::setChartSeriesDataLabelPointRichText)
        .def("remove_chart_series_data_point_format", &Worksheet::removeChartSeriesDataPointFormat)
        .def("set_chart_series_data_point_format", &Worksheet::setChartSeriesDataPointFormat)
        .def("set_chart_series_line_format", &Worksheet::setChartSeriesLineFormat)
        .def("set_chart_series_fill_format", &Worksheet::setChartSeriesFillFormat)
        .def("set_chart_series_marker_format", &Worksheet::setChartSeriesMarkerFormat)
        .def("set_chart_series_leader_line_format", &Worksheet::setChartSeriesLeaderLineFormat)
        .def("remove_chart_series_trendline", &Worksheet::removeChartSeriesTrendline)
        .def("set_chart_series_trendline_line_format", &Worksheet::setChartSeriesTrendlineLineFormat)
        .def("set_chart_series_trendline", &Worksheet::setChartSeriesTrendline)
        .def("add_chart_series_trendline", &Worksheet::addChartSeriesTrendline)
        .def("remove_chart_series_error_bars", &Worksheet::removeChartSeriesErrorBars)
        .def("set_chart_series_error_bars_line_format", &Worksheet::setChartSeriesErrorBarsLineFormat)
        .def("set_chart_series_error_bars", &Worksheet::setChartSeriesErrorBars)
        .def("remove_chart", &Worksheet::removeChart)
        .def("set_vba_code_name", &Worksheet::setVbaCodeName)
        .def_property_readonly("preserved_sheet_format_pr_xml", &Worksheet::preservedSheetFormatPrXml)
        .def("set_loaded_sheet_format_pr_xml", &Worksheet::setLoadedSheetFormatPrXml)
        .def_property_readonly("drawings_dirty", &Worksheet::drawingsDirty)
        .def_property_readonly("drawing_append_dirty", &Worksheet::drawingAppendDirty)
        .def_property_readonly("pivots_dirty", &Worksheet::pivotsDirty)
        .def("clear_dirty", &Worksheet::clearDirty)
        .def_property_readonly("row_dimensions", [](const Worksheet& ws) {
            py::dict result;
            for (const auto& [index, dimension] : ws.rowDimensions()) result[py::int_(index)] = py::cast(&dimension, py::return_value_policy::reference);
            return result;
        })
        .def_property_readonly("column_dimensions", [](const Worksheet& ws) {
            py::dict result;
            for (const auto& [index, dimension] : ws.columnDimensions()) result[py::int_(index)] = py::cast(&dimension, py::return_value_policy::reference);
            return result;
        })
        .def_property_readonly("cells", [](const Worksheet& ws) -> py::list {
            py::list result;
            for (const auto& [key, cell] : ws.cells())
                result.append(py::cast(&cell, py::return_value_policy::reference));
            return result;
        })
        .def_property_readonly("rows", [](Worksheet& ws) {
            py::list result;
            for (auto& row : ws.rows()) result.append(py::cast(row));
            return result;
        })
        .def("row", [](Worksheet& ws, std::size_t n) { return ws.row(n); })
        .def("iter_rows", [](const Worksheet& ws, std::size_t minRow, std::size_t maxRow, std::size_t minCol, std::size_t maxCol) {
            return ws.iterRows(minRow, maxRow, minCol, maxCol);
        }, py::arg("min_row") = 0, py::arg("max_row") = 0, py::arg("min_col") = 0, py::arg("max_col") = 0)
        .def("iter_cols", [](const Worksheet& ws, std::size_t minRow, std::size_t maxRow, std::size_t minCol, std::size_t maxCol) {
            return ws.iterCols(minRow, maxRow, minCol, maxCol);
        }, py::arg("min_row") = 0, py::arg("max_row") = 0, py::arg("min_col") = 0, py::arg("max_col") = 0)
        .def("dirty", &Worksheet::dirty)
        .def("mark_dirty", &Worksheet::markDirty)
        .def("__iter__", [](Worksheet& ws) {
            auto rows = ws.rows();
            py::list result;
            for (auto& row : rows) result.append(py::cast(row));
            return py::iter(result);
        })
        .def("__repr__", [](const Worksheet& ws) { return "<Worksheet '" + ws.name() + "'>"; });

    // === WorksheetExtents ===
    py::class_<WorksheetExtents>(m, "WorksheetExtents")
        .def(py::init<>())
        .def(py::init<std::size_t, std::size_t, std::size_t, std::size_t>(),
             py::arg("min_row"), py::arg("min_column"), py::arg("max_row"), py::arg("max_column"))
        .def_readwrite("min_row", &WorksheetExtents::minRow)
        .def_readwrite("min_column", &WorksheetExtents::minColumn)
        .def_readwrite("max_row", &WorksheetExtents::maxRow)
        .def_readwrite("max_column", &WorksheetExtents::maxColumn);

    // === Row ===
    py::class_<Row>(m, "Row")
        .def(py::init<Worksheet&, std::size_t>(), py::arg("worksheet"), py::arg("row_number"), py::keep_alive<1, 2>())
        .def_property_readonly("number", &Row::number)
        .def("cell", &Row::cell, py::return_value_policy::reference_internal)
        .def("try_cell", [](const Row& r, std::size_t col) -> py::object {
            const auto* c = r.tryCell(col);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def_property_readonly("cells", [](Row& r) -> py::list {
            py::list result;
            for (auto* cell : r.cells()) result.append(py::cast(cell, py::return_value_policy::reference));
            return result;
        })
        .def("__iter__", [](Row& r) {
            auto cells = r.cells();
            py::list result;
            for (auto* c : cells)
                result.append(py::cast(c, py::return_value_policy::reference));
            return py::iter(result);
        })
        .def("values", &Row::values);

    // === Streaming ===
    py::enum_<SharedStringMode>(m, "SharedStringMode")
        .value("DISABLED", SharedStringMode::Disabled)
        .value("HASH", SharedStringMode::Hash)
        .value("BOUNDED_LRU", SharedStringMode::BoundedLru);

    py::class_<StreamingCell>(m, "StreamingCell")
        .def(py::init<>())
        .def_readwrite("address", &StreamingCell::address)
        .def_property("value", [](const StreamingCell& c) { return cellvalue_to_py(c.value); },
            [](StreamingCell& c, const py::object& v) { c.value = py_to_cellvalue(v); })
        .def_readwrite("formula", &StreamingCell::formula)
        .def_readwrite("style_index", &StreamingCell::styleIndex);

    py::class_<StreamingRowIterator>(m, "StreamingRowIterator")
        .def(py::init<>())
        .def("row_number", &StreamingRowIterator::rowNumber);

    py::class_<StreamingReaderOptions>(m, "StreamingReaderOptions")
        .def(py::init<>())
        .def_readwrite("max_file_bytes", &StreamingReaderOptions::maxFileBytes)
        .def_readwrite("max_entry_bytes", &StreamingReaderOptions::maxEntryBytes)
        .def_readwrite("max_total_bytes", &StreamingReaderOptions::maxTotalBytes)
        .def_readwrite("max_entries", &StreamingReaderOptions::maxEntries)
        .def_readwrite("validate_cell_references", &StreamingReaderOptions::validateCellReferences);

    py::class_<StreamingWorksheetReader>(m, "StreamingWorksheetReader")
        .def(py::init<>())
        .def("for_each_row", &StreamingWorksheetReader::forEachRow,
             py::arg("callback"), "Callback receives (row_number, StreamingRow); return False to stop early.")
        .def("__iter__", [](StreamingWorksheetReader& r) {
            py::list result;
            r.forEachRow([&](std::size_t rowNumber, const StreamingRow& row) {
                py::list cells;
                for (const auto& c : row) cells.append(py::cast(c));
                result.append(py::make_tuple(rowNumber, cells));
                return true;
            });
            return py::iter(result);
        })
        .def("begin", &StreamingWorksheetReader::begin)
        .def("end", &StreamingWorksheetReader::end);

    py::class_<StreamingWorkbookReader>(m, "StreamingWorkbookReader")
        .def(py::init<std::filesystem::path>(), py::arg("path"))
        .def(py::init<std::filesystem::path, const StreamingReaderOptions&>(),
             py::arg("path"), py::arg("options"))
        .def("worksheet_names", &StreamingWorkbookReader::worksheetNames)
        .def("worksheet", [](StreamingWorkbookReader& r, const std::string& name) {
            return r.worksheet(name);
        }, py::arg("worksheet_name"), py::return_value_policy::move)
        .def("for_each_row", &StreamingWorkbookReader::forEachRow, py::arg("worksheet_name"), py::arg("callback"));

    py::class_<StreamingWorksheetWriter>(m, "StreamingWorksheetWriter")
        .def(py::init<>())
        .def("append", [](StreamingWorksheetWriter& w, const py::list& values) {
            std::vector<CellValue> cv;
            cv.reserve(values.size());
            for (auto item : values) cv.push_back(py_to_cellvalue(py::reinterpret_borrow<py::object>(item)));
            w.append(cv);
        })
        .def_property_readonly("row_count", &StreamingWorksheetWriter::rowCount)
        .def_property_readonly("name", &StreamingWorksheetWriter::name);

    py::class_<StreamingWorkbookWriter>(m, "StreamingWorkbookWriter")
        .def(py::init<std::filesystem::path, SharedStringMode, std::size_t>(),
             py::arg("output_path"), py::arg("shared_strings") = SharedStringMode::Disabled,
             py::arg("lru_capacity") = 1024)
        .def("add_worksheet", &StreamingWorkbookWriter::addWorksheet, py::return_value_policy::reference_internal)
        .def_property_readonly("sheet_count", &StreamingWorkbookWriter::sheetCount)
        .def("worksheet", py::overload_cast<std::size_t>(&StreamingWorkbookWriter::worksheet), py::return_value_policy::reference_internal)
        .def("close", &StreamingWorkbookWriter::close)
        .def_property_readonly("closed", &StreamingWorkbookWriter::closed)
        .def_property("date_1904", &StreamingWorkbookWriter::date1904, &StreamingWorkbookWriter::setDate1904)
        .def("set_compression_level", &StreamingWorkbookWriter::setCompressionLevel)
        .def("set_compression_strategy", &StreamingWorkbookWriter::setCompressionStrategy)
        .def("set_parallel_workers", &StreamingWorkbookWriter::setParallelWorkers);

    // === Preservation-first external data / Data Model inspection ===
    py::class_<ExternalWorkbookLinkInfo>(m, "ExternalWorkbookLinkInfo")
        .def(py::init<>())
        .def_readwrite("part_name", &ExternalWorkbookLinkInfo::partName).def_readwrite("sheet_names", &ExternalWorkbookLinkInfo::sheetNames).def_readwrite("defined_names", &ExternalWorkbookLinkInfo::definedNames);
    py::class_<WorkbookConnectionInfo>(m, "WorkbookConnectionInfo")
        .def(py::init<>())
        .def_readwrite("part_name", &WorkbookConnectionInfo::partName).def_readwrite("id", &WorkbookConnectionInfo::id).def_readwrite("name", &WorkbookConnectionInfo::name)
        .def_readwrite("description", &WorkbookConnectionInfo::description).def_readwrite("type", &WorkbookConnectionInfo::type)
        .def_readwrite("refresh_on_load", &WorkbookConnectionInfo::refreshOnLoad).def_readwrite("background", &WorkbookConnectionInfo::background).def_readwrite("deleted", &WorkbookConnectionInfo::deleted);
    py::class_<QueryTableInfo>(m, "QueryTableInfo")
        .def(py::init<>())
        .def_readwrite("part_name", &QueryTableInfo::partName).def_readwrite("name", &QueryTableInfo::name).def_readwrite("connection_id", &QueryTableInfo::connectionId).def_readwrite("refresh_on_load", &QueryTableInfo::refreshOnLoad);
    py::class_<ExternalDataInspection>(m, "ExternalDataInspection")
        .def(py::init<>())
        .def_readwrite("external_workbooks", &ExternalDataInspection::externalWorkbooks).def_readwrite("connections", &ExternalDataInspection::connections)
        .def_readwrite("query_tables", &ExternalDataInspection::queryTables).def_readwrite("power_query_parts", &ExternalDataInspection::powerQueryParts)
        .def_readwrite("web_query_parts", &ExternalDataInspection::webQueryParts)
        .def_readwrite("unknown_connection_parts", &ExternalDataInspection::unknownConnectionParts)
        .def_property_readonly("has_external_workbooks", &ExternalDataInspection::hasExternalWorkbooks)
        .def_property_readonly("has_connections", &ExternalDataInspection::hasConnections)
        .def_property_readonly("has_query_tables", &ExternalDataInspection::hasQueryTables)
        .def_property_readonly("has_power_query", &ExternalDataInspection::hasPowerQuery);
    py::class_<DataModelInspection>(m, "DataModelInspection")
        .def(py::init<>())
        .def_readwrite("present", &DataModelInspection::present).def_readwrite("has_olap_pivot_caches", &DataModelInspection::hasOlapPivotCaches)
        .def_readwrite("model_parts", &DataModelInspection::modelParts).def_readwrite("model_relationships", &DataModelInspection::modelRelationships)
        .def_readwrite("olap_pivot_cache_parts", &DataModelInspection::olapPivotCacheParts).def_readwrite("warnings", &DataModelInspection::warnings);

    // === Workbook ===
    py::class_<Workbook>(m, "Workbook")
        .def(py::init<>())
        .def_static("open", [](const std::string& path, const LoadOptions* opts) { Workbook wb; if (opts) wb.load(std::filesystem::path(path), *opts); else wb.load(std::filesystem::path(path)); return wb; }, py::arg("path"), py::arg("options") = nullptr)
        .def("__enter__", [](Workbook& wb) -> Workbook& { return wb; }, py::return_value_policy::reference_internal)
        .def("__exit__", [](Workbook&, py::object, py::object, py::object) { return false; })
        .def("add_worksheet", &Workbook::addWorksheet, py::return_value_policy::reference_internal)
        .def("copy_worksheet", &Workbook::copyWorksheet, py::return_value_policy::reference_internal)
        .def("remove_worksheet", &Workbook::removeWorksheet)
        .def("worksheet",
            [](Workbook& wb, const std::string& name) -> Worksheet* { return wb.worksheet(name); },
            py::return_value_policy::reference_internal)
        .def("get_worksheet",
            [](Workbook& wb, const std::string& name) -> Worksheet* { return wb.worksheet(name); },
            py::return_value_policy::reference_internal)
        .def("__getitem__",
            [](Workbook& wb, const py::object& key) -> Worksheet& {
                if (py::isinstance<py::str>(key)) {
                    auto* ws = wb.worksheet(key.cast<std::string>());
                    if (!ws) throw std::invalid_argument("Worksheet not found");
                    return *ws;
                }
                if (py::isinstance<py::int_>(key))
                    return wb[key.cast<std::size_t>()];
                throw std::invalid_argument("key must be str or int");
            }, py::return_value_policy::reference_internal)
        .def("__len__", &Workbook::sheetCount)
        .def_property_readonly("sheet_count", &Workbook::sheetCount)
        .def_property_readonly("sheet_names", &Workbook::sheetNames)
        .def_property_readonly("active", [](Workbook& wb) -> Worksheet& { return wb[0]; },
                               py::return_value_policy::reference_internal)
        .def("index", &Workbook::index)
        .def("load",
            [](Workbook& wb, const std::string& path) { wb.load(std::filesystem::path(path)); })
        .def("load",
            [](Workbook& wb, const std::string& path, const LoadOptions& opts) {
                wb.load(std::filesystem::path(path), opts);
            }, py::arg("path"), py::arg("options"))
        .def("load_bytes",
            [](Workbook& wb, const py::bytes& data, const LoadOptions* opts) {
                std::string raw = data.cast<std::string>();
                std::istringstream stream(raw);
                if (opts) wb.load(stream, *opts);
                else wb.load(stream);
            }, py::arg("data"), py::arg("options") = nullptr)
        .def("save",
            [](const Workbook& wb, const std::string& path, const SaveOptions* opts) {
                if (opts) wb.save(std::filesystem::path(path), *opts);
                else wb.save(std::filesystem::path(path));
            }, py::arg("path"), py::arg("options") = nullptr)
        .def("save_bytes",
            [](const Workbook& wb, const SaveOptions* opts) -> py::bytes {
                std::ostringstream stream;
                if (opts) wb.save(stream, *opts);
                else wb.save(stream);
                return py::bytes(stream.str());
            }, py::arg("options") = nullptr)
        .def("load_in_place", [](Workbook& wb, const std::string& path, const LoadOptions& opts) {
            wb.loadInPlace(std::filesystem::path(path), opts);
        }, py::arg("path"), py::arg("options"))
        .def("save_in_place", [](const Workbook& wb, const std::string& path, const SaveOptions& opts) {
            wb.saveInPlace(std::filesystem::path(path), opts);
        }, py::arg("path"), py::arg("options"))
        .def_property_readonly("properties", py::overload_cast<>(&Workbook::properties))
        .def("protection", py::overload_cast<>(&Workbook::protection),
             py::return_value_policy::reference_internal)
        .def("calc_properties", py::overload_cast<>(&Workbook::calcProperties),
             py::return_value_policy::reference_internal)
        .def("custom_properties", py::overload_cast<>(&Workbook::customProperties),
             py::return_value_policy::reference_internal)
        .def("add_named_style", &Workbook::addNamedStyle, py::return_value_policy::reference_internal)
        .def("named_style", py::overload_cast<const std::string&>(&Workbook::namedStyle, py::const_),
             py::return_value_policy::reference_internal)
        .def_property_readonly("named_styles", [](Workbook& wb) -> py::list {
            py::list out;
            for (auto& n : wb.namedStyles())
                out.append(py::cast(&n, py::return_value_policy::reference));
            return out;
        })
        .def("apply_named_style", &Workbook::applyNamedStyle)
        .def("add_defined_name", &Workbook::addDefinedName, py::return_value_policy::reference_internal)
        .def("defined_name", [](Workbook& wb, const std::string& name, std::optional<std::size_t> local_sheet_id) -> DefinedName* {
            return wb.definedName(name, local_sheet_id);
        }, py::arg("name"), py::arg("local_sheet_id") = std::nullopt,
           py::return_value_policy::reference_internal)
        .def_property_readonly("defined_names", [](Workbook& wb) -> py::list {
            py::list out;
            for (auto& n : wb.definedNames())
                out.append(py::cast(&n, py::return_value_policy::reference));
            return out;
        })
        .def("calculate_formulas", &Workbook::calculateFormulas, py::arg("options") = CalculationOptions{})
        .def("dependency_graph", &Workbook::dependencyGraph)
        .def("rename_worksheet", &Workbook::renameWorksheet, py::arg("old_name"), py::arg("new_name"), py::arg("options") = WorksheetRenameOptions{})
        .def("synchronize_chart_caches", &Workbook::synchronizeChartCaches, py::arg("options") = ChartCacheSyncOptions{})
        .def("reset_chart_cache_dependency_tracking", &Workbook::resetChartCacheDependencyTracking)
        .def_property_readonly("tracked_chart_cache_dependency_count", &Workbook::trackedChartCacheDependencyCount)
        .def("validate", &Workbook::validate, py::arg("options") = WorkbookValidationOptions{})
        .def("inspect_external_data", &Workbook::inspectExternalData)
        .def("inspect_data_model", &Workbook::inspectDataModel)
        .def("add_vba_project", [](Workbook& wb, const std::string& path) { wb.addVbaProject(std::filesystem::path(path)); }, py::arg("path"))
        .def("set_vba_project", [](Workbook& wb, const py::bytes& data) {
            const std::string raw = data.cast<std::string>();
            wb.setVbaProject(std::vector<unsigned char>(raw.begin(), raw.end()));
        }, py::arg("data"))
        .def_property_readonly("has_vba_project", &Workbook::hasVbaProject)
        .def("remove_vba_project", &Workbook::removeVbaProject)
        .def("set_vba_module", &Workbook::setVbaModule, py::arg("module"))
        .def("set_vba_module_text", &Workbook::setVbaModuleText, py::arg("module_name"), py::arg("source"))
        .def("set_vba_class_module_text", &Workbook::setVbaClassModuleText, py::arg("module_name"), py::arg("source"), py::arg("read_only") = false, py::arg("private_module") = false)
        .def("set_vba_document_module_text", &Workbook::setVbaDocumentModuleText, py::arg("module_name"), py::arg("source"))
        .def("vba_module_text", &Workbook::vbaModuleText, py::arg("module_name"))
        .def_property_readonly("vba_modules", &Workbook::vbaModules)
        .def_property_readonly("vba_project_bytes", [](const Workbook& wb) {
            const auto bytes = wb.vbaProjectBytes();
            return py::bytes(reinterpret_cast<const char*>(bytes.data()), bytes.size());
        })
        .def("save_vba_project", [](const Workbook& wb, const std::string& path) { wb.saveVbaProject(std::filesystem::path(path)); }, py::arg("path"))
        .def_property_readonly("has_vba_signature", &Workbook::hasVbaSignature)
        .def_property_readonly("vba_source_editable", &Workbook::vbaSourceEditable)
        .def_property("vba_project_properties", &Workbook::vbaProjectProperties, &Workbook::setVbaProjectProperties)
        .def("set_vba_project_properties", &Workbook::setVbaProjectProperties, py::arg("properties"))
        .def("remove_vba_module", &Workbook::removeVbaModule, py::arg("module_name"))
        .def("apply_structural_edit", &Workbook::applyStructuralEdit, py::arg("edit"), py::arg("options") = StructuralEditOptions{})
        .def("insert_rows", &Workbook::insertRows, py::arg("sheet_name"), py::arg("index"), py::arg("amount") = 1, py::arg("options") = StructuralEditOptions{})
        .def("delete_rows", &Workbook::deleteRows, py::arg("sheet_name"), py::arg("index"), py::arg("amount") = 1, py::arg("options") = StructuralEditOptions{})
        .def("insert_columns", &Workbook::insertColumns, py::arg("sheet_name"), py::arg("index"), py::arg("amount") = 1, py::arg("options") = StructuralEditOptions{})
        .def("delete_columns", &Workbook::deleteColumns, py::arg("sheet_name"), py::arg("index"), py::arg("amount") = 1, py::arg("options") = StructuralEditOptions{})
        .def_property("date_1904", &Workbook::date1904, &Workbook::setDate1904)
        .def_property("date1904", &Workbook::date1904, &Workbook::setDate1904)
        .def("clear", &Workbook::clear)
        .def_property_readonly("preserved_parts", [](Workbook& wb) -> std::vector<PreservedPart>& { return wb.preservedParts(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("preserved_relationships", [](const Workbook& wb) { return wb.preservedRelationships(); })
        .def_property_readonly("worksheets", [](Workbook& wb) -> py::list {
            py::list result;
            for (auto& ws : wb.worksheets())
                result.append(py::cast(&ws, py::return_value_policy::reference));
            return result;
        })
        .def_property_readonly("strict_namespaces", &Workbook::strictNamespaces)
        .def_property_readonly("diagnostics", [](const Workbook& wb) -> const LoadDiagnostics& { return wb.diagnostics(); }, py::return_value_policy::reference_internal)
        .def("__iter__", [](Workbook& wb) {
            py::list result;
            for (auto& ws : wb.worksheets())
                result.append(py::cast(&ws, py::return_value_policy::reference));
            return py::iter(result);
        })
        .def("__repr__", [](const Workbook& wb) {
            return "<Workbook sheets=" + std::to_string(wb.sheetCount()) + ">";
        });

    m.attr("__version__") = "1.12.0";

    // Excel 365 dynamic-array function prefix helper
    m.def("xlfn", [](const std::string& f) { return xlpp::xlfn(f); },
          py::arg("function"),
          "Prefix an Excel 365 function name with _xlfn. (e.g. SORT -> _xlfn.SORT)");
    m.def("looks_like_encrypted_office_file", [](const std::string& path) {
        return xlpp::looksLikeEncryptedOfficeFile(std::filesystem::path(path));
    }, py::arg("path"));
    m.def("inspect_office_encryption", [](const std::string& path) {
        return xlpp::inspectOfficeEncryption(std::filesystem::path(path));
    }, py::arg("path"));
    m.def("legacy_protection_password_hash", &xlpp::legacyProtectionPasswordHash,
          py::arg("password"));
    m.def("translate_formula_references", &xlpp::translateFormulaReferences,
          py::arg("formula"), py::arg("context_sheet"), py::arg("edit"));
    m.def("translate_range_references", &xlpp::translateRangeReferences,
          py::arg("reference"), py::arg("context_sheet"), py::arg("edit"));
    m.def("rename_worksheet_references", &xlpp::renameWorksheetReferences,
          py::arg("expression"), py::arg("old_worksheet_name"), py::arg("new_worksheet_name"));
    m.def("invalidate_worksheet_references", &xlpp::invalidateWorksheetReferences,
          py::arg("expression"), py::arg("removed_worksheet_name"));
}
