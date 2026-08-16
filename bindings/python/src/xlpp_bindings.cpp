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
    // Fast path using the CPython C API directly (no pybind type dispatch or
    // per-value refcount churn). This is the hot path for append() over large
    // grids.
    PyObject* const o = obj.ptr();
    if (PyUnicode_Check(o)) {
        Py_ssize_t length = 0;
        const char* utf8 = PyUnicode_AsUTF8AndSize(o, &length);
        if (!utf8) throw py::error_already_set();
        return std::string(utf8, static_cast<std::size_t>(length));
    }
    if (o == Py_None) return std::monostate{};
    if (PyBool_Check(o)) return o == Py_True;
    if (PyLong_Check(o)) {
        const auto value = PyLong_AsLongLong(o);
        if (value == -1 && PyErr_Occurred()) throw py::error_already_set();
        return static_cast<double>(value);
    }
    if (PyFloat_Check(o)) {
        const auto value = PyFloat_AsDouble(o);
        if (value == -1.0 && PyErr_Occurred()) throw py::error_already_set();
        return value;
    }
    if (PyDateTime_Check(o) || PyDate_Check(o)) return py_to_datetime(obj);
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

    py::enum_<PackageEncryptionMode>(m, "PackageEncryptionMode")
        .value("AGILE", PackageEncryptionMode::Agile)
        .value("STANDARD", PackageEncryptionMode::Standard);

    py::enum_<PackageEncryptionHash>(m, "PackageEncryptionHash")
        .value("SHA1", PackageEncryptionHash::Sha1)
        .value("SHA256", PackageEncryptionHash::Sha256)
        .value("SHA384", PackageEncryptionHash::Sha384)
        .value("SHA512", PackageEncryptionHash::Sha512);

    py::enum_<PackageEncryptionFormat>(m, "PackageEncryptionFormat")
        .value("NONE", PackageEncryptionFormat::None)
        .value("AGILE", PackageEncryptionFormat::Agile)
        .value("STANDARD", PackageEncryptionFormat::Standard)
        .value("UNSUPPORTED", PackageEncryptionFormat::Unsupported);

    py::class_<PackageEncryptionCertificateKeyInfo>(m, "PackageEncryptionCertificateKeyInfo")
        .def_readonly("x509_certificate", &PackageEncryptionCertificateKeyInfo::x509Certificate)
        .def_readonly("encrypted_key_bytes", &PackageEncryptionCertificateKeyInfo::encryptedKeyBytes)
        .def_readonly("cert_verifier_bytes", &PackageEncryptionCertificateKeyInfo::certVerifierBytes)
        .def_readonly("valid_encoding", &PackageEncryptionCertificateKeyInfo::validEncoding);

    py::class_<PackageEncryptionInfo>(m, "PackageEncryptionInfo")
        .def_readonly("encrypted", &PackageEncryptionInfo::encrypted)
        .def_readonly("format", &PackageEncryptionInfo::format)
        .def_readonly("version_major", &PackageEncryptionInfo::versionMajor)
        .def_readonly("version_minor", &PackageEncryptionInfo::versionMinor)
        .def_readonly("key_bits", &PackageEncryptionInfo::keyBits)
        .def_readonly("block_size", &PackageEncryptionInfo::blockSize)
        .def_readonly("hash_size", &PackageEncryptionInfo::hashSize)
        .def_readonly("spin_count", &PackageEncryptionInfo::spinCount)
        .def_readonly("hash_algorithm", &PackageEncryptionInfo::hashAlgorithm)
        .def_readonly("has_data_integrity", &PackageEncryptionInfo::hasDataIntegrity)
        .def_readonly("supported_for_read", &PackageEncryptionInfo::supportedForRead)
        .def_readonly("supported_for_write", &PackageEncryptionInfo::supportedForWrite)
        .def_readonly("cipher_algorithm", &PackageEncryptionInfo::cipherAlgorithm)
        .def_readonly("cipher_chaining", &PackageEncryptionInfo::cipherChaining)
        .def_readonly("key_encryptor_count", &PackageEncryptionInfo::keyEncryptorCount)
        .def_readonly("password_key_encryptor_count", &PackageEncryptionInfo::passwordKeyEncryptorCount)
        .def_readonly("certificate_key_encryptors", &PackageEncryptionInfo::certificateKeyEncryptors);

    py::class_<PackageEncryptionOptions>(m, "PackageEncryptionOptions")
        .def(py::init<>())
        .def_readwrite("enabled", &PackageEncryptionOptions::enabled)
        .def_readwrite("password", &PackageEncryptionOptions::password)
        .def_readwrite("mode", &PackageEncryptionOptions::mode)
        .def_readwrite("hash_algorithm", &PackageEncryptionOptions::hashAlgorithm)
        .def_readwrite("key_bits", &PackageEncryptionOptions::keyBits)
        .def_readwrite("spin_count", &PackageEncryptionOptions::spinCount);

    py::class_<SaveOptions>(m, "SaveOptions")
        .def(py::init<>())
        .def_readwrite("compression_level", &SaveOptions::compressionLevel)
        .def_readwrite("compression_strategy", &SaveOptions::compressionStrategy)
        .def_readwrite("parallel_workers", &SaveOptions::parallelWorkers)
        .def_readwrite("parallel_sheets", &SaveOptions::parallelSheets)
        .def_readwrite("parallel_rows", &SaveOptions::parallelRows)
        .def_readwrite("strict_namespace", &SaveOptions::strictNamespace)
        .def_readwrite("validate_model_before_save", &SaveOptions::validateModelBeforeSave)
        .def_readwrite("reject_model_warnings_before_save", &SaveOptions::rejectModelWarningsBeforeSave)
        .def_readwrite("validate_package_before_write", &SaveOptions::validatePackageBeforeWrite)
        .def_readwrite("encryption", &SaveOptions::encryption);

    py::class_<LoadOptions>(m, "LoadOptions")
        .def(py::init<>())
        .def_readwrite("lenient", &LoadOptions::lenient)
        .def_readwrite("max_entries", &LoadOptions::maxEntries)
        .def_readwrite("max_entry_bytes", &LoadOptions::maxEntryBytes)
        .def_readwrite("max_total_bytes", &LoadOptions::maxTotalBytes)
        .def_readwrite("max_file_bytes", &LoadOptions::maxFileBytes)
        .def_readwrite("password_to_open", &LoadOptions::passwordToOpen)
        .def_readwrite("max_encryption_spin_count", &LoadOptions::maxEncryptionSpinCount)
        .def_readwrite("max_decrypted_package_bytes", &LoadOptions::maxDecryptedPackageBytes)
        .def_readwrite("allow_standard_encryption", &LoadOptions::allowStandardEncryption)
        .def_readwrite("require_agile_data_integrity", &LoadOptions::requireAgileDataIntegrity)
        .def_readwrite("max_encryption_info_bytes", &LoadOptions::maxEncryptionInfoBytes)
        .def_readwrite("cancel", &LoadOptions::cancel)
        .def_readwrite("progress", &LoadOptions::progress);

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
        .def("__repr__", [](const CellReference& r) { return "<CellRef " + r.address() + ">"; });

    // === DateTime (aggregate) ===
    py::class_<DateTime>(m, "DateTime")
        .def(py::init<int, int, int, int, int, double>(),
             py::arg("year"), py::arg("month"), py::arg("day"),
             py::arg("hour") = 0, py::arg("minute") = 0, py::arg("second") = 0.0)
        .def_readonly("year", &DateTime::year)
        .def_readonly("month", &DateTime::month)
        .def_readonly("day", &DateTime::day)
        .def_readonly("hour", &DateTime::hour)
        .def_readonly("minute", &DateTime::minute)
        .def_readonly("second", &DateTime::second)
        .def_property_readonly("second_int", [](const DateTime& d) { return static_cast<int>(d.second); })
        .def_property_readonly("millisecond", [](const DateTime& d) {
            return static_cast<int>((d.second - static_cast<int>(d.second)) * 1000.0);
        })
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
        .value("GETTING_DATA", CellError::GettingData);

    m.def("cell_error_to_string", [](CellError e) { return toString(e); }, py::arg("error"));
    m.def("cell_error_from_string", [](const std::string& s) { return cellErrorFromString(s); }, py::arg("value"));

    // === Color ===
    py::class_<Color>(m, "Color")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("argb"))
        .def("set_argb", &Color::setArgb, py::return_value_policy::reference_internal)
        .def_property("argb", &Color::argb, &Color::setArgb)
        .def("empty", &Color::empty);

    // === Font ===
    py::class_<Font>(m, "Font")
        .def(py::init<>())
        .def_property("name", &Font::name, [](Font& f, std::string v) -> Font& { f.setName(std::move(v)); return f; })
        .def_property("size", &Font::size, [](Font& f, double v) -> Font& { f.setSize(v); return f; })
        .def_property("bold", &Font::bold, [](Font& f, bool v) -> Font& { f.setBold(v); return f; })
        .def_property("italic", &Font::italic, [](Font& f, bool v) -> Font& { f.setItalic(v); return f; })
        .def_property("underline", &Font::underline, [](Font& f, bool v) -> Font& { f.setUnderline(v); return f; })
        .def_property("strike", &Font::strike, [](Font& f, bool v) -> Font& { f.setStrike(v); return f; })
        .def("color", static_cast<Color& (Font::*)()>(&Font::color), py::return_value_policy::reference_internal);

    // === Fill ===
    py::class_<Fill>(m, "Fill")
        .def(py::init<>())
        .def_property("pattern_type", &Fill::patternType,
            [](Fill& f, std::string v) -> Fill& { f.setPatternType(std::move(v)); return f; })
        .def("foreground", static_cast<Color& (Fill::*)()>(&Fill::foregroundColor), py::return_value_policy::reference_internal)
        .def("background", static_cast<Color& (Fill::*)()>(&Fill::backgroundColor), py::return_value_policy::reference_internal);

    // === BorderSide ===
    py::class_<BorderSide>(m, "BorderSide")
        .def(py::init<>())
        .def_property("style", &BorderSide::style,
            [](BorderSide& b, std::string v) -> BorderSide& { b.setStyle(std::move(v)); return b; })
        .def("color", static_cast<Color& (BorderSide::*)()>(&BorderSide::color), py::return_value_policy::reference_internal);

    // === Border ===
    py::class_<Border>(m, "Border")
        .def(py::init<>())
        .def("left", static_cast<BorderSide& (Border::*)()>(&Border::left), py::return_value_policy::reference_internal)
        .def("right", static_cast<BorderSide& (Border::*)()>(&Border::right), py::return_value_policy::reference_internal)
        .def("top", static_cast<BorderSide& (Border::*)()>(&Border::top), py::return_value_policy::reference_internal)
        .def("bottom", static_cast<BorderSide& (Border::*)()>(&Border::bottom), py::return_value_policy::reference_internal)
        .def("diagonal", static_cast<BorderSide& (Border::*)()>(&Border::diagonal), py::return_value_policy::reference_internal);

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
            [](Alignment& a, int v) -> Alignment& { a.setIndent(v); return a; });

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
        .def("is_default", &Style::isDefault);

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
            [](RichText& rt) -> std::vector<RichTextRun>& { return rt.runs(); },
            [](RichText& rt, const std::vector<RichTextRun>& v) { rt.runs() = v; },
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
        .def_property("and_mode", &FilterColumn::andMode, &FilterColumn::setAndMode)
        .def_property("include_blank", &FilterColumn::includeBlank, &FilterColumn::setIncludeBlank);

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
        .def("set_reference", &AutoFilter::setReference)
        .def_property("reference", &AutoFilter::reference, &AutoFilter::setReference)
        .def_property_readonly("enabled", &AutoFilter::enabled)
        .def("clear", &AutoFilter::clear)
        .def("column", &AutoFilter::column, py::return_value_policy::reference_internal)
        .def("try_column", [](const AutoFilter& f, std::size_t id) -> py::object {
            const auto* c = f.tryColumn(id);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def("sort_state", [](AutoFilter& f) -> SortState& { return f.sortState(); },
             py::return_value_policy::reference_internal);

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

    // === Image ===
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
        .def_property("name", &Image::name, &Image::setName);

    // === Chart ===
    py::enum_<Chart::Type>(m, "ChartType")
        .value("BAR", Chart::Type::Bar).value("LINE", Chart::Type::Line).value("PIE", Chart::Type::Pie)
        .value("SCATTER", Chart::Type::Scatter).value("DOUGHNUT", Chart::Type::Doughnut)
        .value("RADAR", Chart::Type::Radar).value("AREA", Chart::Type::Area).value("BUBBLE", Chart::Type::Bubble);
    py::enum_<Chart::Grouping>(m, "ChartGrouping")
        .value("STANDARD", Chart::Grouping::Standard).value("STACKED", Chart::Grouping::Stacked)
        .value("PERCENT_STACKED", Chart::Grouping::PercentStacked).value("CLUSTERED", Chart::Grouping::Clustered);
    py::enum_<ChartColor::Kind>(m, "ChartColorKind")
        .value("None", ChartColor::Kind::None)
        .value("SRgb", ChartColor::Kind::SRgb)
        .value("Scheme", ChartColor::Kind::Scheme)
        .value("System", ChartColor::Kind::System)
        .value("Preset", ChartColor::Kind::Preset)
        .value("Unknown", ChartColor::Kind::Unknown)
        .export_values();

    py::class_<ChartColor>(m, "ChartColor")
        .def(py::init<>())
        .def_readwrite("kind", &ChartColor::kind)
        .def_readwrite("value", &ChartColor::value);

    py::class_<ChartCustomDashStop>(m, "ChartCustomDashStop")
        .def(py::init<>())
        .def_readwrite("dash", &ChartCustomDashStop::dash)
        .def_readwrite("space", &ChartCustomDashStop::space);

    py::class_<ChartLineFormat>(m, "ChartLineFormat")
        .def(py::init<>())
        .def_readwrite("present", &ChartLineFormat::present)
        .def_readwrite("no_fill", &ChartLineFormat::noFill)
        .def_readwrite("color", &ChartLineFormat::color)
        .def_readwrite("width_points", &ChartLineFormat::widthPoints)
        .def_readwrite("dash", &ChartLineFormat::dash)
        .def_readwrite("cap", &ChartLineFormat::cap)
        .def_readwrite("compound", &ChartLineFormat::compound)
        .def_readwrite("join", &ChartLineFormat::join);

    py::class_<ChartGradientStop>(m, "ChartGradientStop")
        .def(py::init<>())
        .def_readwrite("position", &ChartGradientStop::position)
        .def_readwrite("color", &ChartGradientStop::color);

    py::enum_<ChartFillFormat::Kind>(m, "ChartFillKind")
        .value("None", ChartFillFormat::Kind::None)
        .value("NoFill", ChartFillFormat::Kind::NoFill)
        .value("Solid", ChartFillFormat::Kind::Solid)
        .value("Gradient", ChartFillFormat::Kind::Gradient)
        .value("Pattern", ChartFillFormat::Kind::Pattern)
        .export_values();

    py::class_<ChartFillFormat>(m, "ChartFillFormat")
        .def(py::init<>())
        .def_readwrite("present", &ChartFillFormat::present)
        .def_readwrite("no_fill", &ChartFillFormat::noFill)
        .def_readwrite("color", &ChartFillFormat::color)
        .def_readwrite("kind", &ChartFillFormat::kind)
        .def_readwrite("gradient_angle_degrees", &ChartFillFormat::gradientAngleDegrees)
        .def_readwrite("pattern", &ChartFillFormat::pattern)
        .def_readwrite("foreground_color", &ChartFillFormat::foregroundColor)
        .def_readwrite("background_color", &ChartFillFormat::backgroundColor);

    py::class_<ChartTextRun>(m, "ChartTextRun")
        .def(py::init<>())
        .def_readwrite("text", &ChartTextRun::text)
        .def_readwrite("bold", &ChartTextRun::bold)
        .def_readwrite("italic", &ChartTextRun::italic)
        .def_readwrite("font_size_points", &ChartTextRun::fontSizePoints)
        .def_readwrite("typeface", &ChartTextRun::typeface)
        .def_readwrite("color", &ChartTextRun::color);

    py::class_<ChartTextStyle>(m, "ChartTextStyle")
        .def(py::init<>())
        .def_readwrite("present", &ChartTextStyle::present)
        .def_readwrite("bold", &ChartTextStyle::bold)
        .def_readwrite("italic", &ChartTextStyle::italic)
        .def_readwrite("font_size_points", &ChartTextStyle::fontSizePoints)
        .def_readwrite("typeface", &ChartTextStyle::typeface)
        .def_readwrite("color", &ChartTextStyle::color);

    py::class_<ChartRichText>(m, "ChartRichText")
        .def(py::init<>())
        .def_readwrite("present", &ChartRichText::present)
        .def_readwrite("runs", &ChartRichText::runs)
        .def("plain_text", &ChartRichText::plainText);

    py::class_<ChartCachePoint>(m, "ChartCachePoint")
        .def(py::init<>())
        .def_readwrite("index", &ChartCachePoint::index)
        .def_readwrite("value", &ChartCachePoint::value);

    py::class_<ChartSeriesCache>(m, "ChartSeriesCache")
        .def(py::init<>())
        .def_readwrite("present", &ChartSeriesCache::present)
        .def_readwrite("numeric", &ChartSeriesCache::numeric)
        .def_readwrite("format_code", &ChartSeriesCache::formatCode)
        .def_readwrite("point_count", &ChartSeriesCache::pointCount)
        .def_readwrite("points", &ChartSeriesCache::points)
        .def("effective_point_count", &ChartSeriesCache::effectivePointCount)
        .def("valid", &ChartSeriesCache::valid, py::arg("allow_sparse") = true);

    py::class_<ChartMarkerFormat>(m, "ChartMarkerFormat")
        .def(py::init<>())
        .def_readwrite("present", &ChartMarkerFormat::present)
        .def_readwrite("symbol", &ChartMarkerFormat::symbol)
        .def_readwrite("size", &ChartMarkerFormat::size)
        .def_readwrite("fill", &ChartMarkerFormat::fill)
        .def_readwrite("line", &ChartMarkerFormat::line);

    py::class_<ChartDataLabelPoint>(m, "ChartDataLabelPoint")
        .def(py::init<>())
        .def_readwrite("index", &ChartDataLabelPoint::index)
        .def_readwrite("deleted", &ChartDataLabelPoint::deleted)
        .def_readwrite("show_value", &ChartDataLabelPoint::showValue)
        .def_readwrite("show_category_name", &ChartDataLabelPoint::showCategoryName)
        .def_readwrite("show_series_name", &ChartDataLabelPoint::showSeriesName)
        .def_readwrite("show_percent", &ChartDataLabelPoint::showPercent)
        .def_readwrite("position", &ChartDataLabelPoint::position);

    py::class_<ChartDataLabels>(m, "ChartDataLabels")
        .def(py::init<>())
        .def_readwrite("present", &ChartDataLabels::present)
        .def_readwrite("show_value", &ChartDataLabels::showValue)
        .def_readwrite("show_category_name", &ChartDataLabels::showCategoryName)
        .def_readwrite("show_series_name", &ChartDataLabels::showSeriesName)
        .def_readwrite("show_percent", &ChartDataLabels::showPercent)
        .def_readwrite("position", &ChartDataLabels::position)
        .def_readwrite("separator", &ChartDataLabels::separator);

    py::class_<ChartDataTable>(m, "ChartDataTable")
        .def(py::init<>())
        .def_readwrite("present", &ChartDataTable::present)
        .def_readwrite("show_horizontal_border", &ChartDataTable::showHorizontalBorder)
        .def_readwrite("show_vertical_border", &ChartDataTable::showVerticalBorder)
        .def_readwrite("show_outline", &ChartDataTable::showOutline)
        .def_readwrite("show_legend_keys", &ChartDataTable::showLegendKeys)
        .def_readwrite("fill", &ChartDataTable::fill)
        .def_readwrite("line", &ChartDataTable::line);

    py::class_<ChartManualLayout>(m, "ChartManualLayout")
        .def(py::init<>())
        .def_readwrite("present", &ChartManualLayout::present)
        .def_readwrite("target", &ChartManualLayout::target)
        .def_readwrite("x_mode", &ChartManualLayout::xMode)
        .def_readwrite("y_mode", &ChartManualLayout::yMode)
        .def_readwrite("width_mode", &ChartManualLayout::widthMode)
        .def_readwrite("height_mode", &ChartManualLayout::heightMode)
        .def_readwrite("has_x", &ChartManualLayout::hasX)
        .def_readwrite("has_y", &ChartManualLayout::hasY)
        .def_readwrite("has_width", &ChartManualLayout::hasWidth)
        .def_readwrite("has_height", &ChartManualLayout::hasHeight)
        .def_readwrite("x", &ChartManualLayout::x)
        .def_readwrite("y", &ChartManualLayout::y)
        .def_readwrite("width", &ChartManualLayout::width)
        .def_readwrite("height", &ChartManualLayout::height);

    py::class_<ChartAxisScaling>(m, "ChartAxisScaling")
        .def(py::init<>())
        .def_readwrite("has_minimum", &ChartAxisScaling::hasMinimum)
        .def_readwrite("has_maximum", &ChartAxisScaling::hasMaximum)
        .def_readwrite("has_log_base", &ChartAxisScaling::hasLogBase)
        .def_readwrite("minimum", &ChartAxisScaling::minimum)
        .def_readwrite("maximum", &ChartAxisScaling::maximum)
        .def_readwrite("log_base", &ChartAxisScaling::logBase)
        .def_readwrite("reverse_order", &ChartAxisScaling::reverseOrder);

    py::class_<ChartDisplayUnits>(m, "ChartDisplayUnits")
        .def(py::init<>())
        .def_readwrite("present", &ChartDisplayUnits::present)
        .def_readwrite("built_in_unit", &ChartDisplayUnits::builtInUnit)
        .def_readwrite("has_custom_unit", &ChartDisplayUnits::hasCustomUnit)
        .def_readwrite("custom_unit", &ChartDisplayUnits::customUnit)
        .def_readwrite("show_label", &ChartDisplayUnits::showLabel);

    py::class_<ChartWallFormat>(m, "ChartWallFormat")
        .def(py::init<>())
        .def_readwrite("present", &ChartWallFormat::present)
        .def_readwrite("has_thickness", &ChartWallFormat::hasThickness)
        .def_readwrite("thickness", &ChartWallFormat::thickness)
        .def_readwrite("fill", &ChartWallFormat::fill)
        .def_readwrite("line", &ChartWallFormat::line);

    py::class_<ChartView3D>(m, "ChartView3D")
        .def(py::init<>())
        .def_readwrite("present", &ChartView3D::present)
        .def_readwrite("has_rotation_x", &ChartView3D::hasRotationX)
        .def_readwrite("has_rotation_y", &ChartView3D::hasRotationY)
        .def_readwrite("has_height_percent", &ChartView3D::hasHeightPercent)
        .def_readwrite("has_depth_percent", &ChartView3D::hasDepthPercent)
        .def_readwrite("right_angle_axes", &ChartView3D::rightAngleAxes)
        .def_readwrite("perspective", &ChartView3D::perspective);

    py::enum_<ChartSeries::TrendlineType>(m, "TrendlineType")
        .value("Linear", ChartSeries::TrendlineType::Linear)
        .value("Exponential", ChartSeries::TrendlineType::Exponential)
        .value("Logarithmic", ChartSeries::TrendlineType::Logarithmic)
        .value("Polynomial", ChartSeries::TrendlineType::Polynomial)
        .value("Power", ChartSeries::TrendlineType::Power)
        .value("MovingAverage", ChartSeries::TrendlineType::MovingAverage)
        .export_values();

    py::enum_<ChartSeries::ErrorBarDirection>(m, "ErrorBarDirection")
        .value("X", ChartSeries::ErrorBarDirection::X)
        .value("Y", ChartSeries::ErrorBarDirection::Y)
        .export_values();

    py::enum_<ChartSeries::ErrorBarType>(m, "ErrorBarType")
        .value("Both", ChartSeries::ErrorBarType::Both)
        .value("Plus", ChartSeries::ErrorBarType::Plus)
        .value("Minus", ChartSeries::ErrorBarType::Minus)
        .export_values();

    py::enum_<ChartSeries::ErrorValueType>(m, "ErrorValueType")
        .value("FixedValue", ChartSeries::ErrorValueType::FixedValue)
        .value("Percentage", ChartSeries::ErrorValueType::Percentage)
        .value("StandardDeviation", ChartSeries::ErrorValueType::StandardDeviation)
        .value("StandardError", ChartSeries::ErrorValueType::StandardError)
        .value("Custom", ChartSeries::ErrorValueType::Custom)
        .export_values();

    py::class_<ChartSeries::Trendline>(m, "Trendline")
        .def(py::init<>())
        .def_readwrite("type", &ChartSeries::Trendline::type)
        .def_readwrite("order", &ChartSeries::Trendline::order)
        .def_readwrite("period", &ChartSeries::Trendline::period)
        .def_readwrite("forward", &ChartSeries::Trendline::forward)
        .def_readwrite("backward", &ChartSeries::Trendline::backward)
        .def_readwrite("display_equation", &ChartSeries::Trendline::displayEquation)
        .def_readwrite("display_r_squared", &ChartSeries::Trendline::displayRSquared)
        .def_readwrite("line_format", &ChartSeries::Trendline::lineFormat);

    py::class_<ChartSeries::ErrorBars>(m, "ErrorBars")
        .def(py::init<>())
        .def_readwrite("direction", &ChartSeries::ErrorBars::direction)
        .def_readwrite("bar_type", &ChartSeries::ErrorBars::barType)
        .def_readwrite("value_type", &ChartSeries::ErrorBars::valueType)
        .def_readwrite("value", &ChartSeries::ErrorBars::value)
        .def_readwrite("no_end_cap", &ChartSeries::ErrorBars::noEndCap)
        .def_readwrite("plus_reference", &ChartSeries::ErrorBars::plusReference)
        .def_readwrite("minus_reference", &ChartSeries::ErrorBars::minusReference)
        .def_readwrite("line_format", &ChartSeries::ErrorBars::lineFormat);

    py::class_<ChartSeries>(m, "ChartSeries")
        .def(py::init<std::string>())
        .def_property("title", &ChartSeries::title, &ChartSeries::setTitle)
        .def_property("values_reference", &ChartSeries::valuesReference, &ChartSeries::setValuesReference)
        .def_property("categories_reference", &ChartSeries::categoriesReference, &ChartSeries::setCategoriesReference)
        .def("reference", &ChartSeries::reference, py::arg("sheet_name"), py::arg("range_ref"))
        .def("categories", &ChartSeries::categories, py::arg("sheet_name"), py::arg("range_ref"));
    py::class_<Chart>(m, "Chart")
        .def(py::init<Chart::Type>(), py::arg("type") = Chart::Type::Bar)
        .def_property_readonly("type", &Chart::type)
        .def_property("grouping", &Chart::grouping, &Chart::setGrouping)
        .def_property("title", &Chart::title, &Chart::setTitle)
        .def_property("x_axis_title", &Chart::xAxisTitle, &Chart::setXAxisTitle)
        .def_property("y_axis_title", &Chart::yAxisTitle, &Chart::setYAxisTitle)
        .def_property("style", &Chart::style, &Chart::setStyle)
        .def_property("stable_id", &Chart::stableId, &Chart::setStableId)
        .def_property("width", &Chart::width, &Chart::setWidth)
        .def_property("height", &Chart::height, &Chart::setHeight)
        .def_property("show_legend", &Chart::showLegend, &Chart::setShowLegend)
        .def_property("legend_position", &Chart::legendPosition, &Chart::setLegendPosition)
        .def("add_series", &Chart::addSeries, py::return_value_policy::reference_internal)
        .def_property_readonly("series", [](Chart& c) -> std::vector<ChartSeries>& { return c.series(); }, py::return_value_policy::reference_internal)
        .def_static("type_name", &Chart::typeName, py::arg("type"), py::arg("grouping") = Chart::Grouping::Standard);

    // === Pivot tables ===
    py::class_<PivotCache>(m, "PivotCache")
        .def(py::init<>())
        .def_property("cache_id", &PivotCache::cacheId, &PivotCache::setCacheId)
        .def_property("source_data", &PivotCache::sourceData, &PivotCache::setSourceData);

    py::class_<PivotField>(m, "PivotField")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("name"))
        .def_property("name", &PivotField::name, &PivotField::setName)
        .def_property("axis", &PivotField::axis, &PivotField::setAxis)
        .def_property("show_all", &PivotField::showAll, &PivotField::setShowAll)
        .def_property("sort_type", &PivotField::sortType, &PivotField::setSortType);

    py::class_<PivotFieldReference>(m, "PivotFieldReference")
        .def(py::init<>())
        .def_property("field_index", &PivotFieldReference::fieldIndex, &PivotFieldReference::setFieldIndex);

    py::class_<PivotTable>(m, "PivotTable")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("name"))
        .def_property("name", &PivotTable::name, &PivotTable::setName)
        .def_property("location", &PivotTable::location, &PivotTable::setLocation)
        .def_property_readonly("cache", [](PivotTable& p) -> PivotCache& { return p.cache(); }, py::return_value_policy::reference_internal)
        .def("add_row_field", &PivotTable::addRowField)
        .def("add_column_field", &PivotTable::addColumnField)
        .def("add_page_field", &PivotTable::addPageField)
        .def("add_data_field", [](PivotTable& p) { p.addDataField(); })
        .def_property_readonly("row_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.rowFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("column_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.columnFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("page_fields", [](PivotTable& p) -> std::vector<PivotField>& { return p.pageFields(); }, py::return_value_policy::reference_internal)
        .def_property_readonly("data_fields", [](PivotTable& p) -> std::vector<PivotFieldReference>& { return p.dataFields(); }, py::return_value_policy::reference_internal);

    // === Table ===
    py::class_<TableColumn>(m, "TableColumn")
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
        .def_property_readonly("columns", [](Table& t) -> std::vector<TableColumn>& { return t.columns(); }, py::return_value_policy::reference_internal)
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
        .def_property_readonly("rules", [](ConditionalFormattingEntry& e) -> std::vector<ConditionalRule>& { return e.rules(); }, py::return_value_policy::reference_internal)
        .def("empty", &ConditionalFormattingEntry::empty);

    py::class_<ConditionalFormattingCollection>(m, "ConditionalFormattingCollection")
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
        .def_property_readonly("entries", [](ConditionalFormattingCollection& c) -> std::vector<ConditionalFormattingEntry>& { return c.entries(); }, py::return_value_policy::reference_internal)
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
        .def("__call__", [](DataValidationCollection& c) -> DataValidationCollection& { return c; }, py::return_value_policy::reference_internal)
        .def("add", [](DataValidationCollection& c, DataValidationType type,
                       const std::string& reference) -> DataValidation& {
                return c.add(type, reference);
            }, py::arg("type"), py::arg("reference"),
            py::return_value_policy::reference_internal)
        .def("add_validation", [](DataValidationCollection& c, DataValidation v) -> DataValidation& {
                return c.add(std::move(v));
            }, py::arg("validation"), py::return_value_policy::reference_internal)
        .def_property_readonly("items", [](DataValidationCollection& c) -> std::vector<DataValidation>& { return c.items(); }, py::return_value_policy::reference_internal)
        .def("empty", &DataValidationCollection::empty)
        .def("clear", &DataValidationCollection::clear);

    // === Protection ===
    py::class_<WorksheetProtection>(m, "WorksheetProtection")
        .def(py::init<>())
        .def_property("enabled", &WorksheetProtection::enabled, &WorksheetProtection::setEnabled)
        .def_property("password_hash", &WorksheetProtection::passwordHash, &WorksheetProtection::setPasswordHash)
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
        .def_property("workbook_password_hash", &WorkbookProtection::workbookPasswordHash, &WorkbookProtection::setWorkbookPasswordHash);

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
    py::class_<CalcProperties>(m, "CalcProperties")
        .def(py::init<>())
        .def_property("calc_id", &CalcProperties::calcId, &CalcProperties::setCalcId)
        .def_property("calc_mode", &CalcProperties::calcMode, &CalcProperties::setCalcMode)
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
        .def_property_readonly("items", [](CustomProperties& c) -> std::vector<CustomProperty>& { return c.items(); }, py::return_value_policy::reference_internal)
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
        .def("has_rich_text", &Cell::hasRichText)
        .def("rich_text_value", &Cell::richTextValue)
        .def("set_rich_text", &Cell::setRichText)
        .def("clear_rich_text", &Cell::clearRichText)
        .def("clear", &Cell::clear)
        .def("empty", &Cell::empty)
        .def("is_numeric", &Cell::isNumeric)
        .def("is_string", &Cell::isString)
        .def("is_bool", &Cell::isBoolean)
        .def("is_date", &Cell::isDate)
        .def("value_type", &Cell::valueType)
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
        .def_property_readonly("tables", [](Worksheet& ws) -> std::vector<Table>& { return ws.tables(); }, py::return_value_policy::reference_internal)
        .def("add_image", [](Worksheet& ws, const std::string& path, const std::string& anchor) -> Image& {
            return ws.addImage(path, anchor);
        }, py::return_value_policy::reference_internal)
        .def("add_image", [](Worksheet& ws, const Image& image) -> Image& {
            return ws.addImage(image);
        }, py::return_value_policy::reference_internal)
        .def_property_readonly("images", [](Worksheet& ws) -> std::vector<Image>& { return ws.images(); }, py::return_value_policy::reference_internal)
        .def("sheet_view", py::overload_cast<>(&Worksheet::sheetView),
             py::return_value_policy::reference_internal)
        .def("set_sheet_view", &Worksheet::setSheetView)
        .def("add_chart", &Worksheet::addChart)
        .def("chart", py::overload_cast<std::size_t>(&Worksheet::chart), py::return_value_policy::reference_internal)
        .def_property_readonly("chart_count", &Worksheet::chartCount)
        .def_property_readonly("charts", [](Worksheet& ws) -> std::vector<Chart>& { return ws.charts(); }, py::return_value_policy::reference_internal)
        .def("move_image", &Worksheet::moveImage)
        .def("move_image_absolute", &Worksheet::moveImageAbsolute, py::arg("stable_id"), py::arg("x_emu"), py::arg("y_emu"))
        .def("resize_image", &Worksheet::resizeImage)
        .def("remove_image", &Worksheet::removeImage)
        .def("move_chart", &Worksheet::moveChart)
        .def("move_chart_absolute", &Worksheet::moveChartAbsolute, py::arg("stable_id"), py::arg("x_emu"), py::arg("y_emu"))
        .def("resize_chart", &Worksheet::resizeChart)
        .def("remove_chart", &Worksheet::removeChart)
        .def("set_chart_title", &Worksheet::setChartTitle)
        .def("set_chart_style", &Worksheet::setChartStyle)
        .def("set_chart_title_rich_text", &Worksheet::setChartTitleRichText)
        .def("set_chart_x_axis_title", &Worksheet::setChartXAxisTitle)
        .def("set_chart_y_axis_title", &Worksheet::setChartYAxisTitle)
        .def("set_chart_axis_title", &Worksheet::setChartAxisTitle, py::arg("stable_id"), py::arg("axis_id"), py::arg("title"))
        .def("set_chart_axis_number_format", &Worksheet::setChartAxisNumberFormat, py::arg("stable_id"), py::arg("axis_id"), py::arg("format_code"), py::arg("source_linked") = true)
        .def("set_chart_axis_units", &Worksheet::setChartAxisUnits, py::arg("stable_id"), py::arg("axis_id"), py::arg("major_unit"), py::arg("minor_unit") = 0.0)
        .def("set_chart_axis_scaling", &Worksheet::setChartAxisScaling, py::arg("stable_id"), py::arg("axis_id"), py::arg("scaling"))
        .def("set_chart_axis_crossing", &Worksheet::setChartAxisCrossing, py::arg("stable_id"), py::arg("axis_id"), py::arg("crosses"), py::arg("cross_between") = std::string{})
        .def("set_chart_axis_crosses_at", &Worksheet::setChartAxisCrossesAt, py::arg("stable_id"), py::arg("axis_id"), py::arg("crosses_at"))
        .def("clear_chart_axis_crosses_at", &Worksheet::clearChartAxisCrossesAt)
        .def("set_chart_axis_display_units", &Worksheet::setChartAxisDisplayUnits, py::arg("stable_id"), py::arg("axis_id"), py::arg("display_units"))
        .def("clear_chart_axis_display_units", &Worksheet::clearChartAxisDisplayUnits)
        .def("set_chart_axis_line_format", &Worksheet::setChartAxisLineFormat, py::arg("stable_id"), py::arg("axis_id"), py::arg("format"))
        .def("set_chart_axis_gridline_format", &Worksheet::setChartAxisGridlineFormat, py::arg("stable_id"), py::arg("axis_id"), py::arg("major"), py::arg("format"))
        .def("remove_chart_axis_gridlines", &Worksheet::removeChartAxisGridlines)
        .def("set_chart_area_line_format", &Worksheet::setChartAreaLineFormat)
        .def("set_chart_area_fill_format", &Worksheet::setChartAreaFillFormat)
        .def("set_chart_plot_area_line_format", &Worksheet::setChartPlotAreaLineFormat)
        .def("set_chart_plot_area_fill_format", &Worksheet::setChartPlotAreaFillFormat)
        .def("set_chart_plot_area_layout", &Worksheet::setChartPlotAreaLayout, py::arg("stable_id"), py::arg("layout"))
        .def("set_chart_data_table", &Worksheet::setChartDataTable, py::arg("stable_id"), py::arg("table"))
        .def("remove_chart_data_table", &Worksheet::removeChartDataTable)
        .def("set_chart_legend", &Worksheet::setChartLegend, py::arg("stable_id"), py::arg("show"), py::arg("position") = std::string{})
        .def("set_chart_view_3d", &Worksheet::setChartView3D, py::arg("stable_id"), py::arg("view"))
        .def("set_chart_floor_format", &Worksheet::setChartFloorFormat)
        .def("set_chart_side_wall_format", &Worksheet::setChartSideWallFormat)
        .def("set_chart_back_wall_format", &Worksheet::setChartBackWallFormat)
        .def("set_chart_series_value_cache", &Worksheet::setChartSeriesValueCache, py::arg("stable_id"), py::arg("series_index"), py::arg("cache"))
        .def("set_chart_series_line_format", &Worksheet::setChartSeriesLineFormat, py::arg("stable_id"), py::arg("series_index"), py::arg("format"))
        .def("set_chart_series_fill_format", &Worksheet::setChartSeriesFillFormat, py::arg("stable_id"), py::arg("series_index"), py::arg("format"))
        .def("add_chart_series_trendline", &Worksheet::addChartSeriesTrendline, py::arg("stable_id"), py::arg("series_index"), py::arg("trendline"))
        .def("set_chart_series_error_bars", &Worksheet::setChartSeriesErrorBars, py::arg("stable_id"), py::arg("series_index"), py::arg("error_bars"))
        .def("add_pivot_table", &Worksheet::addPivotTable)
        .def_property_readonly("pivot_tables", [](Worksheet& ws) -> std::vector<PivotTable>& { return ws.pivotTables(); }, py::return_value_policy::reference_internal)
        .def("row", [](Worksheet& ws, std::size_t n) { return ws.row(n); })
        .def("iter_rows", [](const Worksheet& ws, std::size_t minRow, std::size_t maxRow, std::size_t minCol, std::size_t maxCol) {
            return ws.iterRows(minRow, maxRow, minCol, maxCol);
        }, py::arg("min_row") = 0, py::arg("max_row") = 0, py::arg("min_col") = 0, py::arg("max_col") = 0)
        .def("iter_cols", [](const Worksheet& ws, std::size_t minRow, std::size_t maxRow, std::size_t minCol, std::size_t maxCol) {
            return ws.iterCols(minRow, maxRow, minCol, maxCol);
        }, py::arg("min_row") = 0, py::arg("max_row") = 0, py::arg("min_col") = 0, py::arg("max_col") = 0)
        .def("dirty", &Worksheet::dirty)
        .def("mark_dirty", &Worksheet::markDirty)
        .def("save_csv", [](const Worksheet& ws, const std::string& path) { ws.saveCsv(std::filesystem::path(path)); })
        .def("load_csv", [](Worksheet& ws, const std::string& path) { ws.loadCsv(std::filesystem::path(path)); })
        .def_property_readonly("tracked_cell_change_count", &Worksheet::trackedCellChangeCount)
        .def_property_readonly("has_tracked_cell_changes", &Worksheet::hasTrackedCellChanges)
        .def("clear_tracked_cell_changes", &Worksheet::clearTrackedCellChanges)
        .def_property_readonly("cells", [](const Worksheet& ws) -> py::dict {
            py::dict out;
            for (const auto& [key, cell] : ws.cells())
                out[py::cast(cell.address())] = py::cast(&cell, py::return_value_policy::reference);
            return out;
        })
        .def_property_readonly("row_dimensions", [](const Worksheet& ws) -> py::dict {
            py::dict out;
            for (const auto& [row, dim] : ws.rowDimensions())
                out[py::cast(row)] = py::cast(&dim, py::return_value_policy::reference);
            return out;
        })
        .def_property_readonly("column_dimensions", [](const Worksheet& ws) -> py::dict {
            py::dict out;
            for (const auto& [col, dim] : ws.columnDimensions())
                out[py::cast(col)] = py::cast(&dim, py::return_value_policy::reference);
            return out;
        })
        .def("__iter__", [](Worksheet& ws) {
            auto rows = ws.rows();
            py::list result;
            for (auto& row : rows) result.append(py::cast(row));
            return py::iter(result);
        })
        .def("__repr__", [](const Worksheet& ws) { return "<Worksheet '" + ws.name() + "'>"; });

    // === WorksheetExtents ===
    py::class_<WorksheetExtents>(m, "WorksheetExtents")
        .def_readonly("min_row", &WorksheetExtents::minRow)
        .def_readonly("min_column", &WorksheetExtents::minColumn)
        .def_readonly("max_row", &WorksheetExtents::maxRow)
        .def_readonly("max_column", &WorksheetExtents::maxColumn);

    // === Row ===
    py::class_<Row>(m, "Row")
        .def_property_readonly("number", &Row::number)
        .def("cell", &Row::cell, py::return_value_policy::reference_internal)
        .def("try_cell", [](const Row& r, std::size_t col) -> py::object {
            const auto* c = r.tryCell(col);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
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
        .def_readwrite("address", &StreamingCell::address)
        .def_property("value", [](const StreamingCell& c) { return cellvalue_to_py(c.value); },
            [](StreamingCell& c, const py::object& v) { c.value = py_to_cellvalue(v); })
        .def_readwrite("formula", &StreamingCell::formula)
        .def_readwrite("style_index", &StreamingCell::styleIndex);

    py::class_<StreamingRowIterator>(m, "StreamingRowIterator")
        .def(py::init<>())
        .def("row_number", &StreamingRowIterator::rowNumber);

    py::class_<StreamingWorksheetReader>(m, "StreamingWorksheetReader")
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
        .def("worksheet_names", &StreamingWorkbookReader::worksheetNames)
        .def("worksheet", [](StreamingWorkbookReader& r, const std::string& name) {
            return r.worksheet(name);
        }, py::arg("worksheet_name"), py::return_value_policy::move)
        .def("for_each_row", &StreamingWorkbookReader::forEachRow, py::arg("worksheet_name"), py::arg("callback"));

    py::class_<StreamingWorksheetWriter>(m, "StreamingWorksheetWriter")
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

    // === Workbook ===
    py::enum_<WorkbookSheetVisibility>(m, "WorkbookSheetVisibility")
        .value("Visible", WorkbookSheetVisibility::Visible)
        .value("Hidden", WorkbookSheetVisibility::Hidden)
        .value("VeryHidden", WorkbookSheetVisibility::VeryHidden)
        .export_values();

    py::class_<ChartCacheSyncOptions>(m, "ChartCacheSyncOptions")
        .def(py::init<>())
        .def_readwrite("synchronize_titles", &ChartCacheSyncOptions::synchronizeTitles)
        .def_readwrite("synchronize_categories", &ChartCacheSyncOptions::synchronizeCategories)
        .def_readwrite("synchronize_values", &ChartCacheSyncOptions::synchronizeValues)
        .def_readwrite("only_changed_cells", &ChartCacheSyncOptions::onlyChangedCells)
        .def_readwrite("preserve_formula_cached_values", &ChartCacheSyncOptions::preserveFormulaCachedValues)
        .def_readwrite("clear_tracked_changes_after_sync", &ChartCacheSyncOptions::clearTrackedChangesAfterSync)
        .def_readwrite("propagate_formula_dependencies", &ChartCacheSyncOptions::propagateFormulaDependencies)
        .def_readwrite("request_host_recalculation", &ChartCacheSyncOptions::requestHostRecalculationForFormulaDependencies)
        .def_readwrite("max_formula_dependency_depth", &ChartCacheSyncOptions::maxFormulaDependencyDepth)
        .def_readwrite("clear_unsupported_references", &ChartCacheSyncOptions::clearUnsupportedReferences);

    py::class_<ChartCacheSyncReport>(m, "ChartCacheSyncReport")
        .def_readonly("charts_visited", &ChartCacheSyncReport::chartsVisited)
        .def_readonly("series_visited", &ChartCacheSyncReport::seriesVisited)
        .def_readonly("dependencies_visited", &ChartCacheSyncReport::dependenciesVisited)
        .def_readonly("dependencies_matched", &ChartCacheSyncReport::dependenciesMatched)
        .def_readonly("dependencies_skipped_unchanged", &ChartCacheSyncReport::dependenciesSkippedUnchanged)
        .def_readonly("caches_updated", &ChartCacheSyncReport::cachesUpdated)
        .def_readonly("caches_cleared", &ChartCacheSyncReport::cachesCleared)
        .def_readonly("references_skipped", &ChartCacheSyncReport::referencesSkipped)
        .def_readonly("formula_cache_points_reused", &ChartCacheSyncReport::formulaCachePointsReused);

    py::class_<ChartStyleApplyReport>(m, "ChartStyleApplyReport")
        .def_readonly("series_visited", &ChartStyleApplyReport::seriesVisited)
        .def_readonly("series_styled", &ChartStyleApplyReport::seriesStyled)
        .def_readonly("colors_available", &ChartStyleApplyReport::colorsAvailable)
        .def_readonly("fill_styles_available", &ChartStyleApplyReport::fillStylesAvailable)
        .def_readonly("line_styles_available", &ChartStyleApplyReport::lineStylesAvailable);

    py::enum_<VbaModuleType>(m, "VbaModuleType")
        .value("Standard", VbaModuleType::Standard)
        .value("Document", VbaModuleType::Document)
        .value("Class", VbaModuleType::Class)
        .value("Designer", VbaModuleType::Designer)
        .export_values();

    py::class_<VbaModule>(m, "VbaModule")
        .def(py::init<>())
        .def_readwrite("name", &VbaModule::name)
        .def_readwrite("source", &VbaModule::source)
        .def_readwrite("type", &VbaModule::type);

    py::class_<VbaDesignerStorage>(m, "VbaDesignerStorage")
        .def(py::init<>())
        .def_readwrite("name", &VbaDesignerStorage::name);

    py::class_<VbaProjectInfo>(m, "VbaProjectInfo")
        .def(py::init<>())
        .def_readwrite("name", &VbaProjectInfo::name)
        .def_readwrite("description", &VbaProjectInfo::description)
        .def_readwrite("help_file", &VbaProjectInfo::helpFile)
        .def_readwrite("help_context_id", &VbaProjectInfo::helpContextId)
        .def_readwrite("constants", &VbaProjectInfo::constants)
        .def_readwrite("code_page", &VbaProjectInfo::codePage)
        .def_readwrite("project_id", &VbaProjectInfo::projectId);

    py::class_<xlpp::Chartsheet>(m, "Chartsheet")
        .def_property_readonly("name", &xlpp::Chartsheet::name)
        .def_property_readonly("has_chart", &xlpp::Chartsheet::hasChart)
        .def("chart", static_cast<xlpp::Chart& (xlpp::Chartsheet::*)()>(&xlpp::Chartsheet::chart),
             py::return_value_policy::reference_internal);

    py::class_<Workbook>(m, "Workbook")
        .def(py::init<>())
        .def("add_worksheet", &Workbook::addWorksheet, py::return_value_policy::reference_internal)
        .def("copy_worksheet", &Workbook::copyWorksheet, py::return_value_policy::reference_internal)
        .def("remove_worksheet", &Workbook::removeWorksheet)
        .def("rename_worksheet", &Workbook::renameWorksheet, py::arg("old_name"), py::arg("new_name"))
        .def("worksheet",
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
        .def_property_readonly("sheet_names", &Workbook::sheetNames)
        .def_static("is_password_encrypted_file", [](const std::string& path) {
            return Workbook::isPasswordEncryptedFile(std::filesystem::path(path));
        })
        .def_static("inspect_password_encryption_file", [](const std::string& path) {
            return Workbook::inspectPasswordEncryptionFile(std::filesystem::path(path));
        })
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
        .def("defined_name", py::overload_cast<const std::string&>(&Workbook::definedName),
             py::return_value_policy::reference_internal)
        .def_property_readonly("defined_names", [](Workbook& wb) -> py::list {
            py::list out;
            for (auto& n : wb.definedNames())
                out.append(py::cast(&n, py::return_value_policy::reference));
            return out;
        })
        .def_property("date_1904", &Workbook::date1904, &Workbook::setDate1904)
        .def("clear", &Workbook::clear)
        .def("add_chartsheet", &Workbook::addChartsheet, py::return_value_policy::reference_internal)
        .def("chartsheet", [](Workbook& wb, const std::string& name) -> py::object {
            auto* c = wb.chartsheet(name);
            return c ? py::cast(*c, py::return_value_policy::reference) : py::none{};
        })
        .def("remove_chartsheet", &Workbook::removeChartsheet)
        .def("rename_chartsheet", &Workbook::renameChartsheet)
        .def_property_readonly("chartsheet_count", &Workbook::chartsheetCount)
        .def_property_readonly("workbook_sheet_count", &Workbook::workbookSheetCount)
        .def_property_readonly("workbook_sheet_names", &Workbook::workbookSheetNames)
        .def("workbook_sheet_visibility", [](Workbook& wb, std::size_t index) {
            return wb.workbookSheetVisibility(index);
        })
        .def("set_workbook_sheet_visibility", [](Workbook& wb, std::size_t index, WorkbookSheetVisibility v) {
            wb.setWorkbookSheetVisibility(index, v);
        }, py::arg("index"), py::arg("visibility"))
        .def("set_active_workbook_sheet_index", &Workbook::setActiveWorkbookSheetIndex)
        .def("set_active_workbook_sheet", &Workbook::setActiveWorkbookSheet)
        .def("move_workbook_sheet", &Workbook::moveWorkbookSheet)
        .def("synchronize_chart_caches",
            [](Workbook& wb, const ChartCacheSyncOptions& options) { return wb.synchronizeChartCaches(options); },
            py::arg("options") = ChartCacheSyncOptions{})
        .def("synchronize_changed_chart_caches",
            [](Workbook& wb, const ChartCacheSyncOptions& options) { return wb.synchronizeChangedChartCaches(options); },
            py::arg("options") = ChartCacheSyncOptions{})
        .def("add_vba_project", [](Workbook& wb, const std::string& path) { wb.addVbaProject(std::filesystem::path(path)); })
        .def("set_vba_project", [](Workbook& wb, const py::bytes& data) {
            std::string raw = data.cast<std::string>();
            wb.setVbaProject(std::vector<unsigned char>(raw.begin(), raw.end()));
        })
        .def("set_vba_module", &Workbook::setVbaModule)
        .def("set_vba_project_info", &Workbook::setVbaProjectInfo)
        .def("set_vba_designer_storage", &Workbook::setVbaDesignerStorage)
        .def("remove_vba_designer_storage", &Workbook::removeVbaDesignerStorage)
        .def("set_vba_module_text", &Workbook::setVbaModuleText)
        .def("set_vba_class_module_text", &Workbook::setVbaClassModuleText)
        .def("set_vba_document_module_text", &Workbook::setVbaDocumentModuleText)
        .def("remove_vba_module", &Workbook::removeVbaModule)
        .def("apply_chart_color_style",
            [](Workbook& wb, const std::string& sheetName, const std::string& chartStableId,
               bool applyFill, bool applyLine, bool applyMarker) {
                return wb.applyChartColorStyle(sheetName, chartStableId, applyFill, applyLine, applyMarker);
            }, py::arg("worksheet"), py::arg("chart_stable_id"),
               py::arg("apply_fill") = true, py::arg("apply_line") = true, py::arg("apply_marker") = true)
        .def("apply_chart_theme_style_matrix",
            [](Workbook& wb, const std::string& sheetName, const std::string& chartStableId,
               std::size_t fillStyleIndex, std::size_t lineStyleIndex, bool applyMarker) {
                return wb.applyChartThemeStyleMatrix(sheetName, chartStableId, fillStyleIndex, lineStyleIndex, applyMarker);
            }, py::arg("worksheet"), py::arg("chart_stable_id"),
               py::arg("fill_style_index"), py::arg("line_style_index"), py::arg("apply_marker") = true)
        .def_property_readonly("preserved_parts", [](Workbook& wb) -> std::vector<PreservedPart>& { return wb.preservedParts(); }, py::return_value_policy::reference_internal)
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

    m.attr("__version__") = "1.1.2";

    // Excel 365 dynamic-array function prefix helper
    m.def("xlfn", [](const std::string& f) { return xlpp::xlfn(f); },
          py::arg("function"),
          "Prefix an Excel 365 function name with _xlfn. (e.g. SORT -> _xlfn.SORT)");
}
