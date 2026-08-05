#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <pybind11/stl/filesystem.h>
#include <pybind11/functional.h>
#include <pybind11/numpy.h>
#include <XLPP/XLPP.h>
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
    // Fast path: str is most common
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

    py::class_<SaveOptions>(m, "SaveOptions")
        .def(py::init<>())
        .def_readwrite("compression_level", &SaveOptions::compressionLevel)
        .def_readwrite("parallel_workers", &SaveOptions::parallelWorkers);

    py::class_<LoadOptions>(m, "LoadOptions")
        .def(py::init<>())
        .def_readwrite("lenient", &LoadOptions::lenient);

    // === CellReference ===
    py::class_<CellReference>(m, "CellReference")
        .def(py::init<>())
        .def(py::init<std::size_t, std::size_t>(), py::arg("row"), py::arg("column"))
        .def_readwrite("row", &CellReference::row)
        .def_readwrite("column", &CellReference::column)
        .def_static("parse", &CellReference::parse, py::arg("address"))
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
        .def_property_readonly("second_int", [](const DateTime& d) { return static_cast<int>(d.second); })
        .def_property_readonly("millisecond", [](const DateTime& d) {
            return static_cast<int>((d.second - static_cast<int>(d.second)) * 1000.0);
        })
        .def("__str__", [](const DateTime& d) { return toIso8601(d); })
        .def("__repr__", [](const DateTime& d) { return "<DateTime " + toIso8601(d) + ">"; });

    // === Color ===
    py::class_<Color>(m, "Color")
        .def(py::init<>())
        .def("set_argb", &Color::setArgb, py::return_value_policy::reference_internal)
        .def_property("argb", &Color::argb, &Color::setArgb);

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
        .def("bottom", static_cast<BorderSide& (Border::*)()>(&Border::bottom), py::return_value_policy::reference_internal);
    // === Alignment ===
    py::class_<Alignment>(m, "Alignment")
        .def(py::init<>())
        .def_property("horizontal", &Alignment::horizontal,
            [](Alignment& a, std::string v) -> Alignment& { a.setHorizontal(std::move(v)); return a; })
        .def_property("vertical", &Alignment::vertical,
            [](Alignment& a, std::string v) -> Alignment& { a.setVertical(std::move(v)); return a; })
        .def_property("wrap_text", &Alignment::wrapText,
            [](Alignment& a, bool v) -> Alignment& { a.setWrapText(v); return a; });

    // === Style ===
    py::class_<Style>(m, "Style")
        .def(py::init<>())
        .def("font", static_cast<Font& (Style::*)()>(&Style::font), py::return_value_policy::reference_internal)
        .def("fill", static_cast<Fill& (Style::*)()>(&Style::fill), py::return_value_policy::reference_internal)
        .def("border", static_cast<Border& (Style::*)()>(&Style::border), py::return_value_policy::reference_internal)
        .def("alignment", static_cast<Alignment& (Style::*)()>(&Style::alignment), py::return_value_policy::reference_internal)
        .def_property("number_format", &Style::numberFormat, &Style::setNumberFormat);

    // === Hyperlink ===
    py::class_<Hyperlink>(m, "Hyperlink")
        .def(py::init<>())
        .def(py::init<std::string>(), py::arg("target"))
        .def_property("target", &Hyperlink::target, &Hyperlink::setTarget)
        .def_property("display", &Hyperlink::display, &Hyperlink::setDisplay)
        .def_property("tooltip", &Hyperlink::tooltip, &Hyperlink::setTooltip);

    // === Comment ===
    py::class_<Comment>(m, "Comment")
        .def(py::init<>())
        .def(py::init<std::string, std::string>(), py::arg("text"), py::arg("author") = "XL++")
        .def_property("text", &Comment::text, &Comment::setText)
        .def_property("author", &Comment::author, &Comment::setAuthor);

    // === DocumentProperties ===
    py::class_<DocumentProperties>(m, "DocumentProperties")
        .def(py::init<>())
        .def_property("title", &DocumentProperties::title, &DocumentProperties::setTitle)
        .def_property("subject", &DocumentProperties::subject, &DocumentProperties::setSubject)
        .def_property("creator", &DocumentProperties::creator, &DocumentProperties::setCreator)
        .def_property("description", &DocumentProperties::description, &DocumentProperties::setDescription)
        .def_property("keywords", &DocumentProperties::keywords, &DocumentProperties::setKeywords);

    // === AutoFilter ===
    py::class_<AutoFilter>(m, "AutoFilter")
        .def("set_reference", &AutoFilter::setReference)
        .def_property("reference", &AutoFilter::reference, &AutoFilter::setReference)
        .def("column", &AutoFilter::column, py::return_value_policy::reference_internal)
        .def("sort_state", [](AutoFilter& f) -> SortState& { return f.sortState(); },
             py::return_value_policy::reference_internal);

    py::class_<FilterColumn>(m, "FilterColumn")
        .def("add_value", &FilterColumn::addValue)
        .def_property_readonly("values", &FilterColumn::values);

    py::class_<SortState>(m, "SortState")
        .def("set_reference", &SortState::setReference)
        .def("add_condition", &SortState::addCondition, py::arg("ref"), py::arg("descending") = false);

    // === PageSetup ===
    py::enum_<PageOrientation>(m, "PageOrientation")
        .value("DEFAULT", PageOrientation::Default)
        .value("PORTRAIT", PageOrientation::Portrait)
        .value("LANDSCAPE", PageOrientation::Landscape);

    py::class_<PageSetup>(m, "PageSetup")
        .def(py::init<>())
        .def_property("orientation", &PageSetup::orientation, &PageSetup::setOrientation)
        .def_property("scale", &PageSetup::scale, &PageSetup::setScale);

    py::class_<PageMargins>(m, "PageMargins")
        .def(py::init<>())
        .def_property("left", &PageMargins::left, &PageMargins::setLeft)
        .def_property("right", &PageMargins::right, &PageMargins::setRight)
        .def_property("top", &PageMargins::top, &PageMargins::setTop)
        .def_property("bottom", &PageMargins::bottom, &PageMargins::setBottom);

    // === Image ===
    py::class_<Image>(m, "Image")
        .def_static("from_file", &Image::fromFile);

    // === Table ===
    py::class_<TableColumn>(m, "TableColumn")
        .def_property_readonly("id", &TableColumn::id)
        .def_property("name", &TableColumn::name, &TableColumn::setName);

    py::class_<Table>(m, "Table")
        .def(py::init<>())
        .def(py::init<std::string, std::string>())
        .def_property_readonly("name", &Table::name)
        .def_property_readonly("display_name", &Table::displayName)
        .def_property("reference", &Table::reference, &Table::setReference)
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
        .def_property_readonly("value", &DefinedName::value);

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

    py::class_<ConditionalRule>(m, "ConditionalRule")
        .def_static("formula", &ConditionalRule::formula, py::arg("expression"))
        .def_static("cell_is", &ConditionalRule::cellIs, py::arg("op"), py::arg("value"))
        .def_static("cell_is_between", &ConditionalRule::cellIsBetween,
                    py::arg("lower"), py::arg("upper"), py::arg("negate") = false)
        .def_static("data_bar", &ConditionalRule::dataBar, py::arg("color") = "FF638EC6")
        .def_static("color_scale", &ConditionalRule::colorScale)
        .def_static("icon_set", &ConditionalRule::iconSet, py::arg("icons") = "3Arrows")
        .def_property_readonly("type", &ConditionalRule::type)
        .def_property_readonly("op", &ConditionalRule::op)
        .def_property_readonly("formulas", &ConditionalRule::formulas)
        .def_property("priority", &ConditionalRule::priority, &ConditionalRule::setPriority)
        .def_property("stop_if_true", &ConditionalRule::stopIfTrue, &ConditionalRule::setStopIfTrue)
        .def("__repr__", [](const ConditionalRule& r) { return "<ConditionalRule>"; });

    py::class_<ConditionalFormattingEntry>(m, "ConditionalFormattingEntry")
        .def_property_readonly("reference", &ConditionalFormattingEntry::reference)
        .def_property_readonly("rules", [](ConditionalFormattingEntry& e) -> py::list {
            py::list out;
            for (auto& r : e.rules()) out.append(py::cast(&r, py::return_value_policy::reference));
            return out;
        });

    py::class_<ConditionalFormattingCollection>(m, "ConditionalFormattingCollection")
        .def("add_rule", [](ConditionalFormattingCollection& c, const std::string& ref,
                            ConditionalRule rule) -> ConditionalRule& {
                return c.addRule(ref, std::move(rule));
            }, py::arg("reference"), py::arg("rule"),
            py::return_value_policy::reference_internal)
        .def_property_readonly("entries", [](ConditionalFormattingCollection& c) -> py::list {
            py::list out;
            for (auto& e : c.entries()) out.append(py::cast(&e, py::return_value_policy::reference));
            return out;
        });

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

    py::class_<DataValidation>(m, "DataValidation")
        .def(py::init<>())
        .def(py::init<DataValidationType>(), py::arg("type"))
        .def_property_readonly("type", &DataValidation::type)
        .def_property_readonly("op", &DataValidation::op)
        .def_property("formula1", &DataValidation::formula1, &DataValidation::setFormula1)
        .def_property("formula2", &DataValidation::formula2, &DataValidation::setFormula2)
        .def_property("reference", &DataValidation::reference, &DataValidation::setReference)
        .def_property("allow_blank", &DataValidation::allowBlank, &DataValidation::setAllowBlank)
        .def("__repr__", [](const DataValidation& v) { return "<DataValidation>"; });

    py::class_<DataValidationCollection>(m, "DataValidationCollection")
        .def("add", [](DataValidationCollection& c, DataValidationType type,
                       const std::string& reference) -> DataValidation& {
                return c.add(type, reference);
            }, py::arg("type"), py::arg("reference"),
            py::return_value_policy::reference_internal)
        .def("add_validation", [](DataValidationCollection& c, DataValidation v) -> DataValidation& {
                return c.add(std::move(v));
            }, py::arg("validation"), py::return_value_policy::reference_internal)
        .def_property_readonly("items", [](DataValidationCollection& c) -> py::list {
            py::list out;
            for (auto& v : c.items()) out.append(py::cast(&v, py::return_value_policy::reference));
            return out;
        });

    // === Cell ===
    py::class_<Cell>(m, "Cell")
        .def_property_readonly("address", &Cell::address)
        .def_property_readonly("row", &Cell::row)
        .def_property_readonly("column", &Cell::column)
        .def_property("value",
            [](const Cell& c) { return cellvalue_to_py(c.value()); },
            [](Cell& c, const py::object& v) {
                auto cv = py_to_cellvalue(v);
                if (auto* dt = std::get_if<DateTime>(&cv)) {
                    // Applying the date via setDate/setDateTime installs a
                    // matching number format so the value survives save/load
                    // as a real date instead of a bare serial number.
                    if (dt->hour == 0 && dt->minute == 0 && dt->second == 0.0)
                        c.setDate(*dt);
                    else
                        c.setDateTime(*dt);
                } else {
                    c.setValue(cv);
                }
            })
        .def("set_formula", [](Cell& c, const std::string& f) { c.setFormula(f); })
        .def("set_dynamic_array_formula", [](Cell& c, const std::string& f, const std::string& ref) {
            c.setDynamicArrayFormula(f, ref);
        })
        .def_property_readonly("formula", &Cell::formula)
        .def("has_formula", &Cell::hasFormula)
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
        .def("has_hyperlink", &Cell::hasHyperlink)
        .def("hyperlink", &Cell::hyperlink, py::return_value_policy::reference_internal)
        .def("set_hyperlink", &Cell::setHyperlink)
        .def("has_comment", &Cell::hasComment)
        .def("comment", &Cell::comment, py::return_value_policy::reference_internal)
        .def("set_comment", [](Cell& c, const Comment& v) { c.setComment(v); })
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
        .def("range", py::overload_cast<const std::string&>(&Worksheet::range))
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
                // Header row
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
                // Data always starts at row 2, symmetric with from_records(header=True).
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
        .def("set_print_area", &Worksheet::setPrintArea)
        .def_property_readonly("print_area", &Worksheet::printArea)
        .def_property_readonly("max_row", &Worksheet::maxRow)
        .def_property_readonly("max_column", &Worksheet::maxColumn)
        .def("dimensions", &Worksheet::dimensions)
        .def_property_readonly("empty", &Worksheet::empty)
        .def_property_readonly("row_count", &Worksheet::rowCount)
        .def_property_readonly("col_count", &Worksheet::columnCount)
        .def("insert_rows", &Worksheet::insertRows, py::arg("index"), py::arg("amount") = 1)
        .def("delete_rows", &Worksheet::deleteRows, py::arg("index"), py::arg("amount") = 1)
        .def("insert_columns", &Worksheet::insertColumns, py::arg("index"), py::arg("amount") = 1)
        .def("delete_columns", &Worksheet::deleteColumns, py::arg("index"), py::arg("amount") = 1)
        .def("auto_filter", py::overload_cast<>(&Worksheet::autoFilter),
             py::return_value_policy::reference_internal)
        .def("conditional_formatting", py::overload_cast<>(&Worksheet::conditionalFormatting),
             py::return_value_policy::reference_internal)
        .def("data_validations", py::overload_cast<>(&Worksheet::dataValidations),
             py::return_value_policy::reference_internal)
        .def("page_setup", py::overload_cast<>(&Worksheet::pageSetup),
             py::return_value_policy::reference_internal)
        .def("page_margins", py::overload_cast<>(&Worksheet::pageMargins),
             py::return_value_policy::reference_internal)
        .def("protection", py::overload_cast<>(&Worksheet::protection),
             py::return_value_policy::reference_internal)
        .def("add_table", &Worksheet::addTable, py::return_value_policy::reference_internal)
        .def("add_image", [](Worksheet& ws, const std::string& path, const std::string& anchor) -> Image& {
            return ws.addImage(path, anchor);
        }, py::return_value_policy::reference_internal)
        .def("__iter__", [](Worksheet& ws) {
            // Materialize the rows into a Python-owned list: the underlying
            // std::vector<Row> would otherwise dangle as soon as the lambda
            // returns. The returned iterator owns the list; callers must keep
            // the worksheet alive (same lifetime contract as C++).
            auto rows = ws.rows();
            py::list result;
            for (auto& row : rows) result.append(py::cast(row));
            return py::iter(result);
        })
        .def("__repr__", [](const Worksheet& ws) { return "<Worksheet '" + ws.name() + "'>"; });

    // === Row ===
    py::class_<Row>(m, "Row")
        .def_property_readonly("number", &Row::number)
        .def("cell", &Row::cell, py::return_value_policy::reference_internal)
        .def("__iter__", [](Row& r) {
            // Materialize cell pointers; the underlying std::vector<Cell*>
            // would otherwise dangle. Cell objects reference the worksheet's
            // stable storage.
            auto cells = r.cells();
            py::list result;
            for (auto* c : cells)
                result.append(py::cast(c, py::return_value_policy::reference));
            return py::iter(result);
        })
        .def("values", &Row::values);

    // === Workbook ===
    py::class_<Workbook>(m, "Workbook")
        .def(py::init<>())
        .def("add_worksheet", &Workbook::addWorksheet, py::return_value_policy::reference_internal)
        .def("copy_worksheet", &Workbook::copyWorksheet, py::return_value_policy::reference_internal)
        .def("remove_worksheet", &Workbook::removeWorksheet)
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
        .def_property_readonly("active", [](Workbook& wb) -> Worksheet& { return wb[0]; },
                               py::return_value_policy::reference_internal)
        .def("index", &Workbook::index)
        .def("load",
            [](Workbook& wb, const std::string& path) { wb.load(std::filesystem::path(path)); })
        .def("load",
            [](Workbook& wb, const std::string& path, const LoadOptions& opts) {
                wb.load(std::filesystem::path(path), opts);
            }, py::arg("path"), py::arg("options"))
        .def("save",
            [](const Workbook& wb, const std::string& path, const SaveOptions* opts) {
                if (opts) wb.save(std::filesystem::path(path), *opts);
                else wb.save(std::filesystem::path(path));
            }, py::arg("path"), py::arg("options") = nullptr)
        .def_property_readonly("properties", py::overload_cast<>(&Workbook::properties))
        .def("add_named_style", &Workbook::addNamedStyle, py::return_value_policy::reference_internal)
        .def("named_style", py::overload_cast<const std::string&>(&Workbook::namedStyle, py::const_),
             py::return_value_policy::reference_internal)
        .def("apply_named_style", &Workbook::applyNamedStyle)
        .def("add_defined_name", &Workbook::addDefinedName, py::return_value_policy::reference_internal)
        .def_property_readonly("defined_names", [](Workbook& wb) -> py::list {
            py::list out;
            for (auto& n : wb.definedNames())
                out.append(py::cast(&n, py::return_value_policy::reference));
            return out;
        })
        .def_property("date_1904", &Workbook::date1904, &Workbook::setDate1904)
        .def("clear", &Workbook::clear)
        .def("__iter__", [](Workbook& wb) {
            // Yield references into the workbook's stable worksheet storage.
            // The returned iterator owns the materialized list; callers must
            // keep the workbook alive (same lifetime contract as C++).
            py::list result;
            for (auto& ws : wb.worksheets())
                result.append(py::cast(&ws, py::return_value_policy::reference));
            return py::iter(result);
        })
        .def("__repr__", [](const Workbook& wb) {
            return "<Workbook sheets=" + std::to_string(wb.sheetCount()) + ">";
        });

    m.attr("__version__") = "1.1.1";

    // Excel 365 dynamic-array function prefix helper
    m.def("xlfn", [](const std::string& f) { return xlpp::xlfn(f); },
          py::arg("function"),
          "Prefix an Excel 365 function name with _xlfn. (e.g. SORT -> _xlfn.SORT)");
}
