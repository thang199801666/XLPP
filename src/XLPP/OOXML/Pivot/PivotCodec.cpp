#include "OOXML/Pivot/PivotCodec.h"

#include "OOXML/Common/Namespaces.h"
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include "OOXML/Common/PackageRelationships.h"

#include <XLPP/Cell/CellReference.h>
#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Cell/DateTime.h>

#include <algorithm>
#include <iomanip>
#include <functional>
#include <set>
#include <sstream>
#include <stdexcept>
#include <utility>
#include <vector>

namespace xlpp::internal::ooxml {

using xlpp::internal::xmlEscape;

int resolvedPivotFieldIndex(const xlpp::PivotCache& cache, int index, const std::string& name) {
    if (index >= 0 && static_cast<std::size_t>(index) < cache.fields().size()) return index;
    if (!name.empty()) {
        const auto resolved = cache.fieldIndex(name);
        if (resolved >= 0) return resolved;
    }
    throw std::invalid_argument("Pivot field '" + name + "' is not present in the cache fields");
}


const xlpp::PivotField* pivotAxisFieldByCacheIndex(const xlpp::PivotTable& pt, int cacheIndex) {
    const auto findField = [cacheIndex, &pt](const std::vector<xlpp::PivotField>& fields) -> const xlpp::PivotField* {
        const auto it = std::find_if(fields.begin(), fields.end(), [&](const auto& field) {
            try { return resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()) == cacheIndex; }
            catch (...) { return false; }
        });
        return it == fields.end() ? nullptr : &*it;
    };
    if (const auto* value = findField(pt.rowFields())) return value;
    if (const auto* value = findField(pt.columnFields())) return value;
    return findField(pt.pageFields());
}

std::vector<std::vector<std::size_t>> pivotAxisCombinations(
    const xlpp::PivotTable& pt,
    const std::vector<int>& axisIndexes,
    const std::vector<std::vector<std::string>>& sharedItems) {
    std::vector<std::vector<std::size_t>> combinations{{}};
    for (const auto cacheIndex : axisIndexes) {
        const auto* model = pivotAxisFieldByCacheIndex(pt, cacheIndex);
        std::vector<std::size_t> visible;
        const auto count = sharedItems[static_cast<std::size_t>(cacheIndex)].size();
        visible.reserve(count);
        for (std::size_t item = 0; item < count; ++item) {
            if (!model || !model->itemHidden(static_cast<int>(item))) visible.push_back(item);
        }
        // Empty caches can occur when saveData=false. Keep a synthetic zero slot so
        // the view remains structurally refreshable without inventing cache data.
        if (visible.empty()) visible.push_back(0);
        std::vector<std::vector<std::size_t>> next;
        for (const auto& prefix : combinations) for (const auto item : visible) {
            auto row = prefix; row.push_back(item); next.push_back(std::move(row));
        }
        combinations = std::move(next);
    }
    return combinations;
}

void writePivotAxisItems(std::ostringstream& xml, const char* tag,
                         const std::vector<std::vector<std::size_t>>& combinations,
                         std::size_t axisDepth, bool includeGrandTotal) {
    const auto count = combinations.size() + (includeGrandTotal ? 1u : 0u);
    xml << '<' << tag << " count=\"" << count << "\">";
    for (const auto& tuple : combinations) {
        xml << "<i>";
        for (const auto item : tuple) xml << "<x v=\"" << item << "\"/>";
        xml << "</i>";
    }
    if (includeGrandTotal) {
        xml << "<i t=\"grand\">";
        for (std::size_t i = 0; i < axisDepth; ++i) xml << "<x/>";
        xml << "</i>";
    }
    xml << "</" << tag << '>';
}

bool pivotShowAsRequiresX14(const std::string& value) {
    static const std::vector<std::string> extended{
        "percentOfParent", "percentOfParentRow", "percentOfParentCol",
        "percentOfRunningTotal", "rankAscending", "rankDescending"
    };
    return std::find(extended.begin(), extended.end(), value) != extended.end();
}

std::string pivotTableXml(const xlpp::PivotTable& pt, std::size_t id, bool strict) {
    std::ostringstream xml;
    const auto fieldCount = pt.cache().fields().size();
    if (pt.name().empty()) throw std::invalid_argument("Pivot table name cannot be empty");
    if (fieldCount == 0) throw std::invalid_argument("Pivot table cache must contain at least one field");

    std::vector<int> rowIndexes, columnIndexes, pageIndexes, dataIndexes;
    rowIndexes.reserve(pt.rowFields().size());
    columnIndexes.reserve(pt.columnFields().size());
    pageIndexes.reserve(pt.pageFields().size());
    dataIndexes.reserve(pt.dataFields().size());
    for (const auto& field : pt.rowFields()) rowIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.columnFields()) columnIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.pageFields()) pageIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));
    for (const auto& field : pt.dataFields()) dataIndexes.push_back(resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()));

    std::vector<std::vector<std::string>> sharedItems(fieldCount);
    for (const auto& record : pt.cache().records()) {
        if (record.size() != fieldCount)
            throw std::invalid_argument("Pivot cache record width must match field count");
        for (std::size_t i = 0; i < fieldCount; ++i) {
            if (std::find(sharedItems[i].begin(), sharedItems[i].end(), record[i]) == sharedItems[i].end())
                sharedItems[i].push_back(record[i]);
        }
    }

    // A PivotTable location is the complete occupied view range, not merely the
    // requested top-left cell.  When the caller supplies an anchor, derive the
    // smallest view rectangle from the emitted row-item matrix.  This keeps the
    // location, rowItems and colItems mutually consistent for desktop Excel.
    std::string location = pt.location().empty() ? "D2" : pt.location();
    if (location.find(':') == std::string::npos) {
        const auto first = xlpp::CellReference::parse(location);
        const auto rowTuples = rowIndexes.empty() ? std::size_t{1} : pivotAxisCombinations(pt, rowIndexes, sharedItems).size();
        const auto rowItemCount = rowIndexes.empty() ? std::size_t{1} : rowTuples + (pt.rowGrandTotals() ? 1u : 0u);
        auto columnTuples = columnIndexes.empty() ? std::size_t{1} : pivotAxisCombinations(pt, columnIndexes, sharedItems).size();
        if (dataIndexes.size() > 1) columnTuples *= dataIndexes.size();
        const auto outputRows = std::max<std::size_t>(2, rowItemCount + 1); // field/header row
        const auto rowHeaderWidth = pt.layout() == xlpp::PivotLayout::Compact ? std::size_t{1} : std::max<std::size_t>(1, rowIndexes.size());
        const auto outputColumns = std::max<std::size_t>(2, rowHeaderWidth + std::max<std::size_t>(1, columnTuples));
        xlpp::CellReference last{first.row + outputRows - 1, first.column + outputColumns - 1};
        location += ":" + last.address();
    }

    // Match the conservative view metadata emitted by desktop Excel.  The
    // extension list and revision UID are optional, but the layout/version flags
    // below make the intended non-OLAP compact view unambiguous.
    const bool needsX14 = std::any_of(pt.rowFields().begin(), pt.rowFields().end(), [](const auto& field){ return field.repeatItemLabels(); })
        || std::any_of(pt.columnFields().begin(), pt.columnFields().end(), [](const auto& field){ return field.repeatItemLabels(); })
        || std::any_of(pt.dataFields().begin(), pt.dataFields().end(), [](const auto& field){ return pivotShowAsRequiresX14(field.showDataAs()); });
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotTableDefinition xmlns=\""
        << nsMain(strict) << "\"";
    if (needsX14) xml << " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"";
    xml << " name=\"" << xmlEscape(pt.name()) << "\" cacheId=\"" << id
        << "\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\""
        << " applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\""
        << " dataCaption=\"" << xmlEscape(pt.dataCaption()) << "\" updatedVersion=\"8\" minRefreshableVersion=\"3\""
        << " useAutoFormatting=\"" << (pt.useAutoFormatting() ? 1 : 0) << "\" preserveFormatting=\"" << (pt.preserveFormatting() ? 1 : 0) << "\""
        << " rowGrandTotals=\"" << (pt.rowGrandTotals() ? 1 : 0) << "\" colGrandTotals=\"" << (pt.columnGrandTotals() ? 1 : 0) << "\""
        << " itemPrintTitles=\"1\" createdVersion=\"8\" indent=\"0\" multipleFieldFilters=\"" << (pt.multipleFieldFilters() ? 1 : 0) << "\""
        << " showEmptyRow=\"" << (pt.showEmptyRow() ? 1 : 0) << "\" showEmptyCol=\"" << (pt.showEmptyColumn() ? 1 : 0) << "\""
        << " showDrill=\"" << (pt.showDrill() ? 1 : 0) << "\" enableDrill=\"" << (pt.enableDrill() ? 1 : 0) << "\""
        << " showDataTips=\"" << (pt.showDataTips() ? 1 : 0) << "\" showMemberPropertyTips=\"" << (pt.showMemberPropertyTips() ? 1 : 0) << "\""
        << " showHeaders=\"" << (pt.showHeaders() ? 1 : 0) << "\" showValuesRow=\"" << (pt.showValuesRow() ? 1 : 0) << "\""
        << " subtotalHiddenItems=\"" << (pt.subtotalHiddenItems() ? 1 : 0) << "\" pageWrap=\"" << pt.pageWrap() << "\""
        << " pageOverThenDown=\"" << (pt.pageOverThenDown() ? 1 : 0) << "\"";
    switch (pt.layout()) {
        case xlpp::PivotLayout::Compact: xml << " compact=\"1\" compactData=\"1\" outline=\"1\" outlineData=\"1\""; break;
        case xlpp::PivotLayout::Outline: xml << " compact=\"0\" compactData=\"0\" outline=\"1\" outlineData=\"1\""; break;
        case xlpp::PivotLayout::Tabular: xml << " compact=\"0\" compactData=\"0\" outline=\"0\" outlineData=\"0\""; break;
    }
    xml << ">";
    xml << "<location ref=\"" << xmlEscape(location) << "\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\"/>";

    xml << "<pivotFields count=\"" << fieldCount << "\">";
    for (std::size_t fieldIndex = 0; fieldIndex < fieldCount; ++fieldIndex) {
        const auto contains = [fieldIndex](const std::vector<int>& values) {
            return std::find(values.begin(), values.end(), static_cast<int>(fieldIndex)) != values.end();
        };
        const auto isRow = contains(rowIndexes);
        const auto isCol = contains(columnIndexes);
        const auto isPage = contains(pageIndexes);
        const auto isData = contains(dataIndexes);
        const xlpp::PivotField* modelField = nullptr;
        if (isRow) modelField = &pt.rowFields()[static_cast<std::size_t>(std::distance(rowIndexes.begin(), std::find(rowIndexes.begin(), rowIndexes.end(), static_cast<int>(fieldIndex))))];
        else if (isCol) modelField = &pt.columnFields()[static_cast<std::size_t>(std::distance(columnIndexes.begin(), std::find(columnIndexes.begin(), columnIndexes.end(), static_cast<int>(fieldIndex))))];
        else if (isPage) modelField = &pt.pageFields()[static_cast<std::size_t>(std::distance(pageIndexes.begin(), std::find(pageIndexes.begin(), pageIndexes.end(), static_cast<int>(fieldIndex))))];

        xml << "<pivotField";
        if (isRow) xml << " axis=\"axisRow\"";
        else if (isCol) xml << " axis=\"axisCol\"";
        else if (isPage) xml << " axis=\"axisPage\"";
        if (isData) xml << " dataField=\"1\"";
        xml << " showAll=\"" << ((modelField && modelField->showAll()) ? 1 : 0) << "\"";
        if (modelField) {
            if (modelField->sortType() == 1) xml << " sortType=\"ascending\"";
            else if (modelField->sortType() == 2) xml << " sortType=\"descending\"";
            xml << " subtotalTop=\"" << (modelField->subtotalTop() ? 1 : 0) << "\""
                << " insertBlankRow=\"" << (modelField->insertBlankRow() ? 1 : 0) << "\""
                << " includeNewItemsInFilter=\"" << (modelField->includeNewItemsInFilter() ? 1 : 0) << "\""
                << " multipleItemSelectionAllowed=\"" << (modelField->multipleItemSelectionAllowed() ? 1 : 0) << "\""
                << " compact=\"" << (modelField->compact() ? 1 : 0) << "\" outline=\"" << (modelField->outline() ? 1 : 0) << "\""
                << " insertPageBreak=\"" << (modelField->insertPageBreak() ? 1 : 0) << "\" showDropDowns=\"" << (modelField->showDropDowns() ? 1 : 0) << "\"";
            for (const auto& subtotal : modelField->subtotals()) xml << ' ' << subtotal << "Subtotal=\"1\"";
        }
        if (isRow || isCol || isPage) xml << " defaultSubtotal=\"" << ((modelField && modelField->defaultSubtotal()) ? 1 : 0) << "\"";
        if (isRow || isCol || isPage) {
            // The axis item list contains only concrete cache items.  A synthetic
            // default item is not needed for an unfiltered field and can make the
            // row/column view indexes ambiguous to Excel.
            xml << "><items count=\"" << sharedItems[fieldIndex].size() << "\">";
            for (std::size_t itemIndex = 0; itemIndex < sharedItems[fieldIndex].size(); ++itemIndex) {
                xml << "<item x=\"" << itemIndex << "\"";
                if (modelField && modelField->itemHidden(static_cast<int>(itemIndex))) xml << " h=\"1\"";
                xml << "/>";
            }
            xml << "</items>";
            if (modelField && modelField->repeatItemLabels())
                xml << "<extLst><ext uri=\"{2946ED86-A175-432a-8AC1-64E0C546D7DE}\"><x14:pivotField fillDownLabels=\"1\"/></ext></extLst>";
            xml << "</pivotField>";
        } else {
            xml << "/>";
        }
    }
    xml << "</pivotFields>";

    if (!rowIndexes.empty()) {
        xml << "<rowFields count=\"" << rowIndexes.size() << "\">";
        for (const auto index : rowIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</rowFields>";
        writePivotAxisItems(xml, "rowItems", pivotAxisCombinations(pt, rowIndexes, sharedItems), rowIndexes.size(), pt.rowGrandTotals());
    }

    std::vector<int> effectiveColumnIndexes = columnIndexes;
    const bool valuesOnColumns = dataIndexes.size() > 1;
    if (valuesOnColumns) effectiveColumnIndexes.push_back(-2); // Excel's synthetic Values field.
    if (!effectiveColumnIndexes.empty()) {
        xml << "<colFields count=\"" << effectiveColumnIndexes.size() << "\">";
        for (const auto index : effectiveColumnIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</colFields>";
    }
    if (columnIndexes.empty() && !valuesOnColumns) {
        xml << "<colItems count=\"1\"><i/></colItems>";
    } else {
        auto combinations = columnIndexes.empty() ? std::vector<std::vector<std::size_t>>{{}} : pivotAxisCombinations(pt, columnIndexes, sharedItems);
        if (valuesOnColumns) {
            std::vector<std::vector<std::size_t>> expanded;
            for (const auto& tuple : combinations) for (std::size_t dataIndex = 0; dataIndex < dataIndexes.size(); ++dataIndex) {
                auto value = tuple; value.push_back(dataIndex); expanded.push_back(std::move(value));
            }
            combinations = std::move(expanded);
        }
        writePivotAxisItems(xml, "colItems", combinations, effectiveColumnIndexes.size(), pt.columnGrandTotals());
    }
    if (!pageIndexes.empty()) {
        xml << "<pageFields count=\"" << pageIndexes.size() << "\">";
        for (std::size_t pageIndex = 0; pageIndex < pageIndexes.size(); ++pageIndex) {
            const auto fieldIndex = pageIndexes[pageIndex];
            const auto& field = pt.pageFields()[pageIndex];
            xml << "<pageField fld=\"" << fieldIndex << "\" hier=\"-1\"";
            if (field.selectedItemIndex() >= 0) {
                const auto selected = static_cast<std::size_t>(field.selectedItemIndex());
                const auto& items = sharedItems[static_cast<std::size_t>(fieldIndex)];
                if (selected >= items.size())
                    throw std::invalid_argument("Pivot page-field selected item is outside the cache item range");
                xml << " item=\"" << selected << "\"";
            }
            xml << "/>";
        }
        xml << "</pageFields>";
    }
    if (!dataIndexes.empty()) {
        xml << "<dataFields count=\"" << dataIndexes.size() << "\">";
        for (std::size_t i = 0; i < dataIndexes.size(); ++i) {
            const auto index = dataIndexes[i];
            const auto& field = pt.dataFields()[i];
            const auto& cacheName = pt.cache().fields()[static_cast<std::size_t>(index)];
            const auto sourceName = field.name().empty() ? cacheName : field.name();
            const auto subtotal = field.subtotal().empty() ? std::string("sum") : field.subtotal();
            const auto displayName = field.caption().empty() ? (subtotal == "sum" ? "Sum of " + sourceName : sourceName) : field.caption();
            xml << "<dataField name=\"" << xmlEscape(displayName) << "\" fld=\"" << index
                << "\" baseField=\"" << field.baseField() << "\" baseItem=\"" << field.baseItem() << "\"";
            if (subtotal != "sum") xml << " subtotal=\"" << xmlEscape(subtotal) << "\"";
            if (field.numberFormatId() != 0) xml << " numFmtId=\"" << field.numberFormatId() << "\"";
            const bool extendedShowAs = pivotShowAsRequiresX14(field.showDataAs());
            if (!extendedShowAs && !field.showDataAs().empty() && field.showDataAs() != "normal")
                xml << " showDataAs=\"" << xmlEscape(field.showDataAs()) << "\"";
            if (extendedShowAs) {
                xml << "><extLst><ext uri=\"{E15A36E0-9728-4E99-A89B-3F7291B0FE68}\"><x14:dataField pivotShowAs=\""
                    << xmlEscape(field.showDataAs()) << "\"/></ext></extLst></dataField>";
            } else {
                xml << "/>";
            }
        }
        xml << "</dataFields>";
    }
    if (!pt.filters().empty()) {
        xml << "<filters count=\"" << pt.filters().size() << "\">";
        std::size_t filterId = 1;
        for (const auto& filter : pt.filters()) {
            if (filter.fieldIndex < 0 || static_cast<std::size_t>(filter.fieldIndex) >= fieldCount) throw std::invalid_argument("Pivot filter field index is outside the cache");
            xml << "<filter fld=\"" << filter.fieldIndex << "\" type=\"" << xmlEscape(filter.type) << "\" id=\"" << filterId++ << "\"";
            if (filter.measureFieldIndex >= 0) xml << " iMeasureFld=\"" << filter.measureFieldIndex << "\"";
            if (!filter.value1.empty()) xml << " stringValue1=\"" << xmlEscape(filter.value1) << "\"";
            if (!filter.value2.empty()) xml << " stringValue2=\"" << xmlEscape(filter.value2) << "\"";
            xml << ">";
            if (filter.type == "count" || filter.type == "percent" || filter.type == "sum")
                xml << "<autoFilter ref=\"A1\"><filterColumn colId=\"0\"><top10 top=\"" << (filter.top10Top ? 1 : 0) << "\" percent=\"" << (filter.top10Percent ? 1 : 0) << "\" val=\"" << filter.top10Value << "\"/></filterColumn></autoFilter>";
            xml << "</filter>";
        }
        xml << "</filters>";
    }
    xml << "<pivotTableStyleInfo name=\"" << xmlEscape(pt.styleName())
        << "\" showRowHeaders=\"" << (pt.showRowHeaders() ? 1 : 0)
        << "\" showColHeaders=\"" << (pt.showColumnHeaders() ? 1 : 0)
        << "\" showRowStripes=\"" << (pt.showRowStripes() ? 1 : 0)
        << "\" showColStripes=\"" << (pt.showColumnStripes() ? 1 : 0)
        << "\" showLastColumn=\"" << (pt.showLastColumn() ? 1 : 0) << "\"/>";
    xml << "</pivotTableDefinition>";
    return xml.str();
}

enum class PivotValueKind { Blank, Number, Boolean, Error, String };

struct ParsedPivotValue {
    PivotValueKind kind{PivotValueKind::String};
    double number{0.0};
};

ParsedPivotValue parsePivotValue(const std::string& value) {
    if (value.empty()) return {PivotValueKind::Blank, 0.0};
    if (value == "true") return {PivotValueKind::Boolean, 1.0};
    if (value == "false") return {PivotValueKind::Boolean, 0.0};
    if (value.front() == '#') return {PivotValueKind::Error, 0.0};
    char* end = nullptr;
    const auto number = std::strtod(value.c_str(), &end);
    if (end && end != value.c_str() && *end == '\0' && std::isfinite(number))
        return {PivotValueKind::Number, number};
    return {PivotValueKind::String, 0.0};
}

std::vector<std::vector<std::string>> pivotSharedItems(const xlpp::PivotCache& cache) {
    std::vector<std::vector<std::string>> result(cache.fields().size());
    for (const auto& record : cache.records()) {
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        for (std::size_t i = 0; i < record.size(); ++i) {
            auto& items = result[i];
            if (std::find(items.begin(), items.end(), record[i]) == items.end())
                items.push_back(record[i]);
        }
    }
    return result;
}

bool pivotFieldIsPureData(const xlpp::PivotTable& pt, std::size_t fieldIndex) {
    const auto matchesPivotField = [&](const auto& fields) {
        return std::any_of(fields.begin(), fields.end(), [&](const auto& field) {
            return resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()) == static_cast<int>(fieldIndex);
        });
    };
    const bool onAxis = matchesPivotField(pt.rowFields()) || matchesPivotField(pt.columnFields())
        || matchesPivotField(pt.pageFields());
    const bool isData = std::any_of(pt.dataFields().begin(), pt.dataFields().end(), [&](const auto& field) {
        return resolvedPivotFieldIndex(pt.cache(), field.fieldIndex(), field.name()) == static_cast<int>(fieldIndex);
    });
    return isData && !onAxis;
}

void writePivotValue(std::ostringstream& xml, const std::string& value) {
    const auto parsed = parsePivotValue(value);
    switch (parsed.kind) {
        case PivotValueKind::Blank: xml << "<m/>"; break;
        case PivotValueKind::Number: xml << "<n v=\"" << std::setprecision(15) << parsed.number << "\"/>"; break;
        case PivotValueKind::Boolean: xml << "<b v=\"" << (parsed.number != 0.0 ? 1 : 0) << "\"/>"; break;
        case PivotValueKind::Error: xml << "<e v=\"" << xmlEscape(value) << "\"/>"; break;
        case PivotValueKind::String: xml << "<s v=\"" << xmlEscape(value) << "\"/>"; break;
    }
}

std::string pivotCacheXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    const auto& cache = pt.cache();
    const auto fieldCount = cache.fields().size();
    const auto sharedItems = pivotSharedItems(cache);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheDefinition xmlns=\""
        << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict)
        << "\" r:id=\"rId1\" saveData=\"" << (cache.saveData() ? 1 : 0)
        << "\" refreshOnLoad=\"" << (cache.refreshOnLoad() ? 1 : 0)
        << "\" enableRefresh=\"" << (cache.enableRefresh() ? 1 : 0) << "\""
        << " missingItemsLimit=\"" << cache.missingItemsLimit() << "\" backgroundQuery=\"" << (cache.backgroundQuery() ? 1 : 0) << "\""
        << " optimizeMemory=\"" << (cache.optimizeMemory() ? 1 : 0) << "\" upgradeOnRefresh=\"" << (cache.upgradeOnRefresh() ? 1 : 0) << "\""
        << " supportSubquery=\"" << (cache.supportSubquery() ? 1 : 0) << "\" supportAdvancedDrill=\"" << (cache.supportAdvancedDrill() ? 1 : 0) << "\""
        << (cache.refreshedBy().empty() ? std::string{} : std::string(" refreshedBy=\"") + xmlEscape(cache.refreshedBy()) + "\"")
        << " createdVersion=\"3\" refreshedVersion=\"3\" minRefreshableVersion=\"3\" recordCount=\""
        << cache.records().size() << "\">";
    std::string sourceSheet;
    std::string sourceRef = cache.sourceData();
    const auto bang = sourceRef.find('!');
    if (bang != std::string::npos) {
        sourceSheet = sourceRef.substr(0, bang);
        sourceRef = sourceRef.substr(bang + 1);
        if (sourceSheet.size() >= 2 && sourceSheet.front() == '\'' && sourceSheet.back() == '\'') {
            sourceSheet = sourceSheet.substr(1, sourceSheet.size() - 2);
            std::size_t pos = 0;
            while ((pos = sourceSheet.find("''", pos)) != std::string::npos) sourceSheet.replace(pos, 2, "'");
        }
    }
    sourceRef.erase(std::remove(sourceRef.begin(), sourceRef.end(), '$'), sourceRef.end());
    xml << "<cacheSource type=\"worksheet\"><worksheetSource ref=\"" << xmlEscape(sourceRef) << "\"";
    if (!sourceSheet.empty()) xml << " sheet=\"" << xmlEscape(sourceSheet) << "\"";
    xml << "/></cacheSource><cacheFields count=\"" << fieldCount << "\">";

    for (std::size_t i = 0; i < fieldCount; ++i) {
        const auto& items = sharedItems[i];
        bool containsString = false, containsNumber = false, containsBoolean = false;
        bool containsError = false, containsBlank = false, allNumbersInteger = true;
        double minValue = 0.0, maxValue = 0.0;
        bool hasNumericRange = false;
        std::set<PivotValueKind> nonBlankKinds;
        for (const auto& value : items) {
            const auto parsed = parsePivotValue(value);
            if (parsed.kind != PivotValueKind::Blank) nonBlankKinds.insert(parsed.kind);
            switch (parsed.kind) {
                case PivotValueKind::Blank: containsBlank = true; break;
                case PivotValueKind::Number:
                    containsNumber = true;
                    allNumbersInteger = allNumbersInteger && std::floor(parsed.number) == parsed.number;
                    if (!hasNumericRange) { minValue = maxValue = parsed.number; hasNumericRange = true; }
                    else { minValue = std::min(minValue, parsed.number); maxValue = std::max(maxValue, parsed.number); }
                    break;
                case PivotValueKind::Boolean: containsBoolean = true; break;
                case PivotValueKind::Error: containsError = true; break;
                case PivotValueKind::String: containsString = true; break;
            }
        }

        xml << "<cacheField name=\"" << xmlEscape(cache.fields()[i]) << "\" numFmtId=\"0\"><sharedItems";
        if (!containsString) xml << " containsString=\"0\"";
        if (containsNumber) xml << " containsNumber=\"1\"";
        if (containsNumber && allNumbersInteger) xml << " containsInteger=\"1\"";
        if (containsBoolean) xml << " containsBoolean=\"1\"";
        if (containsError) xml << " containsError=\"1\"";
        if (containsBlank) xml << " containsBlank=\"1\"";
        if (nonBlankKinds.size() <= 1 && (containsNumber || containsBoolean || containsError))
            xml << " containsSemiMixedTypes=\"0\"";
        if (hasNumericRange) xml << " minValue=\"" << std::setprecision(15) << minValue
                                 << "\" maxValue=\"" << std::setprecision(15) << maxValue << "\"";
        const bool writeChildren = !pivotFieldIsPureData(pt, i);
        if (writeChildren) xml << " count=\"" << items.size() << "\"";

        if (!writeChildren || items.empty()) {
            xml << "/>";
        } else {
            xml << ">";
            for (const auto& value : items) writePivotValue(xml, value);
            xml << "</sharedItems>";
        }
        if (const auto* model = pivotAxisFieldByCacheIndex(pt, static_cast<int>(i)); model && model->grouping().active()) {
            const auto& group = model->grouping();
            xml << "<fieldGroup base=\"" << i << "\"><rangePr";
            xml << " autoStart=\"" << (group.autoStart ? 1 : 0) << "\" autoEnd=\"" << (group.autoEnd ? 1 : 0) << "\"";
            if (group.kind == xlpp::PivotGrouping::Kind::Numeric) {
                if (!group.autoStart) xml << " startNum=\"" << std::setprecision(15) << group.start << "\"";
                if (!group.autoEnd) xml << " endNum=\"" << std::setprecision(15) << group.end << "\"";
                xml << " groupBy=\"range\" groupInterval=\"" << group.interval << "\"";
            } else {
                if (!group.autoStart && !group.startDate.empty()) xml << " startDate=\"" << xmlEscape(group.startDate) << "\"";
                if (!group.autoEnd && !group.endDate.empty()) xml << " endDate=\"" << xmlEscape(group.endDate) << "\"";
                const char* groupBy = "months";
                switch (group.datePart) {
                    case xlpp::PivotGrouping::DatePart::Seconds: groupBy="seconds"; break; case xlpp::PivotGrouping::DatePart::Minutes: groupBy="minutes"; break;
                    case xlpp::PivotGrouping::DatePart::Hours: groupBy="hours"; break; case xlpp::PivotGrouping::DatePart::Days: groupBy="days"; break;
                    case xlpp::PivotGrouping::DatePart::Months: groupBy="months"; break; case xlpp::PivotGrouping::DatePart::Quarters: groupBy="quarters"; break;
                    case xlpp::PivotGrouping::DatePart::Years: groupBy="years"; break;
                }
                xml << " groupBy=\"" << groupBy << "\"";
            }
            xml << "/></fieldGroup>";
        }
        xml << "</cacheField>";
    }
    xml << "</cacheFields></pivotCacheDefinition>";
    return xml.str();
}

std::string pivotCacheRecordsXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    const auto& cache = pt.cache();
    const auto sharedItems = pivotSharedItems(cache);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheRecords xmlns=\""
        << nsMain(strict) << "\" count=\"" << cache.records().size() << "\">";
    for (const auto& record : cache.records()) {
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        xml << "<r>";
        for (std::size_t fieldIndex = 0; fieldIndex < record.size(); ++fieldIndex) {
            if (pivotFieldIsPureData(pt, fieldIndex)) {
                writePivotValue(xml, record[fieldIndex]);
            } else {
                const auto& items = sharedItems[fieldIndex];
                const auto item = std::find(items.begin(), items.end(), record[fieldIndex]);
                if (item == items.end()) throw std::logic_error("Pivot shared item lookup failed");
                xml << "<x v=\"" << std::distance(items.begin(), item) << "\"/>";
            }
        }
        xml << "</r>";
    }
    xml << "</pivotCacheRecords>";
    return xml.str();
}

std::string quotePivotSheetName(const std::string& name) {
    std::string escaped = name;
    std::size_t pos = 0;
    while ((pos = escaped.find('\'', pos)) != std::string::npos) {
        escaped.insert(pos, 1, '\'');
        pos += 2;
    }
    return "'" + escaped + "'";
}

std::string pivotCellText(const xlpp::Cell* cell) {
    if (!cell || !cell->hasValue()) return {};
    const auto& value = cell->value();
    if (const auto* text = std::get_if<std::string>(&value)) return *text;
    if (const auto* number = std::get_if<double>(&value)) {
        std::ostringstream out;
        out << std::setprecision(15) << *number;
        return out.str();
    }
    if (const auto* boolean = std::get_if<bool>(&value)) return *boolean ? "true" : "false";
    if (const auto* error = std::get_if<xlpp::CellError>(&value)) return xlpp::toString(*error);
    if (const auto* date = std::get_if<xlpp::DateTime>(&value)) return xlpp::toIso8601(*date);
    return {};
}

struct PivotSourceReference {
    std::string sheetName;
    xlpp::CellReference first;
    xlpp::CellReference last;
};

PivotSourceReference parsePivotSourceReference(const std::string& sourceData, const std::string& defaultSheet) {
    std::string sheetName = defaultSheet;
    std::string rangeText = sourceData;
    const auto bang = rangeText.find('!');
    if (bang != std::string::npos) {
        sheetName = rangeText.substr(0, bang);
        rangeText = rangeText.substr(bang + 1);
        if (sheetName.size() >= 2 && sheetName.front() == '\'' && sheetName.back() == '\'') {
            sheetName = sheetName.substr(1, sheetName.size() - 2);
            std::size_t pos = 0;
            while ((pos = sheetName.find("''", pos)) != std::string::npos) sheetName.replace(pos, 2, "'");
        }
    }
    const auto colon = rangeText.find(':');
    auto first = xlpp::CellReference::parse(colon == std::string::npos ? rangeText : rangeText.substr(0, colon));
    auto last = xlpp::CellReference::parse(colon == std::string::npos ? rangeText : rangeText.substr(colon + 1));
    if (first.row > last.row) std::swap(first.row, last.row);
    if (first.column > last.column) std::swap(first.column, last.column);
    return {std::move(sheetName), first, last};
}

xlpp::PivotTable effectivePivotTable(const xlpp::PivotTable& source,
                                     const std::deque<xlpp::Worksheet>& sheets,
                                     const xlpp::Worksheet& owner,
                                     std::size_t cacheId) {
    auto result = source;
    result.cache().setCacheId(static_cast<int>(cacheId));
    if (result.cache().sourceData().empty()) {
        result.cache().setSourceData(quotePivotSheetName(owner.name()) + "!" + owner.dimensions());
    }

    if (result.cache().fields().empty() || result.cache().records().empty()) {
        const auto parsed = parsePivotSourceReference(result.cache().sourceData(), owner.name());
        const auto sourceSheet = std::find_if(sheets.begin(), sheets.end(), [&](const auto& sheet) {
            return sheet.name() == parsed.sheetName;
        });
        if (sourceSheet == sheets.end())
            throw std::invalid_argument("Pivot source worksheet not found: " + parsed.sheetName);

        const auto width = parsed.last.column - parsed.first.column + 1;
        if (result.cache().fields().empty()) {
            std::vector<std::string> fields;
            fields.reserve(width);
            for (std::size_t column = parsed.first.column; column <= parsed.last.column; ++column) {
                auto name = pivotCellText(sourceSheet->tryCell(parsed.first.row, column));
                if (name.empty()) name = "Field" + std::to_string(column - parsed.first.column + 1);
                if (std::find(fields.begin(), fields.end(), name) != fields.end())
                    name += "_" + std::to_string(column - parsed.first.column + 1);
                fields.push_back(std::move(name));
            }
            result.cache().setFields(std::move(fields));
        }
        if (result.cache().fields().size() != width)
            throw std::invalid_argument("Pivot source width does not match cache field count");

        if (result.cache().records().empty() && parsed.first.row < parsed.last.row) {
            for (std::size_t row = parsed.first.row + 1; row <= parsed.last.row; ++row) {
                std::vector<std::string> record;
                record.reserve(width);
                for (std::size_t column = parsed.first.column; column <= parsed.last.column; ++column)
                    record.push_back(pivotCellText(sourceSheet->tryCell(row, column)));
                result.cache().addRecord(std::move(record));
            }
        }
    }
    return result;
}


void loadImportedPivotModels(xlpp::Worksheet& sheet,
                             const xlpp::internal::ZipArchive& archive,
                             const std::string& sourceSheetPart,
                             const std::vector<xlpp::PreservedRelationship>& relationships) {
    for (const auto& relationship : relationshipsForSource(relationships, sourceSheetPart)) {
        if (relationshipKind(relationship) != "pivotTable" || relationship.targetMode == "External") continue;
        const auto pivotPart = resolvePackagePart(sourceSheetPart, relationship.target);
        if (pivotPart.empty() || !archive.contains(pivotPart)) continue;
        const auto pivotXml = archive.get(pivotPart);
        const auto definitions = xlpp::internal::tags(pivotXml, "pivotTableDefinition");
        if (definitions.empty()) continue;
        const auto& definition = definitions.front();
        xlpp::PivotTable pivot(xlpp::internal::attribute(definition, "name"));
        if (pivot.name().empty()) pivot.setName("ImportedPivot" + std::to_string(sheet.loadedPivotCount() + 1));
        const auto attrBool = [](const std::string& xml, const char* name, bool fallback) {
            const auto value = xlpp::internal::attribute(xml, name);
            if (value.empty()) return fallback;
            return value != "0" && value != "false" && value != "FALSE";
        };
        pivot.setRowGrandTotals(attrBool(definition, "rowGrandTotals", true));
        pivot.setColumnGrandTotals(attrBool(definition, "colGrandTotals", true));
        pivot.setPreserveFormatting(attrBool(definition, "preserveFormatting", true));
        pivot.setUseAutoFormatting(attrBool(definition, "useAutoFormatting", true));
        pivot.setShowEmptyRow(attrBool(definition, "showEmptyRow", false));
        pivot.setShowEmptyColumn(attrBool(definition, "showEmptyCol", false));
        pivot.setShowDrill(attrBool(definition, "showDrill", true));
        pivot.setEnableDrill(attrBool(definition, "enableDrill", true));
        pivot.setShowDataTips(attrBool(definition, "showDataTips", true));
        pivot.setShowMemberPropertyTips(attrBool(definition, "showMemberPropertyTips", true));
        pivot.setShowHeaders(attrBool(definition, "showHeaders", true));
        pivot.setMultipleFieldFilters(attrBool(definition, "multipleFieldFilters", false));
        pivot.setShowValuesRow(attrBool(definition, "showValuesRow", false));
        pivot.setSubtotalHiddenItems(attrBool(definition, "subtotalHiddenItems", false));
        pivot.setPageOverThenDown(attrBool(definition, "pageOverThenDown", false));
        if (const auto pageWrap = xlpp::internal::attribute(definition, "pageWrap"); !pageWrap.empty()) { try { pivot.setPageWrap(std::stoi(pageWrap)); } catch (...) {} }
        if (const auto caption = xlpp::internal::attribute(definition, "dataCaption"); !caption.empty()) pivot.setDataCaption(caption);
        const auto compact = attrBool(definition, "compact", true);
        const auto outline = attrBool(definition, "outline", true);
        pivot.setLayout(compact ? xlpp::PivotLayout::Compact : (outline ? xlpp::PivotLayout::Outline : xlpp::PivotLayout::Tabular));
        if (const auto styles = xlpp::internal::tags(definition, "pivotTableStyleInfo"); !styles.empty()) {
            const auto& style = styles.front();
            if (const auto name = xlpp::internal::attribute(style, "name"); !name.empty()) pivot.setStyleName(name);
            pivot.setShowRowHeaders(attrBool(style, "showRowHeaders", true));
            pivot.setShowColumnHeaders(attrBool(style, "showColHeaders", true));
            pivot.setShowRowStripes(attrBool(style, "showRowStripes", false));
            pivot.setShowColumnStripes(attrBool(style, "showColStripes", false));
            pivot.setShowLastColumn(attrBool(style, "showLastColumn", true));
        }
        if (const auto locations = xlpp::internal::tags(definition, "location"); !locations.empty())
            pivot.setLocation(xlpp::internal::attribute(locations.front(), "ref"));
        const auto cacheIdText = xlpp::internal::attribute(definition, "cacheId");
        if (!cacheIdText.empty()) {
            try { pivot.cache().setCacheId(std::max(1, std::stoi(cacheIdText))); } catch (...) {}
        }

        std::string cachePart;
        for (const auto& cacheRel : relationshipsForSource(relationships, pivotPart)) {
            if (relationshipKind(cacheRel) == "pivotCacheDefinition" && cacheRel.targetMode != "External") {
                cachePart = resolvePackagePart(pivotPart, cacheRel.target); break;
            }
        }
        std::vector<std::vector<std::string>> importedSharedItems;
        std::vector<xlpp::PivotGrouping> importedGroupings;
        if (!cachePart.empty() && archive.contains(cachePart)) {
            const auto cacheXml = archive.get(cachePart);
            if (const auto cacheDefinitions = xlpp::internal::tags(cacheXml, "pivotCacheDefinition"); !cacheDefinitions.empty()) {
                const auto& cacheDefinition = cacheDefinitions.front();
                pivot.cache().setRefreshOnLoad(attrBool(cacheDefinition, "refreshOnLoad", false));
                pivot.cache().setSaveData(attrBool(cacheDefinition, "saveData", true));
                pivot.cache().setEnableRefresh(attrBool(cacheDefinition, "enableRefresh", true));
                pivot.cache().setBackgroundQuery(attrBool(cacheDefinition, "backgroundQuery", false));
                pivot.cache().setOptimizeMemory(attrBool(cacheDefinition, "optimizeMemory", false));
                pivot.cache().setUpgradeOnRefresh(attrBool(cacheDefinition, "upgradeOnRefresh", false));
                pivot.cache().setSupportSubquery(attrBool(cacheDefinition, "supportSubquery", true));
                pivot.cache().setSupportAdvancedDrill(attrBool(cacheDefinition, "supportAdvancedDrill", true));
                if (const auto limit = xlpp::internal::attribute(cacheDefinition, "missingItemsLimit"); !limit.empty()) { try { pivot.cache().setMissingItemsLimit(std::stoi(limit)); } catch (...) {} }
                if (const auto refreshedBy = xlpp::internal::attribute(cacheDefinition, "refreshedBy"); !refreshedBy.empty()) pivot.cache().setRefreshedBy(refreshedBy);
            }
            if (const auto sources = xlpp::internal::tags(cacheXml, "worksheetSource"); !sources.empty()) {
                const auto ref = xlpp::internal::attribute(sources.front(), "ref");
                const auto sourceSheet = xlpp::internal::attribute(sources.front(), "sheet");
                if (!ref.empty()) pivot.cache().setSourceData(sourceSheet.empty() ? ref : quotePivotSheetName(sourceSheet) + "!" + ref);
            }
            auto directChildren = [](std::string_view node) {
                std::vector<std::string> result;
                const auto openEnd = node.find('>');
                const auto closeBegin = node.rfind("</");
                if (openEnd == std::string_view::npos || closeBegin == std::string_view::npos || closeBegin <= openEnd) return result;
                std::size_t pos = openEnd + 1;
                while (pos < closeBegin) {
                    const auto lt = node.find('<', pos);
                    if (lt == std::string_view::npos || lt >= closeBegin) break;
                    if (lt + 1 >= closeBegin || node[lt + 1] == '/') { pos = lt + 1; continue; }
                    const auto gt = node.find('>', lt + 1);
                    if (gt == std::string_view::npos || gt > closeBegin) break;
                    std::size_t nameEnd = lt + 1;
                    while (nameEnd < gt && node[nameEnd] != ' ' && node[nameEnd] != '\t' && node[nameEnd] != '\r' && node[nameEnd] != '\n' && node[nameEnd] != '/' && node[nameEnd] != '>') ++nameEnd;
                    const auto name = std::string(node.substr(lt + 1, nameEnd - lt - 1));
                    if (gt > lt && node[gt - 1] == '/') {
                        result.emplace_back(node.substr(lt, gt - lt + 1)); pos = gt + 1; continue;
                    }
                    const auto close = node.find("</" + name + ">", gt + 1);
                    if (close == std::string_view::npos || close >= closeBegin) { pos = gt + 1; continue; }
                    result.emplace_back(node.substr(lt, close + name.size() + 3 - lt));
                    pos = close + name.size() + 3;
                }
                return result;
            };
            auto childLocalName = [](const std::string& node) {
                if (node.size() < 2 || node[0] != '<') return std::string{};
                auto begin = std::size_t{1}; auto end = begin;
                while (end < node.size() && node[end] != ' ' && node[end] != '\t' && node[end] != '\r' && node[end] != '\n' && node[end] != '/' && node[end] != '>') ++end;
                auto name = node.substr(begin, end - begin); const auto colon = name.find(':'); if (colon != std::string::npos) name.erase(0, colon + 1); return name;
            };
            std::vector<std::string> fields;
            for (const auto& field : xlpp::internal::tags(cacheXml, "cacheField")) {
                auto name = xlpp::internal::attribute(field, "name");
                if (name.empty()) name = "Field" + std::to_string(fields.size() + 1);
                fields.push_back(std::move(name));
                std::vector<std::string> shared;
                if (const auto nodes = xlpp::internal::tags(field, "sharedItems"); !nodes.empty()) {
                    for (const auto& child : directChildren(nodes.front())) {
                        const auto kind = childLocalName(child);
                        if (kind == "m") shared.emplace_back();
                        else shared.push_back(xlpp::internal::attribute(child, "v"));
                    }
                }
                importedSharedItems.push_back(std::move(shared));
                xlpp::PivotGrouping grouping;
                if (const auto groups = xlpp::internal::tags(field, "fieldGroup"); !groups.empty()) {
                    if (const auto ranges = xlpp::internal::tags(groups.front(), "rangePr"); !ranges.empty()) {
                        const auto& range = ranges.front();
                        const auto groupBy = xlpp::internal::attribute(range, "groupBy");
                        grouping.kind = groupBy == "range" ? xlpp::PivotGrouping::Kind::Numeric : xlpp::PivotGrouping::Kind::Date;
                        grouping.autoStart = attrBool(range, "autoStart", true); grouping.autoEnd = attrBool(range, "autoEnd", true);
                        if (const auto v=xlpp::internal::attribute(range,"startNum"); !v.empty()) { try { grouping.start=std::stod(v); } catch (...) {} }
                        if (const auto v=xlpp::internal::attribute(range,"endNum"); !v.empty()) { try { grouping.end=std::stod(v); } catch (...) {} }
                        if (const auto v=xlpp::internal::attribute(range,"groupInterval"); !v.empty()) { try { grouping.interval=std::stod(v); } catch (...) {} }
                        grouping.startDate = xlpp::internal::attribute(range, "startDate"); grouping.endDate = xlpp::internal::attribute(range, "endDate");
                        if (groupBy=="seconds") grouping.datePart=xlpp::PivotGrouping::DatePart::Seconds; else if(groupBy=="minutes") grouping.datePart=xlpp::PivotGrouping::DatePart::Minutes;
                        else if(groupBy=="hours") grouping.datePart=xlpp::PivotGrouping::DatePart::Hours; else if(groupBy=="days") grouping.datePart=xlpp::PivotGrouping::DatePart::Days;
                        else if(groupBy=="quarters") grouping.datePart=xlpp::PivotGrouping::DatePart::Quarters; else if(groupBy=="years") grouping.datePart=xlpp::PivotGrouping::DatePart::Years;
                        else grouping.datePart=xlpp::PivotGrouping::DatePart::Months;
                    }
                }
                importedGroupings.push_back(std::move(grouping));
            }
            if (!fields.empty()) pivot.cache().setFields(std::move(fields));

            std::string recordsPart;
            for (const auto& recordsRel : relationshipsForSource(relationships, cachePart)) {
                if (relationshipKind(recordsRel) == "pivotCacheRecords" && recordsRel.targetMode != "External") { recordsPart = resolvePackagePart(cachePart, recordsRel.target); break; }
            }
            if (!recordsPart.empty() && archive.contains(recordsPart)) {
                const auto recordsXml = archive.get(recordsPart);
                for (const auto& recordNode : xlpp::internal::tags(recordsXml, "r")) {
                    std::vector<std::string> record;
                    std::size_t fieldIndex = 0;
                    for (const auto& child : directChildren(recordNode)) {
                        if (fieldIndex >= pivot.cache().fields().size()) break;
                        const auto kind = childLocalName(child);
                        if (kind == "x") {
                            try {
                                const auto index = static_cast<std::size_t>(std::stoull(xlpp::internal::attribute(child, "v")));
                                if (fieldIndex < importedSharedItems.size() && index < importedSharedItems[fieldIndex].size()) record.push_back(importedSharedItems[fieldIndex][index]);
                                else record.emplace_back();
                            } catch (...) { record.emplace_back(); }
                        } else if (kind == "m") record.emplace_back();
                        else record.push_back(xlpp::internal::attribute(child, "v"));
                        ++fieldIndex;
                    }
                    while (record.size() < pivot.cache().fields().size()) record.emplace_back();
                    if (record.size() == pivot.cache().fields().size()) pivot.cache().addRecord(std::move(record));
                }
            }
        }

        auto addAxis = [&](const char* containerName, const char* childName, const char* indexAttr, const char* axis) {
            const auto containers = xlpp::internal::tags(definition, containerName);
            if (containers.empty()) return;
            for (const auto& node : xlpp::internal::tags(containers.front(), childName)) {
                const auto text = xlpp::internal::attribute(node, indexAttr);
                if (text.empty()) continue;
                try {
                    const auto index = std::stoi(text);
                    if (index < 0 || static_cast<std::size_t>(index) >= pivot.cache().fields().size()) continue;
                    xlpp::PivotField* added = nullptr;
                    if (std::string_view(axis) == "row") added = &pivot.addRowField(pivot.cache().fields()[static_cast<std::size_t>(index)]);
                    else if (std::string_view(axis) == "col") added = &pivot.addColumnField(pivot.cache().fields()[static_cast<std::size_t>(index)]);
                    else {
                        added = &pivot.addPageField(pivot.cache().fields()[static_cast<std::size_t>(index)]);
                        if (const auto item = xlpp::internal::attribute(node, "item"); !item.empty()) {
                            try { added->setSelectedItemIndex(std::stoi(item)); } catch (...) {}
                        }
                    }
                    added->setFieldIndex(index);
                } catch (...) {}
            }
        };
        addAxis("rowFields", "field", "x", "row");
        addAxis("colFields", "field", "x", "col");
        addAxis("pageFields", "pageField", "fld", "page");
        if (const auto data = xlpp::internal::tags(definition, "dataFields"); !data.empty()) {
            for (const auto& node : xlpp::internal::tags(data.front(), "dataField")) {
                const auto text = xlpp::internal::attribute(node, "fld");
                if (text.empty()) continue;
                try {
                    const auto index = std::stoi(text);
                    if (index < 0 || static_cast<std::size_t>(index) >= pivot.cache().fields().size()) continue;
                    auto subtotal = xlpp::internal::attribute(node, "subtotal");
                    if (subtotal.empty()) subtotal = "sum";
                    auto& added = pivot.addDataField(pivot.cache().fields()[static_cast<std::size_t>(index)], subtotal);
                    added.setFieldIndex(index);
                    if (const auto caption = xlpp::internal::attribute(node, "name"); !caption.empty()) added.setCaption(caption);
                    if (const auto numFmt = xlpp::internal::attribute(node, "numFmtId"); !numFmt.empty()) { try { added.setNumberFormatId(static_cast<std::uint32_t>(std::stoul(numFmt))); } catch (...) {} }
                    if (const auto showDataAs = xlpp::internal::attribute(node, "showDataAs"); !showDataAs.empty()) {
                        added.setShowDataAs(showDataAs);
                    } else if (const auto extended = xlpp::internal::tags(node, "x14:dataField"); !extended.empty()) {
                        if (const auto pivotShowAs = xlpp::internal::attribute(extended.front(), "pivotShowAs"); !pivotShowAs.empty())
                            added.setShowDataAs(pivotShowAs);
                    }
                    if (const auto baseField = xlpp::internal::attribute(node, "baseField"); !baseField.empty()) { try { added.setBaseField(std::stoi(baseField)); } catch (...) {} }
                    if (const auto baseItem = xlpp::internal::attribute(node, "baseItem"); !baseItem.empty()) { try { added.setBaseItem(std::stoi(baseItem)); } catch (...) {} }
                } catch (...) {}
            }
        }

        // Recover common field flags from pivotFields by cache index.
        if (const auto fieldGroups = xlpp::internal::tags(definition, "pivotFields"); !fieldGroups.empty()) {
            const auto fieldNodes = xlpp::internal::tags(fieldGroups.front(), "pivotField");
            auto applyFlags = [&](std::vector<xlpp::PivotField>& values) {
                for (auto& value : values) {
                    const auto index = value.fieldIndex();
                    if (index < 0 || static_cast<std::size_t>(index) >= fieldNodes.size()) continue;
                    const auto& node = fieldNodes[static_cast<std::size_t>(index)];
                    value.setShowAll(attrBool(node, "showAll", true));
                    const auto sortType = xlpp::internal::attribute(node, "sortType");
                    value.setSortType(sortType == "ascending" ? 1 : (sortType == "descending" ? 2 : 0));
                    value.setSubtotalTop(attrBool(node, "subtotalTop", true));
                    value.setInsertBlankRow(attrBool(node, "insertBlankRow", false));
                    value.setIncludeNewItemsInFilter(attrBool(node, "includeNewItemsInFilter", false));
                    value.setMultipleItemSelectionAllowed(attrBool(node, "multipleItemSelectionAllowed", false));
                    value.setCompact(attrBool(node, "compact", true));
                    value.setOutline(attrBool(node, "outline", true));
                    value.setInsertPageBreak(attrBool(node, "insertPageBreak", false));
                    value.setShowDropDowns(attrBool(node, "showDropDowns", true));
                    value.setDefaultSubtotal(attrBool(node, "defaultSubtotal", true));
                    static const std::vector<std::string> subtotalNames{"sum","countA","avg","max","min","product","count","stdDev","stdDevP","var","varP"};
                    std::vector<std::string> subtotals;
                    for (const auto& subtotal : subtotalNames) if (attrBool(node, (subtotal + "Subtotal").c_str(), false)) subtotals.push_back(subtotal);
                    value.setSubtotals(std::move(subtotals));
                    std::vector<int> hidden;
                    if (const auto itemLists = xlpp::internal::tags(node, "items"); !itemLists.empty()) {
                        const auto itemNodes = xlpp::internal::tags(itemLists.front(), "item");
                        for (std::size_t itemIndex = 0; itemIndex < itemNodes.size(); ++itemIndex) if (attrBool(itemNodes[itemIndex], "h", false)) hidden.push_back(static_cast<int>(itemIndex));
                    }
                    value.setHiddenItemIndexes(std::move(hidden));
                    value.setRepeatItemLabels(!xlpp::internal::tags(node, "x14:pivotField").empty());
                    if (index >= 0 && static_cast<std::size_t>(index) < importedGroupings.size() && importedGroupings[static_cast<std::size_t>(index)].active()) value.setGrouping(importedGroupings[static_cast<std::size_t>(index)]);
                }
            };
            applyFlags(pivot.rowFields()); applyFlags(pivot.columnFields()); applyFlags(pivot.pageFields());
        }
        if (const auto filterLists = xlpp::internal::tags(definition, "filters"); !filterLists.empty()) {
            for (const auto& node : xlpp::internal::tags(filterLists.front(), "filter")) {
                xlpp::PivotFilter filter;
                filter.type = xlpp::internal::attribute(node, "type");
                try { filter.fieldIndex = std::stoi(xlpp::internal::attribute(node, "fld")); } catch (...) { continue; }
                if (const auto v=xlpp::internal::attribute(node,"iMeasureFld"); !v.empty()) { try { filter.measureFieldIndex=std::stoi(v); } catch (...) {} }
                filter.value1=xlpp::internal::attribute(node,"stringValue1"); filter.value2=xlpp::internal::attribute(node,"stringValue2");
                if (const auto top10=xlpp::internal::tags(node,"top10"); !top10.empty()) {
                    filter.top10Top=attrBool(top10.front(),"top",true); filter.top10Percent=attrBool(top10.front(),"percent",false);
                    if (const auto v=xlpp::internal::attribute(top10.front(),"val"); !v.empty()) { try { filter.top10Value=std::stod(v); } catch (...) {} }
                }
                pivot.addFilter(std::move(filter));
            }
        }
        sheet.addLoadedPivotTable(std::move(pivot));
    }
}


} // namespace xlpp::internal::ooxml
