#include "WorkbookPivotWrite.h"
#include "WorkbookNamespaces.h"
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Cell/Cell.h>
#include <XLPP/Cell/CellReference.h>
#include "../XML/XmlUtilities.h"
#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdlib>
#include <iomanip>
#include <set>
#include <stdexcept>
#include <utility>

namespace xlpp {
namespace internal {
int resolvedPivotFieldIndex(const xlpp::PivotCache& cache, int index, const std::string& name) {
    if (index >= 0 && static_cast<std::size_t>(index) < cache.fields().size()) return index;
    if (!name.empty()) {
        const auto resolved = cache.fieldIndex(name);
        if (resolved >= 0) return resolved;
    }
    throw std::invalid_argument("Pivot field '" + name + "' is not present in the cache fields");
}

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

PivotValueKind pivotValueKind(xlpp::PivotCacheValueKind kind) {
    using K = xlpp::PivotCacheValueKind;
    switch (kind) {
    case K::Missing: return PivotValueKind::Blank;
    case K::Number: return PivotValueKind::Number;
    case K::String: return PivotValueKind::String;
    case K::Boolean: return PivotValueKind::Boolean;
    case K::Error: return PivotValueKind::Error;
    case K::DateTime: return PivotValueKind::DateTime;
    }
    return PivotValueKind::String;
}

xlpp::PivotCacheValueKind publicPivotValueKind(std::string_view nodeType) {
    if (nodeType == "m") return xlpp::PivotCacheValueKind::Missing;
    if (nodeType == "n") return xlpp::PivotCacheValueKind::Number;
    if (nodeType == "b") return xlpp::PivotCacheValueKind::Boolean;
    if (nodeType == "e") return xlpp::PivotCacheValueKind::Error;
    if (nodeType == "d") return xlpp::PivotCacheValueKind::DateTime;
    return xlpp::PivotCacheValueKind::String;
}

xlpp::PivotCacheValueKind publicPivotValueKind(PivotValueKind kind) {
    switch (kind) {
    case PivotValueKind::Blank: return xlpp::PivotCacheValueKind::Missing;
    case PivotValueKind::Number: return xlpp::PivotCacheValueKind::Number;
    case PivotValueKind::Boolean: return xlpp::PivotCacheValueKind::Boolean;
    case PivotValueKind::Error: return xlpp::PivotCacheValueKind::Error;
    case PivotValueKind::DateTime: return xlpp::PivotCacheValueKind::DateTime;
    case PivotValueKind::String: return xlpp::PivotCacheValueKind::String;
    }
    return xlpp::PivotCacheValueKind::String;
}

bool samePivotSharedItem(const PivotSharedItem& a, const PivotSharedItem& b) {
    return a.kind == b.kind && a.value == b.value;
}

PivotValueKind recordValueKind(const xlpp::PivotCache& cache, std::size_t row, std::size_t column,
                               const std::string& value) {
    if (cache.hasTypedRecordKinds()) return pivotValueKind(cache.recordKinds()[row][column]);
    return parsePivotValue(value).kind;
}

std::vector<std::vector<PivotSharedItem>> pivotSharedItems(const xlpp::PivotCache& cache) {
    std::vector<std::vector<PivotSharedItem>> result(cache.fields().size());
    for (std::size_t rowIndex = 0; rowIndex < cache.records().size(); ++rowIndex) {
        const auto& record = cache.records()[rowIndex];
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        for (std::size_t i = 0; i < record.size(); ++i) {
            auto& items = result[i];
            PivotSharedItem candidate{record[i], recordValueKind(cache, rowIndex, i, record[i])};
            if (std::find_if(items.begin(), items.end(), [&](const auto& item) { return samePivotSharedItem(item, candidate); }) == items.end())
                items.push_back(std::move(candidate));
        }
    }
    return result;
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

    const auto sharedItems = pivotSharedItems(pt.cache());

    // A PivotTable location is the complete occupied view range, not merely the
    // requested top-left cell.  When the caller supplies an anchor, derive the
    // smallest view rectangle from the emitted row-item matrix.  This keeps the
    // location, rowItems and colItems mutually consistent for desktop Excel.
    std::string location = pt.location().empty() ? "D2" : pt.location();
    if (location.find(':') == std::string::npos) {
        const auto first = xlpp::CellReference::parse(location);
        const auto rowItemCount = rowIndexes.empty()
            ? std::size_t{1}
            : sharedItems[static_cast<std::size_t>(rowIndexes.front())].size() + 1; // grand total
        const auto outputRows = std::max<std::size_t>(2, rowItemCount + 1);          // header row
        const auto outputColumns = std::max<std::size_t>(2, 1 + dataIndexes.size());
        xlpp::CellReference last{first.row + outputRows - 1, first.column + outputColumns - 1};
        location += ":" + last.address();
    }

    // Match the conservative view metadata emitted by desktop Excel.  The
    // extension list and revision UID are optional, but the layout/version flags
    // below make the intended non-OLAP compact view unambiguous.
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotTableDefinition xmlns=\""
        << nsMain(strict) << "\" name=\"" << xmlEscape(pt.name()) << "\" cacheId=\"" << id
        << "\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\""
        << " applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\""
        << " dataCaption=\"Values\" updatedVersion=\"8\" minRefreshableVersion=\"3\""
        << " useAutoFormatting=\"" << (pt.useAutoFormatting() ? 1 : 0) << "\" preserveFormatting=\"" << (pt.preserveFormatting() ? 1 : 0) << "\""
        << " rowGrandTotals=\"" << (pt.rowGrandTotals() ? 1 : 0) << "\" colGrandTotals=\"" << (pt.columnGrandTotals() ? 1 : 0) << "\""
        << " showDrill=\"" << (pt.showDrill() ? 1 : 0) << "\" compact=\"" << (pt.compact() ? 1 : 0) << "\""
        << " itemPrintTitles=\"1\" createdVersion=\"8\" indent=\"0\" outline=\"" << (pt.outline() ? 1 : 0)
        << "\" outlineData=\"" << (pt.outline() ? 1 : 0) << "\" multipleFieldFilters=\"" << (pt.multipleFieldFilters() ? 1 : 0) << "\"";
    if (pt.chartFormatIndex()) xml << " chartFormat=\"" << *pt.chartFormatIndex() << "\"";
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
            // Keep the generated XML minimal and backward-compatible. Excel's
            // defaults are compact/outline=true, so only materialize explicit
            // non-default values. Imported values are still represented in the
            // object model and are written when they differ from the defaults.
            if (!modelField->compact()) xml << " compact=\"0\"";
            if (!modelField->outline()) xml << " outline=\"0\"";
            if (modelField->sortType() == 1) xml << " sortType=\"ascending\"";
            else if (modelField->sortType() == 2) xml << " sortType=\"descending\"";
            for (const auto& subtotal : modelField->subtotals()) {
                const auto attribute = subtotal + "Subtotal";
                xml << " " << attribute << "=\"1\"";
            }
        }
        if (isRow || isCol || isPage) xml << " defaultSubtotal=\"" << ((modelField && modelField->defaultSubtotal()) ? 1 : 0) << "\"";
        if (isRow || isCol || isPage) {
            // The axis item list contains only concrete cache items.  A synthetic
            // default item is not needed for an unfiltered field and can make the
            // row/column view indexes ambiguous to Excel.
            xml << "><items count=\"" << sharedItems[fieldIndex].size() << "\">";
            for (std::size_t itemIndex = 0; itemIndex < sharedItems[fieldIndex].size(); ++itemIndex) {
                const xlpp::PivotFieldItem* itemModel = nullptr;
                if (modelField) {
                    const auto found = std::find_if(modelField->items().begin(), modelField->items().end(), [&](const auto& candidate) {
                        if (candidate.hasCacheValue) {
                            const auto& current = sharedItems[fieldIndex][itemIndex];
                            return candidate.cacheValue == current.value && pivotValueKind(candidate.cacheValueKind) == current.kind;
                        }
                        return candidate.cacheIndex == static_cast<int>(itemIndex);
                    });
                    if (found != modelField->items().end()) itemModel = &*found;
                }
                const bool hidden = itemModel ? itemModel->hidden : (modelField && std::find(modelField->hiddenItems().begin(), modelField->hiddenItems().end(), static_cast<int>(itemIndex)) != modelField->hiddenItems().end());
                xml << "<item x=\"" << itemIndex << "\"";
                if (itemModel) {
                    if (!itemModel->type.empty()) xml << " t=\"" << xmlEscape(itemModel->type) << "\"";
                    if (!itemModel->caption.empty()) xml << " n=\"" << xmlEscape(itemModel->caption) << "\"";
                    if (!itemModel->showDetails) xml << " sd=\"0\"";
                    if (itemModel->formula) xml << " f=\"1\"";
                    if (itemModel->missing) xml << " m=\"1\"";
                }
                if (hidden) xml << " h=\"1\"";
                xml << "/>";
            }
            xml << "</items></pivotField>";
        } else {
            xml << "/>";
        }
    }
    xml << "</pivotFields>";

    if (!rowIndexes.empty()) {
        xml << "<rowFields count=\"" << rowIndexes.size() << "\">";
        for (const auto index : rowIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</rowFields>";
        // A complete multi-axis item matrix is significantly more complex.  The
        // first row field is sufficient for the common one-row-axis pivot and is
        // accepted by Excel/LibreOffice; additional row fields are refreshed from
        // the cache when the workbook is opened.
        const auto itemCount = sharedItems[static_cast<std::size_t>(rowIndexes.front())].size();
        xml << "<rowItems count=\"" << (itemCount + 1) << "\">";
        for (std::size_t itemIndex = 0; itemIndex < itemCount; ++itemIndex)
            xml << "<i><x v=\"" << itemIndex << "\"/></i>";
        xml << "<i t=\"grand\"><x/></i></rowItems>";
    }
    if (!columnIndexes.empty()) {
        xml << "<colFields count=\"" << columnIndexes.size() << "\">";
        for (const auto index : columnIndexes) xml << "<field x=\"" << index << "\"/>";
        xml << "</colFields>";
    }
    if (columnIndexes.empty()) {
        // With no explicit column field and a single values field, the canonical
        // column item is an empty data item.  Do not invent a field/item index.
        xml << "<colItems count=\"1\"><i/></colItems>";
    } else {
        xml << "<colItems count=\"1\"><i t=\"grand\"><x/></i></colItems>";
    }
    if (!pageIndexes.empty()) {
        xml << "<pageFields count=\"" << pageIndexes.size() << "\">";
        for (std::size_t i = 0; i < pageIndexes.size(); ++i) {
            const auto index = pageIndexes[i];
            const xlpp::PivotPageField* setting = i < pt.pageFieldSettings().size() ? &pt.pageFieldSettings()[i] : nullptr;
            xml << "<pageField fld=\"" << index << "\" hier=\"" << (setting ? setting->hierarchy() : -1) << "\"";
            if (setting && setting->item() >= 0) xml << " item=\"" << setting->item() << "\"";
            if (setting && !setting->name().empty()) xml << " name=\"" << xmlEscape(setting->name()) << "\"";
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
            const auto displayName = field.displayName().empty() ? (subtotal == "sum" ? "Sum of " + sourceName : sourceName) : field.displayName();
            xml << "<dataField name=\"" << xmlEscape(displayName)
                << "\" fld=\"" << index << "\"";
            if (field.baseField() >= 0) xml << " baseField=\"" << field.baseField() << "\"";
            if (field.baseItem() >= 0) xml << " baseItem=\"" << field.baseItem() << "\"";
            if (subtotal != "sum") xml << " subtotal=\"" << xmlEscape(subtotal) << "\"";
            if (field.showDataAs() != "normal") xml << " showDataAs=\"" << xmlEscape(field.showDataAs()) << "\"";
            if (field.numberFormatId() > 0) xml << " numFmtId=\"" << field.numberFormatId() << "\"";
            xml << "/>";
        }
        xml << "</dataFields>";
    }
    // CT_PivotTableDefinition requires chartFormats before the style node and
    // filters after it.  Keeping this schema order matters for strict consumers.
    if (!pt.chartFormats().empty()) {
        xml << "<chartFormats count=\"" << pt.chartFormats().size() << "\">";
        for (const auto& format : pt.chartFormats()) {
            xml << "<chartFormat chart=\"" << format.chartIndex << "\" format=\""
                << format.formatId << "\" series=\"" << (format.series ? 1 : 0) << "\"";
            if (format.pivotAreaXml.empty()) xml << "/>";
            else xml << ">" << format.pivotAreaXml << "</chartFormat>";
        }
        xml << "</chartFormats>";
    }
    xml << "<pivotTableStyleInfo name=\"" << xmlEscape(pt.styleName()) << "\" showRowHeaders=\"" << (pt.showRowHeaders() ? 1 : 0)
        << "\" showColHeaders=\"" << (pt.showColumnHeaders() ? 1 : 0) << "\" showRowStripes=\"" << (pt.showRowStripes() ? 1 : 0)
        << "\" showColStripes=\"" << (pt.showColumnStripes() ? 1 : 0) << "\" showLastColumn=\"1\"/>";
    if (!pt.filters().empty()) {
        xml << "<filters count=\"" << pt.filters().size() << "\">";
        for (const auto& filter : pt.filters()) {
            if (filter.fieldIndex < 0 || filter.type.empty())
                throw std::invalid_argument("Pivot filter requires a field index and type");
            xml << "<filter fld=\"" << filter.fieldIndex << "\" type=\"" << xmlEscape(filter.type)
                << "\" id=\"" << filter.id << "\"";
            if (filter.evaluationOrder != 0) xml << " evalOrder=\"" << filter.evaluationOrder << "\"";
            if (filter.measureField >= 0) xml << " iMeasureFld=\"" << filter.measureField << "\"";
            if (filter.measureHierarchy >= 0) xml << " iMeasureHier=\"" << filter.measureHierarchy << "\"";
            if (filter.memberPropertyField >= 0) xml << " mpFld=\"" << filter.memberPropertyField << "\"";
            if (!filter.name.empty()) xml << " name=\"" << xmlEscape(filter.name) << "\"";
            if (!filter.description.empty()) xml << " description=\"" << xmlEscape(filter.description) << "\"";
            if (!filter.stringValue1.empty()) xml << " stringValue1=\"" << xmlEscape(filter.stringValue1) << "\"";
            if (!filter.stringValue2.empty()) xml << " stringValue2=\"" << xmlEscape(filter.stringValue2) << "\"";
            if (filter.autoFilterXml.empty()) xml << "/>";
            else xml << ">" << filter.autoFilterXml << "</filter>";
        }
        xml << "</filters>";
    }
    xml << "</pivotTableDefinition>";
    return xml.str();
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

void writePivotValue(std::ostringstream& xml, const std::string& value, PivotValueKind kind) {
    switch (kind) {
        case PivotValueKind::Blank: xml << "<m/>"; break;
        case PivotValueKind::Number: {
            const auto parsed = parsePivotValue(value);
            if (parsed.kind != PivotValueKind::Number)
                throw std::invalid_argument("Typed Pivot numeric value is not a finite number: " + value);
            xml << "<n v=\"" << std::setprecision(15) << parsed.number << "\"/>";
            break;
        }
        case PivotValueKind::Boolean:
            if (value != "true" && value != "false" && value != "1" && value != "0")
                throw std::invalid_argument("Typed Pivot boolean value must be true/false/1/0");
            xml << "<b v=\"" << ((value == "true" || value == "1") ? 1 : 0) << "\"/>";
            break;
        case PivotValueKind::Error: xml << "<e v=\"" << xmlEscape(value) << "\"/>"; break;
        case PivotValueKind::DateTime: xml << "<d v=\"" << xmlEscape(value) << "\"/>"; break;
        case PivotValueKind::String: xml << "<s v=\"" << xmlEscape(value) << "\"/>"; break;
    }
}

void writePivotValue(std::ostringstream& xml, const std::string& value) {
    writePivotValue(xml, value, parsePivotValue(value).kind);
}

std::string pivotCacheXml(const xlpp::PivotTable& pt, bool strict) {
    std::ostringstream xml;
    const auto& cache = pt.cache();
    const auto fieldCount = cache.fields().size();
    const auto sharedItems = pivotSharedItems(cache);
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><pivotCacheDefinition xmlns=\""
        << nsMain(strict) << "\" xmlns:r=\"" << nsRelsDoc(strict)
        << "\" r:id=\"rId1\" saveData=\"" << (cache.saveData() ? 1 : 0) << "\" refreshOnLoad=\"" << (cache.refreshOnLoad() ? 1 : 0)
        << "\" enableRefresh=\"" << (cache.enableRefresh() ? 1 : 0) << "\""
        << " createdVersion=\"3\" refreshedVersion=\"3\" minRefreshableVersion=\"3\""
        << (cache.missingItemsLimit() >= 0 ? " missingItemsLimit=\"" + std::to_string(cache.missingItemsLimit()) + "\"" : std::string{})
        << " recordCount=\""
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
    const auto* olap = cache.olap();
    if (olap && olap->present()) {
        // OLAP pivot: emit the olap cacheSource plus any calculatedMember nodes.
        // When a raw olapInfo subtree was loaded, reuse it so unmodeled children
        // (kpi, hierarchy metadata) stay byte-preserved.
        xml << "<cacheSource type=\"olap\"";
        if (olap->connectionId >= 0) xml << " connectionId=\"" << olap->connectionId << "\"";
        xml << ">";
        if (!olap->rawOlapInfoXml.empty()) {
            xml << olap->rawOlapInfoXml;
        } else {
            xml << "<olapInfo preserveFormatting=\"" << (olap->preserveFormatting ? 1 : 0) << "\"";
            if (!olap->localCube.empty()) xml << " localCube=\"" << xmlEscape(olap->localCube) << "\"";
            if (!olap->localConnection.empty()) xml << " localConnection=\"" << xmlEscape(olap->localConnection) << "\"";
            xml << "/>";
        }
        xml << "</cacheSource>";
        for (const auto& member : cache.calculatedMembers()) {
            if (!member.rawXml.empty()) {
                xml << member.rawXml;
            } else {
                xml << "<calculatedMember name=\"" << xmlEscape(member.name)
                    << "\" mdx=\"" << xmlEscape(member.mdx) << "\"";
                if (!member.memberName.empty()) xml << " memberName=\"" << xmlEscape(member.memberName) << "\"";
                if (member.hierarchy >= 0) xml << " hierarchy=\"" << member.hierarchy << "\"";
                if (!member.solveOrder.empty()) xml << " solveOrder=\"" << xmlEscape(member.solveOrder) << "\"";
                if (!member.set.empty()) xml << " set=\"" << xmlEscape(member.set) << "\"";
                xml << "/>";
            }
        }
        xml << "<cacheFields count=\"" << fieldCount << "\">";
    } else {
        xml << "<cacheSource type=\"worksheet\"><worksheetSource";
        if (!cache.sourceName().empty()) xml << " name=\"" << xmlEscape(cache.sourceName()) << "\"";
        else xml << " ref=\"" << xmlEscape(sourceRef) << "\"";
        if (!sourceSheet.empty()) xml << " sheet=\"" << xmlEscape(sourceSheet) << "\"";
        xml << "/></cacheSource><cacheFields count=\"" << fieldCount << "\">";
    }

    for (std::size_t i = 0; i < fieldCount; ++i) {
        const auto& items = sharedItems[i];
        bool containsString = false, containsNumber = false, containsBoolean = false;
        bool containsError = false, containsBlank = false, allNumbersInteger = true;
        double minValue = 0.0, maxValue = 0.0;
        bool hasNumericRange = false;
        std::set<PivotValueKind> nonBlankKinds;
        bool containsDate = false;
        for (const auto& item : items) {
            const auto parsed = item.kind == PivotValueKind::Number ? parsePivotValue(item.value) : ParsedPivotValue{item.kind, 0.0};
            if (item.kind != PivotValueKind::Blank) nonBlankKinds.insert(item.kind);
            switch (item.kind) {
                case PivotValueKind::Blank: containsBlank = true; break;
                case PivotValueKind::Number:
                    if (parsed.kind != PivotValueKind::Number)
                        throw std::invalid_argument("Typed Pivot numeric shared item is not a finite number: " + item.value);
                    containsNumber = true;
                    allNumbersInteger = allNumbersInteger && std::floor(parsed.number) == parsed.number;
                    if (!hasNumericRange) { minValue = maxValue = parsed.number; hasNumericRange = true; }
                    else { minValue = std::min(minValue, parsed.number); maxValue = std::max(maxValue, parsed.number); }
                    break;
                case PivotValueKind::Boolean: containsBoolean = true; break;
                case PivotValueKind::Error: containsError = true; break;
                case PivotValueKind::DateTime: containsDate = true; break;
                case PivotValueKind::String: containsString = true; break;
            }
        }

        xml << "<cacheField name=\"" << xmlEscape(cache.fields()[i]) << "\"";
        if (!cache.fieldCaption(i).empty()) xml << " caption=\"" << xmlEscape(cache.fieldCaption(i)) << "\"";
        if (cache.isCalculatedField(i)) xml << " formula=\"" << xmlEscape(cache.fieldFormula(i)) << "\"";
        if (!cache.fieldDatabaseField(i)) xml << " databaseField=\"0\"";
        xml << " numFmtId=\"" << cache.fieldNumberFormatId(i) << "\"><sharedItems";
        if (!containsString) xml << " containsString=\"0\"";
        if (containsNumber) xml << " containsNumber=\"1\"";
        if (containsNumber && allNumbersInteger) xml << " containsInteger=\"1\"";
        if (containsBlank) xml << " containsBlank=\"1\"";
        if (containsDate) xml << " containsDate=\"1\"";
        if (containsString || containsNumber || containsBoolean || containsError) xml << " containsNonDate=\"1\"";
        if (nonBlankKinds.size() > 1) xml << " containsMixedTypes=\"1\"";
        else if (!nonBlankKinds.empty()) xml << " containsMixedTypes=\"0\"";
        if (nonBlankKinds.size() <= 1 && (containsNumber || containsBoolean || containsError || containsDate))
            xml << " containsSemiMixedTypes=\"0\"";
        if (hasNumericRange) xml << " minValue=\"" << std::setprecision(15) << minValue
                                 << "\" maxValue=\"" << std::setprecision(15) << maxValue << "\"";
        const bool writeChildren = !pivotFieldIsPureData(pt, i);
        if (writeChildren) xml << " count=\"" << items.size() << "\"";

        if (!writeChildren || items.empty()) {
            xml << "/>";
        } else {
            xml << ">";
            for (const auto& item : items) writePivotValue(xml, item.value, item.kind);
            xml << "</sharedItems>";
        }
        if (const auto* group = cache.tryFieldGroup(i)) {
            xml << "<fieldGroup";
            if (group->parentField >= 0) xml << " par=\"" << group->parentField << "\"";
            if (group->baseField >= 0) xml << " base=\"" << group->baseField << "\"";
            const bool hasRange = !group->groupBy.empty() || group->startNumber.has_value() || group->endNumber.has_value()
                || group->interval.has_value() || !group->startDate.empty() || !group->endDate.empty()
                || !group->autoStart || !group->autoEnd;
            if (!hasRange && group->items.empty()) {
                xml << "/>";
            } else {
                xml << ">";
                if (hasRange) {
                    xml << "<rangePr";
                    if (!group->groupBy.empty()) xml << " groupBy=\"" << xmlEscape(group->groupBy) << "\"";
                    if (!group->autoStart) xml << " autoStart=\"0\"";
                    if (!group->autoEnd) xml << " autoEnd=\"0\"";
                    if (group->startNumber) xml << " startNum=\"" << std::setprecision(15) << *group->startNumber << "\"";
                    if (group->endNumber) xml << " endNum=\"" << std::setprecision(15) << *group->endNumber << "\"";
                    if (group->interval) xml << " groupInterval=\"" << std::setprecision(15) << *group->interval << "\"";
                    if (!group->startDate.empty()) xml << " startDate=\"" << xmlEscape(group->startDate) << "\"";
                    if (!group->endDate.empty()) xml << " endDate=\"" << xmlEscape(group->endDate) << "\"";
                    xml << "/>";
                }
                const auto itemCount = group->typedItems.empty() ? group->items.size() : group->typedItems.size();
                if (itemCount != 0) {
                    xml << "<groupItems count=\"" << itemCount << "\">";
                    if (group->typedItems.empty()) {
                        for (const auto& item : group->items) xml << "<s v=\"" << xmlEscape(item) << "\"/>";
                    } else {
                        for (const auto& item : group->typedItems) {
                            switch (item.kind) {
                                case xlpp::PivotCacheValueKind::Number: xml << "<n v=\"" << xmlEscape(item.value) << "\"/>"; break;
                                case xlpp::PivotCacheValueKind::DateTime: xml << "<d v=\"" << xmlEscape(item.value) << "\"/>"; break;
                                case xlpp::PivotCacheValueKind::Boolean: xml << "<b v=\"" << (item.value == "1" ? "1" : "0") << "\"/>"; break;
                                case xlpp::PivotCacheValueKind::Error: xml << "<e v=\"" << xmlEscape(item.value) << "\"/>"; break;
                                case xlpp::PivotCacheValueKind::Missing: xml << "<m/>"; break;
                                case xlpp::PivotCacheValueKind::String: default: xml << "<s v=\"" << xmlEscape(item.value) << "\"/>"; break;
                            }
                        }
                    }
                    xml << "</groupItems>";
                }
                xml << "</fieldGroup>";
            }
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
    for (std::size_t rowIndex = 0; rowIndex < cache.records().size(); ++rowIndex) {
        const auto& record = cache.records()[rowIndex];
        if (record.size() != cache.fields().size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        xml << "<r>";
        for (std::size_t fieldIndex = 0; fieldIndex < record.size(); ++fieldIndex) {
            const auto kind = recordValueKind(cache, rowIndex, fieldIndex, record[fieldIndex]);
            if (pivotFieldIsPureData(pt, fieldIndex)) {
                writePivotValue(xml, record[fieldIndex], kind);
            } else {
                const auto& items = sharedItems[fieldIndex];
                PivotSharedItem candidate{record[fieldIndex], kind};
                const auto item = std::find_if(items.begin(), items.end(), [&](const auto& actual) { return samePivotSharedItem(actual, candidate); });
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

struct PivotCellValue {
    std::string text;
    xlpp::PivotCacheValueKind kind{xlpp::PivotCacheValueKind::Missing};
};

PivotCellValue pivotCellValue(const xlpp::Cell* cell) {
    if (!cell || !cell->hasValue()) return {};
    const auto& value = cell->value();
    if (const auto* text = std::get_if<std::string>(&value)) return {*text, xlpp::PivotCacheValueKind::String};
    if (const auto* number = std::get_if<double>(&value)) {
        std::ostringstream out;
        out << std::setprecision(15) << *number;
        return {out.str(), xlpp::PivotCacheValueKind::Number};
    }
    if (const auto* boolean = std::get_if<bool>(&value))
        return {*boolean ? "true" : "false", xlpp::PivotCacheValueKind::Boolean};
    if (const auto* error = std::get_if<xlpp::CellError>(&value))
        return {xlpp::toString(*error), xlpp::PivotCacheValueKind::Error};
    if (const auto* date = std::get_if<xlpp::DateTime>(&value))
        return {xlpp::toIso8601(*date), xlpp::PivotCacheValueKind::DateTime};
    return {};
}

std::string pivotCellText(const xlpp::Cell* cell) { return pivotCellValue(cell).text; }

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

    const bool brokenSourceReference = result.cache().sourceData().find("#REF!") != std::string::npos;
    if (brokenSourceReference) {
        // Structural deletion can legitimately leave an existing PivotCache
        // with a broken worksheet source while its cached schema/data remains
        // usable for display. Do not attempt to dereference #REF! during save.
        // Excel/Calc can surface the broken source and a later refresh can be
        // repaired by the caller. A schema-less generated Pivot still cannot
        // be serialized safely.
        if (result.cache().fields().empty())
            throw std::invalid_argument("Pivot source is #REF! and the cache has no field schema");
        result.cache().setRefreshOnLoad(true);
        return result;
    }

    if (result.cache().fields().empty() || static_cast<const xlpp::PivotCache&>(result.cache()).records().empty()) {
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

        if (static_cast<const xlpp::PivotCache&>(result.cache()).records().empty() && parsed.first.row < parsed.last.row) {
            for (std::size_t row = parsed.first.row + 1; row <= parsed.last.row; ++row) {
                std::vector<std::string> record;
                std::vector<xlpp::PivotCacheValueKind> kinds;
                record.reserve(width);
                kinds.reserve(width);
                for (std::size_t column = parsed.first.column; column <= parsed.last.column; ++column) {
                    auto value = pivotCellValue(sourceSheet->tryCell(row, column));
                    record.push_back(std::move(value.text));
                    kinds.push_back(value.kind);
                }
                result.cache().addTypedRecord(std::move(record), std::move(kinds));
            }
        }
    }
    return result;
}

bool pivotCachesEquivalent(const xlpp::PivotCache& left, const xlpp::PivotCache& right) {
    if (left.sourceData() != right.sourceData() || left.sourceName() != right.sourceName()
        || left.refreshOnLoad() != right.refreshOnLoad() || left.saveData() != right.saveData()
        || left.enableRefresh() != right.enableRefresh() || left.missingItemsLimit() != right.missingItemsLimit()
        || left.fields() != right.fields() || left.records() != right.records()) return false;
    if (left.hasTypedRecordKinds() != right.hasTypedRecordKinds()) return false;
    if (left.hasTypedRecordKinds() && left.recordKinds() != right.recordKinds()) return false;
    for (std::size_t index = 0; index < left.fields().size(); ++index)
        if (left.fieldFormula(index) != right.fieldFormula(index)) return false;
    for (std::size_t index = 0; index < left.fields().size(); ++index) {
        const auto* a = left.tryFieldGroup(index);
        const auto* b = right.tryFieldGroup(index);
        if ((a == nullptr) != (b == nullptr)) return false;
        if (!a) continue;
        if (a->parentField != b->parentField || a->baseField != b->baseField || a->groupBy != b->groupBy
            || a->autoStart != b->autoStart || a->autoEnd != b->autoEnd || a->startNumber != b->startNumber
            || a->endNumber != b->endNumber || a->interval != b->interval || a->startDate != b->startDate
            || a->endDate != b->endDate || a->items != b->items) return false;
    }
    return true;
}
} // namespace internal
} // namespace xlpp

