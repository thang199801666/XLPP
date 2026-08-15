#include <XLPP/Workbook/Workbook.h>
#include "../XML/XmlUtilities.h"

#include <algorithm>
#include <array>
#include <cctype>
#include <cstdint>
#include <optional>
#include <sstream>
#include <stdexcept>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

namespace xlpp {
using internal::xmlEscape;

namespace {
std::string resolvePackagePart(const std::string& basePart, std::string relativeTarget) {
    if (relativeTarget.empty()) return {};
    const bool absoluteTarget = relativeTarget.front() == '/';
    if (absoluteTarget) relativeTarget.erase(relativeTarget.begin());
    std::vector<std::string> segments;
    const auto slash = basePart.find_last_of('/');
    std::string combined = absoluteTarget
        ? relativeTarget
        : (slash == std::string::npos ? std::string{} : basePart.substr(0, slash + 1)) + relativeTarget;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = combined.substr(begin, end == std::string::npos ? std::string::npos : end - begin);
        if (segment == "..") { if (!segments.empty()) segments.pop_back(); }
        else if (!segment.empty() && segment != ".") segments.push_back(segment);
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    std::ostringstream result;
    for (std::size_t i = 0; i < segments.size(); ++i) {
        if (i) result << '/';
        result << segments[i];
    }
    return result.str();
}

struct PivotXmlRange {
    std::size_t begin{std::string::npos};
    std::size_t end{std::string::npos}; // exclusive
    explicit operator bool() const noexcept { return begin != std::string::npos && end != std::string::npos && end >= begin; }
};

bool pivotXmlTagBoundary(char value) {
    return value == ' ' || value == '\t' || value == '\r' || value == '\n' || value == '>' || value == '/';
}

PivotXmlRange pivotBalancedElementRange(const std::string& xml, std::string_view localName,
                                        std::size_t searchBegin = 0,
                                        std::size_t searchEnd = std::string::npos) {
    const auto openToken = "<" + std::string(localName);
    const auto closeToken = "</" + std::string(localName) + ">";
    if (searchEnd == std::string::npos || searchEnd > xml.size()) searchEnd = xml.size();
    auto begin = searchBegin;
    while ((begin = xml.find(openToken, begin)) != std::string::npos && begin < searchEnd) {
        const auto afterName = begin + openToken.size();
        if (afterName < xml.size() && pivotXmlTagBoundary(xml[afterName])) break;
        begin = afterName;
    }
    if (begin == std::string::npos || begin >= searchEnd) return {};
    const auto headEnd = xml.find('>', begin + openToken.size());
    if (headEnd == std::string::npos || headEnd >= searchEnd) return {};
    if (headEnd > begin && xml[headEnd - 1] == '/') return {begin, headEnd + 1};
    std::size_t depth = 1;
    std::size_t pos = headEnd + 1;
    while (depth != 0 && pos < searchEnd) {
        auto nextOpen = xml.find(openToken, pos);
        while (nextOpen != std::string::npos && nextOpen < searchEnd) {
            const auto afterName = nextOpen + openToken.size();
            if (afterName < xml.size() && pivotXmlTagBoundary(xml[afterName])) break;
            nextOpen = xml.find(openToken, afterName);
        }
        const auto nextClose = xml.find(closeToken, pos);
        if (nextClose == std::string::npos || nextClose >= searchEnd) return {};
        if (nextOpen != std::string::npos && nextOpen < nextClose && nextOpen < searchEnd) {
            const auto nestedHeadEnd = xml.find('>', nextOpen + openToken.size());
            if (nestedHeadEnd == std::string::npos || nestedHeadEnd >= searchEnd) return {};
            if (nestedHeadEnd == nextOpen || xml[nestedHeadEnd - 1] != '/') ++depth;
            pos = nestedHeadEnd + 1;
        } else {
            --depth;
            pos = nextClose + closeToken.size();
        }
    }
    return depth == 0 ? PivotXmlRange{begin, pos} : PivotXmlRange{};
}

PivotXmlRange pivotNthChildRange(const std::string& xml, const PivotXmlRange& container,
                                 std::string_view localName, std::size_t index) {
    if (!container) return {};
    const auto headEnd = xml.find('>', container.begin);
    if (headEnd == std::string::npos || headEnd + 1 >= container.end) return {};
    std::size_t pos = headEnd + 1;
    for (std::size_t i = 0; i <= index; ++i) {
        const auto child = pivotBalancedElementRange(xml, localName, pos, container.end);
        if (!child) return {};
        if (i == index) return child;
        pos = child.end;
    }
    return {};
}

PivotXmlRange pivotNthCacheValueRange(const std::string& xml, const PivotXmlRange& record,
                                      std::size_t index) {
    if (!record) return {};
    const auto headEnd = xml.find('>', record.begin);
    if (headEnd == std::string::npos || headEnd + 1 >= record.end) return {};
    static constexpr std::array<std::string_view, 7> kValueTags{"m", "n", "b", "e", "s", "d", "x"};
    std::size_t pos = headEnd + 1;
    for (std::size_t current = 0; current <= index; ++current) {
        PivotXmlRange best;
        for (const auto tag : kValueTags) {
            const auto candidate = pivotBalancedElementRange(xml, tag, pos, record.end);
            if (candidate && (!best || candidate.begin < best.begin)) best = candidate;
        }
        if (!best) return {};
        if (current == index) return best;
        pos = best.end;
    }
    return {};
}

std::string pivotPatchAttribute(std::string head, const std::string& name,
                                const std::optional<std::string>& value) {
    const auto key = name + "=\"";
    const auto attrBegin = head.find(key);
    if (!value) return head;
    if (value->empty()) {
        if (attrBegin != std::string::npos) {
            auto eraseBegin = attrBegin;
            while (eraseBegin > 0 && std::isspace(static_cast<unsigned char>(head[eraseBegin - 1]))) --eraseBegin;
            const auto valueBegin = attrBegin + key.size();
            const auto valueEnd = head.find('\"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed Pivot XML attribute");
            head.erase(eraseBegin, valueEnd - eraseBegin + 1);
        }
        return head;
    }
    if (attrBegin != std::string::npos) {
        const auto valueBegin = attrBegin + key.size();
        const auto valueEnd = head.find('\"', valueBegin);
        if (valueEnd == std::string::npos) throw std::runtime_error("Malformed Pivot XML attribute");
        head.replace(valueBegin, valueEnd - valueBegin, *value);
        return head;
    }
    auto insertPos = head.size() - 1;
    if (insertPos > 0 && head[insertPos - 1] == '/') --insertPos;
    head.insert(insertPos, " " + name + "=\"" + *value + "\"");
    return head;
}

} // namespace

bool Workbook::updateImportedPivotCacheOptions(const std::string& worksheetName,
                                               const std::string& pivotTableName,
                                               const PivotCacheOptionsPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
        return sheet.name() == worksheetName;
    });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) {
        return pivot.name() == pivotTableName;
    });
    if (pivotIt == sheetIt->pivotTables_.end()) return false;
    const auto cachePart = pivotIt->cache().sharedCacheKey();
    if (cachePart.empty() || cachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) return false;

    auto partIt = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& part) {
        return part.name == cachePart;
    });
    if (partIt == preservedParts_.end()) return false;

    auto xml = partIt->data;
    const auto rootBegin = xml.find("<pivotCacheDefinition");
    if (rootBegin == std::string::npos) return false;
    const auto rootEnd = xml.find('>', rootBegin);
    if (rootEnd == std::string::npos) return false;
    auto root = xml.substr(rootBegin, rootEnd - rootBegin + 1);
    const auto setAttribute = [&](const std::string& name, const std::string& value) {
        const auto key = name + "=\"";
        const auto begin = root.find(key);
        if (begin != std::string::npos) {
            const auto valueBegin = begin + key.size();
            const auto valueEnd = root.find('"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed pivot cache root attribute");
            root.replace(valueBegin, valueEnd - valueBegin, value);
            return;
        }
        auto insertPos = root.size() - 1;
        if (insertPos > 0 && root[insertPos - 1] == '/') --insertPos;
        root.insert(insertPos, " " + name + "=\"" + value + "\"");
    };
    if (patch.refreshOnLoad) setAttribute("refreshOnLoad", *patch.refreshOnLoad ? "1" : "0");
    if (patch.saveData) setAttribute("saveData", *patch.saveData ? "1" : "0");
    if (patch.enableRefresh) setAttribute("enableRefresh", *patch.enableRefresh ? "1" : "0");
    if (patch.missingItemsLimit) setAttribute("missingItemsLimit", std::to_string(*patch.missingItemsLimit));
    xml.replace(rootBegin, rootEnd - rootBegin + 1, root);
    partIt->data = std::move(xml);

    // Keep every loaded model sharing this physical cache coherent without
    // flipping Worksheet::pivotsDirty_: the package XML is selectively patched.
    for (auto& sheet : sheets_) {
        for (auto& pivot : sheet.pivotTables_) {
            if (pivot.cache().sharedCacheKey() != cachePart) continue;
            if (patch.refreshOnLoad) pivot.cache().setRefreshOnLoad(*patch.refreshOnLoad);
            if (patch.saveData) pivot.cache().setSaveData(*patch.saveData);
            if (patch.enableRefresh) pivot.cache().setEnableRefresh(*patch.enableRefresh);
            if (patch.missingItemsLimit) pivot.cache().setMissingItemsLimit(*patch.missingItemsLimit);
        }
    }
    return true;
}

bool Workbook::updateImportedPivotCacheField(const std::string& worksheetName,
                                             const std::string& pivotTableName,
                                             std::size_t fieldIndex,
                                             const PivotCacheFieldPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
        return sheet.name() == worksheetName;
    });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) {
        return pivot.name() == pivotTableName;
    });
    if (pivotIt == sheetIt->pivotTables_.end() || fieldIndex >= pivotIt->cache().fields().size()) return false;
    const auto cachePart = pivotIt->cache().sharedCacheKey();
    if (cachePart.empty() || cachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) return false;
    if (patch.name && patch.name->empty()) throw std::invalid_argument("Pivot cache field name cannot be empty");
    if (patch.numberFormatId && *patch.numberFormatId < 0)
        throw std::invalid_argument("Pivot cache field number format ID cannot be negative");

    auto partIt = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& part) {
        return part.name == cachePart;
    });
    if (partIt == preservedParts_.end()) return false;
    auto xml = partIt->data;
    std::size_t search = 0;
    std::size_t begin = std::string::npos;
    for (std::size_t i = 0; i <= fieldIndex; ++i) {
        while (true) {
            begin = xml.find("<cacheField", search);
            if (begin == std::string::npos) return false;
            const auto after = begin + std::string("<cacheField").size();
            if (after >= xml.size() || std::isspace(static_cast<unsigned char>(xml[after])) || xml[after] == '>' || xml[after] == '/') break;
            search = after;
        }
        if (i != fieldIndex) search = begin + 1;
    }
    const auto end = xml.find('>', begin);
    if (end == std::string::npos) return false;
    auto head = xml.substr(begin, end - begin + 1);
    const auto setAttribute = [&](const std::string& name, const std::optional<std::string>& value) {
        const auto key = name + "=\"";
        const auto attrBegin = head.find(key);
        if (!value || value->empty()) {
            if (attrBegin != std::string::npos) {
                auto eraseBegin = attrBegin;
                while (eraseBegin > 0 && std::isspace(static_cast<unsigned char>(head[eraseBegin - 1]))) --eraseBegin;
                const auto valueBegin = attrBegin + key.size();
                const auto valueEnd = head.find('"', valueBegin);
                if (valueEnd == std::string::npos) throw std::runtime_error("Malformed pivot cache field attribute");
                head.erase(eraseBegin, valueEnd - eraseBegin + 1);
            }
            return;
        }
        if (attrBegin != std::string::npos) {
            const auto valueBegin = attrBegin + key.size();
            const auto valueEnd = head.find('"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed pivot cache field attribute");
            head.replace(valueBegin, valueEnd - valueBegin, *value);
            return;
        }
        auto insertPos = head.size() - 1;
        if (insertPos > 0 && head[insertPos - 1] == '/') --insertPos;
        head.insert(insertPos, " " + name + "=\"" + *value + "\"");
    };
    if (patch.name) setAttribute("name", xmlEscape(*patch.name));
    if (patch.caption) setAttribute("caption", patch.caption->empty() ? std::optional<std::string>{} : std::optional<std::string>{xmlEscape(*patch.caption)});
    if (patch.formula) setAttribute("formula", patch.formula->empty() ? std::optional<std::string>{} : std::optional<std::string>{xmlEscape(*patch.formula)});
    if (patch.numberFormatId) setAttribute("numFmtId", std::to_string(*patch.numberFormatId));
    if (patch.databaseField) setAttribute("databaseField", *patch.databaseField ? "1" : "0");
    xml.replace(begin, end - begin + 1, head);
    partIt->data = std::move(xml);

    for (auto& sheet : sheets_) {
        for (auto& pivot : sheet.pivotTables_) {
            if (pivot.cache().sharedCacheKey() != cachePart || fieldIndex >= pivot.cache().fields().size()) continue;
            if (patch.name) {
                const auto oldName = pivot.cache().fields()[fieldIndex];
                pivot.cache().fields()[fieldIndex] = *patch.name;
                const auto syncName = [&](auto& fields) {
                    for (auto& field : fields) {
                        if (field.fieldIndex() == static_cast<int>(fieldIndex) || field.name() == oldName) field.setName(*patch.name);
                    }
                };
                syncName(pivot.rowFields());
                syncName(pivot.columnFields());
                syncName(pivot.pageFields());
                for (auto& field : pivot.dataFields()) {
                    if (field.fieldIndex() == static_cast<int>(fieldIndex) || field.name() == oldName) field.setName(*patch.name);
                }
            }
            if (patch.caption) pivot.cache().setFieldCaption(fieldIndex, *patch.caption);
            if (patch.formula) pivot.cache().setFieldFormula(fieldIndex, *patch.formula);
            if (patch.numberFormatId) pivot.cache().setFieldNumberFormatId(fieldIndex, *patch.numberFormatId);
            if (patch.databaseField) pivot.cache().setFieldDatabaseField(fieldIndex, *patch.databaseField);
        }
    }
    return true;
}

bool Workbook::updateImportedPivotFieldItem(const std::string& worksheetName,
                                            const std::string& pivotTableName,
                                            std::size_t fieldIndex,
                                            std::size_t itemIndex,
                                            const PivotFieldItemPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return sheet.name() == worksheetName; });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) { return pivot.name() == pivotTableName; });
    if (pivotIt == sheetIt->pivotTables_.end()) return false;
    if (patch.cacheIndex && *patch.cacheIndex < -1) throw std::invalid_argument("Pivot field item cache index cannot be less than -1");

    auto findPivotPart = [&]() -> PreservedPart* {
        std::vector<std::string> candidateParts;
        for (std::size_t i = 0; i < sourceSheetNames_.size() && i < sourceSheetParts_.size(); ++i) {
            if (sourceSheetNames_[i] != worksheetName) continue;
            const auto& sourcePart = sourceSheetParts_[i];
            for (const auto& relationship : preservedRelationships_) {
                if (relationship.sourcePart != sourcePart || relationship.targetMode == "External"
                    || relationship.type.find("/pivotTable") == std::string::npos) continue;
                candidateParts.push_back(resolvePackagePart(sourcePart, relationship.target));
            }
        }
        if (candidateParts.empty()) {
            for (const auto& part : preservedParts_)
                if (part.name.rfind("xl/pivotTables/pivotTable", 0) == 0 && part.name.ends_with(".xml")) candidateParts.push_back(part.name);
        }
        for (const auto& candidate : candidateParts) {
            auto part = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& value) { return value.name == candidate; });
            if (part == preservedParts_.end()) continue;
            const auto roots = internal::tags(part->data, "pivotTableDefinition");
            if (!roots.empty() && internal::attribute(roots.front(), "name") == pivotTableName) return &*part;
        }
        return nullptr;
    };
    auto* part = findPivotPart();
    if (!part) return false;
    auto xml = part->data;
    const auto fields = pivotBalancedElementRange(xml, "pivotFields");
    const auto field = pivotNthChildRange(xml, fields, "pivotField", fieldIndex);
    if (!field) return false;
    const auto items = pivotBalancedElementRange(xml, "items", field.begin, field.end);
    const auto item = pivotNthChildRange(xml, items, "item", itemIndex);
    if (!item) return false;
    const auto headEnd = xml.find('>', item.begin);
    if (headEnd == std::string::npos || headEnd >= item.end) return false;
    auto head = xml.substr(item.begin, headEnd - item.begin + 1);
    if (patch.cacheIndex) head = pivotPatchAttribute(std::move(head), "x", *patch.cacheIndex < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.cacheIndex)});
    if (patch.type) head = pivotPatchAttribute(std::move(head), "t", patch.type->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.type)});
    if (patch.caption) head = pivotPatchAttribute(std::move(head), "n", patch.caption->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.caption)});
    if (patch.hidden) head = pivotPatchAttribute(std::move(head), "h", *patch.hidden ? std::optional<std::string>{"1"} : std::optional<std::string>{std::string{}});
    if (patch.showDetails) head = pivotPatchAttribute(std::move(head), "sd", *patch.showDetails ? std::optional<std::string>{std::string{}} : std::optional<std::string>{"0"});
    if (patch.formula) head = pivotPatchAttribute(std::move(head), "f", *patch.formula ? std::optional<std::string>{"1"} : std::optional<std::string>{std::string{}});
    if (patch.missing) head = pivotPatchAttribute(std::move(head), "m", *patch.missing ? std::optional<std::string>{"1"} : std::optional<std::string>{std::string{}});
    xml.replace(item.begin, headEnd - item.begin + 1, head);
    part->data = std::move(xml);

    const auto syncField = [&](PivotField& modelField) {
        if (modelField.fieldIndex() != static_cast<int>(fieldIndex) || itemIndex >= modelField.items().size()) return;
        auto& modelItem = modelField.items()[itemIndex];
        if (patch.cacheIndex) modelItem.cacheIndex = *patch.cacheIndex;
        if (patch.type) modelItem.type = *patch.type;
        if (patch.caption) modelItem.caption = *patch.caption;
        if (patch.hidden) modelItem.hidden = *patch.hidden;
        if (patch.showDetails) modelItem.showDetails = *patch.showDetails;
        if (patch.formula) modelItem.formula = *patch.formula;
        if (patch.missing) modelItem.missing = *patch.missing;
        modelField.hiddenItems().clear();
        for (const auto& candidate : modelField.items())
            if (candidate.hidden && candidate.cacheIndex >= 0) modelField.hiddenItems().push_back(candidate.cacheIndex);
        const auto current = pivotBalancedElementRange(part->data, "pivotFields");
        const auto currentField = pivotNthChildRange(part->data, current, "pivotField", fieldIndex);
        const auto currentItems = pivotBalancedElementRange(part->data, "items", currentField.begin, currentField.end);
        const auto currentItem = pivotNthChildRange(part->data, currentItems, "item", itemIndex);
        if (currentItem) modelItem.rawXml = part->data.substr(currentItem.begin, currentItem.end - currentItem.begin);
    };
    for (auto& modelField : pivotIt->rowFields()) syncField(modelField);
    for (auto& modelField : pivotIt->columnFields()) syncField(modelField);
    for (auto& modelField : pivotIt->pageFields()) syncField(modelField);
    return true;
}

bool Workbook::updateImportedPivotFilter(const std::string& worksheetName,
                                         const std::string& pivotTableName,
                                         std::size_t filterIndex,
                                         const PivotFilterPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return sheet.name() == worksheetName; });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) { return pivot.name() == pivotTableName; });
    if (pivotIt == sheetIt->pivotTables_.end() || filterIndex >= pivotIt->filters().size()) return false;
    if (patch.fieldIndex && *patch.fieldIndex < 0) throw std::invalid_argument("Pivot filter field index cannot be negative");
    if (patch.type && patch.type->empty()) throw std::invalid_argument("Pivot filter type cannot be empty");

    auto findPivotPart = [&]() -> PreservedPart* {
        std::vector<std::string> candidateParts;
        for (std::size_t i = 0; i < sourceSheetNames_.size() && i < sourceSheetParts_.size(); ++i) {
            if (sourceSheetNames_[i] != worksheetName) continue;
            const auto& sourcePart = sourceSheetParts_[i];
            for (const auto& relationship : preservedRelationships_) {
                if (relationship.sourcePart != sourcePart || relationship.targetMode == "External"
                    || relationship.type.find("/pivotTable") == std::string::npos) continue;
                candidateParts.push_back(resolvePackagePart(sourcePart, relationship.target));
            }
        }
        if (candidateParts.empty()) {
            for (const auto& part : preservedParts_)
                if (part.name.rfind("xl/pivotTables/pivotTable", 0) == 0 && part.name.ends_with(".xml")) candidateParts.push_back(part.name);
        }
        for (const auto& candidate : candidateParts) {
            auto part = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& value) { return value.name == candidate; });
            if (part == preservedParts_.end()) continue;
            const auto roots = internal::tags(part->data, "pivotTableDefinition");
            if (!roots.empty() && internal::attribute(roots.front(), "name") == pivotTableName) return &*part;
        }
        return nullptr;
    };
    auto* part = findPivotPart();
    if (!part) return false;
    auto xml = part->data;
    const auto filters = pivotBalancedElementRange(xml, "filters");
    const auto filter = pivotNthChildRange(xml, filters, "filter", filterIndex);
    if (!filter) return false;
    const auto headEnd = xml.find('>', filter.begin);
    if (headEnd == std::string::npos || headEnd >= filter.end) return false;
    auto head = xml.substr(filter.begin, headEnd - filter.begin + 1);
    if (patch.fieldIndex) head = pivotPatchAttribute(std::move(head), "fld", std::to_string(*patch.fieldIndex));
    if (patch.type) head = pivotPatchAttribute(std::move(head), "type", xmlEscape(*patch.type));
    if (patch.id) head = pivotPatchAttribute(std::move(head), "id", std::to_string(*patch.id));
    if (patch.evaluationOrder) head = pivotPatchAttribute(std::move(head), "evalOrder", *patch.evaluationOrder == 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.evaluationOrder)});
    if (patch.measureField) head = pivotPatchAttribute(std::move(head), "iMeasureFld", *patch.measureField < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.measureField)});
    if (patch.measureHierarchy) head = pivotPatchAttribute(std::move(head), "iMeasureHier", *patch.measureHierarchy < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.measureHierarchy)});
    if (patch.memberPropertyField) head = pivotPatchAttribute(std::move(head), "mpFld", *patch.memberPropertyField < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.memberPropertyField)});
    if (patch.name) head = pivotPatchAttribute(std::move(head), "name", patch.name->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.name)});
    if (patch.description) head = pivotPatchAttribute(std::move(head), "description", patch.description->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.description)});
    if (patch.stringValue1) head = pivotPatchAttribute(std::move(head), "stringValue1", patch.stringValue1->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.stringValue1)});
    if (patch.stringValue2) head = pivotPatchAttribute(std::move(head), "stringValue2", patch.stringValue2->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.stringValue2)});

    std::string replacement;
    const bool wasSelfClosing = head.size() >= 2 && head[head.size() - 2] == '/';
    if (patch.autoFilterXml) {
        std::string body;
        if (!wasSelfClosing) body = part->data.substr(headEnd + 1, filter.end - (headEnd + 1) - std::string("</filter>").size());
        const auto autoRange = pivotBalancedElementRange(body, "autoFilter");
        if (autoRange) body.replace(autoRange.begin, autoRange.end - autoRange.begin, *patch.autoFilterXml);
        else if (!patch.autoFilterXml->empty()) body += *patch.autoFilterXml;
        if (patch.autoFilterXml->empty() && autoRange) body.erase(autoRange.begin, autoRange.end - autoRange.begin);
        if (body.empty()) {
            if (!head.empty() && head.back() == '>') {
                if (!wasSelfClosing) head.insert(head.size() - 1, "/");
            }
            replacement = head;
        } else {
            if (wasSelfClosing) head.erase(head.size() - 2, 1);
            replacement = head + body + "</filter>";
        }
    } else {
        replacement = head + part->data.substr(headEnd + 1, filter.end - (headEnd + 1));
    }
    xml.replace(filter.begin, filter.end - filter.begin, replacement);
    part->data = std::move(xml);

    auto& model = pivotIt->filters()[filterIndex];
    if (patch.fieldIndex) model.fieldIndex = *patch.fieldIndex;
    if (patch.type) model.type = *patch.type;
    if (patch.id) model.id = *patch.id;
    if (patch.evaluationOrder) model.evaluationOrder = *patch.evaluationOrder;
    if (patch.measureField) model.measureField = *patch.measureField;
    if (patch.measureHierarchy) model.measureHierarchy = *patch.measureHierarchy;
    if (patch.memberPropertyField) model.memberPropertyField = *patch.memberPropertyField;
    if (patch.name) model.name = *patch.name;
    if (patch.description) model.description = *patch.description;
    if (patch.stringValue1) model.stringValue1 = *patch.stringValue1;
    if (patch.stringValue2) model.stringValue2 = *patch.stringValue2;
    if (patch.autoFilterXml) model.autoFilterXml = *patch.autoFilterXml;
    return true;
}

bool Workbook::updateImportedPivotDataField(const std::string& worksheetName,
                                            const std::string& pivotTableName,
                                            std::size_t dataFieldIndex,
                                            const PivotDataFieldPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return sheet.name() == worksheetName; });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) { return pivot.name() == pivotTableName; });
    if (pivotIt == sheetIt->pivotTables_.end() || dataFieldIndex >= pivotIt->dataFields().size()) return false;
    if (patch.fieldIndex && (*patch.fieldIndex < 0 || static_cast<std::size_t>(*patch.fieldIndex) >= pivotIt->cache().fields().size()))
        throw std::invalid_argument("Pivot data field index is out of cache-field range");
    if (patch.numberFormatId && *patch.numberFormatId < 0) throw std::invalid_argument("Pivot data field number format ID cannot be negative");
    if (patch.subtotal) { PivotFieldReference probe; probe.setSubtotal(*patch.subtotal); }
    if (patch.showDataAs) { PivotFieldReference probe; probe.setShowDataAs(*patch.showDataAs); }

    auto findPivotPart = [&]() -> PreservedPart* {
        std::vector<std::string> candidateParts;
        for (std::size_t i = 0; i < sourceSheetNames_.size() && i < sourceSheetParts_.size(); ++i) {
            if (sourceSheetNames_[i] != worksheetName) continue;
            const auto& sourcePart = sourceSheetParts_[i];
            for (const auto& relationship : preservedRelationships_) {
                if (relationship.sourcePart != sourcePart || relationship.targetMode == "External"
                    || relationship.type.find("/pivotTable") == std::string::npos) continue;
                candidateParts.push_back(resolvePackagePart(sourcePart, relationship.target));
            }
        }
        if (candidateParts.empty()) {
            for (const auto& part : preservedParts_)
                if (part.name.rfind("xl/pivotTables/pivotTable", 0) == 0 && part.name.ends_with(".xml")) candidateParts.push_back(part.name);
        }
        for (const auto& candidate : candidateParts) {
            auto part = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& value) { return value.name == candidate; });
            if (part == preservedParts_.end()) continue;
            const auto roots = internal::tags(part->data, "pivotTableDefinition");
            if (!roots.empty() && internal::attribute(roots.front(), "name") == pivotTableName) return &*part;
        }
        return nullptr;
    };
    auto* part = findPivotPart();
    if (!part) return false;
    auto xml = part->data;
    const auto container = pivotBalancedElementRange(xml, "dataFields");
    const auto field = pivotNthChildRange(xml, container, "dataField", dataFieldIndex);
    if (!field) return false;
    const auto headEnd = xml.find('>', field.begin);
    if (headEnd == std::string::npos || headEnd >= field.end) return false;
    auto head = xml.substr(field.begin, headEnd - field.begin + 1);
    if (patch.fieldIndex) head = pivotPatchAttribute(std::move(head), "fld", std::to_string(*patch.fieldIndex));
    if (patch.name) head = pivotPatchAttribute(std::move(head), "name", patch.name->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.name)});
    if (patch.subtotal) head = pivotPatchAttribute(std::move(head), "subtotal", *patch.subtotal == "sum" ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.subtotal)});
    if (patch.showDataAs) head = pivotPatchAttribute(std::move(head), "showDataAs", *patch.showDataAs == "normal" ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.showDataAs)});
    if (patch.baseField) head = pivotPatchAttribute(std::move(head), "baseField", *patch.baseField < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.baseField)});
    if (patch.baseItem) head = pivotPatchAttribute(std::move(head), "baseItem", *patch.baseItem < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.baseItem)});
    if (patch.numberFormatId) head = pivotPatchAttribute(std::move(head), "numFmtId", *patch.numberFormatId == 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.numberFormatId)});
    xml.replace(field.begin, headEnd - field.begin + 1, head);
    part->data = std::move(xml);

    auto& model = pivotIt->dataFields()[dataFieldIndex];
    if (patch.fieldIndex) {
        model.setFieldIndex(*patch.fieldIndex);
        model.setName(pivotIt->cache().fields()[static_cast<std::size_t>(*patch.fieldIndex)]);
    }
    if (patch.name) model.setDisplayName(*patch.name);
    if (patch.subtotal) model.setSubtotal(*patch.subtotal);
    if (patch.showDataAs) model.setShowDataAs(*patch.showDataAs);
    if (patch.baseField) model.setBaseField(*patch.baseField);
    if (patch.baseItem) model.setBaseItem(*patch.baseItem);
    if (patch.numberFormatId) model.setNumberFormatId(*patch.numberFormatId);
    return true;
}

bool Workbook::updateImportedPivotPageField(const std::string& worksheetName,
                                            const std::string& pivotTableName,
                                            std::size_t pageFieldIndex,
                                            const PivotPageFieldPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) { return sheet.name() == worksheetName; });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) { return pivot.name() == pivotTableName; });
    if (pivotIt == sheetIt->pivotTables_.end() || pageFieldIndex >= pivotIt->pageFieldSettings().size()) return false;
    if (patch.fieldIndex && (*patch.fieldIndex < 0 || static_cast<std::size_t>(*patch.fieldIndex) >= pivotIt->cache().fields().size()))
        throw std::invalid_argument("Pivot page field index is out of cache-field range");

    auto findPivotPart = [&]() -> PreservedPart* {
        std::vector<std::string> candidateParts;
        for (std::size_t i = 0; i < sourceSheetNames_.size() && i < sourceSheetParts_.size(); ++i) {
            if (sourceSheetNames_[i] != worksheetName) continue;
            const auto& sourcePart = sourceSheetParts_[i];
            for (const auto& relationship : preservedRelationships_) {
                if (relationship.sourcePart != sourcePart || relationship.targetMode == "External"
                    || relationship.type.find("/pivotTable") == std::string::npos) continue;
                candidateParts.push_back(resolvePackagePart(sourcePart, relationship.target));
            }
        }
        if (candidateParts.empty()) {
            for (const auto& part : preservedParts_)
                if (part.name.rfind("xl/pivotTables/pivotTable", 0) == 0 && part.name.ends_with(".xml")) candidateParts.push_back(part.name);
        }
        for (const auto& candidate : candidateParts) {
            auto part = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& value) { return value.name == candidate; });
            if (part == preservedParts_.end()) continue;
            const auto roots = internal::tags(part->data, "pivotTableDefinition");
            if (!roots.empty() && internal::attribute(roots.front(), "name") == pivotTableName) return &*part;
        }
        return nullptr;
    };
    auto* part = findPivotPart();
    if (!part) return false;
    auto xml = part->data;
    const auto container = pivotBalancedElementRange(xml, "pageFields");
    const auto field = pivotNthChildRange(xml, container, "pageField", pageFieldIndex);
    if (!field) return false;
    const auto headEnd = xml.find('>', field.begin);
    if (headEnd == std::string::npos || headEnd >= field.end) return false;
    auto head = xml.substr(field.begin, headEnd - field.begin + 1);
    if (patch.fieldIndex) head = pivotPatchAttribute(std::move(head), "fld", std::to_string(*patch.fieldIndex));
    if (patch.item) head = pivotPatchAttribute(std::move(head), "item", *patch.item < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.item)});
    if (patch.hierarchy) head = pivotPatchAttribute(std::move(head), "hier", *patch.hierarchy < 0 ? std::optional<std::string>{std::string{}} : std::optional<std::string>{std::to_string(*patch.hierarchy)});
    if (patch.name) head = pivotPatchAttribute(std::move(head), "name", patch.name->empty() ? std::optional<std::string>{std::string{}} : std::optional<std::string>{xmlEscape(*patch.name)});
    xml.replace(field.begin, headEnd - field.begin + 1, head);
    part->data = std::move(xml);

    auto& setting = pivotIt->pageFieldSettings()[pageFieldIndex];
    if (patch.fieldIndex) {
        setting.setFieldIndex(*patch.fieldIndex);
        if (pageFieldIndex < pivotIt->pageFields().size()) {
            pivotIt->pageFields()[pageFieldIndex].setFieldIndex(*patch.fieldIndex);
            pivotIt->pageFields()[pageFieldIndex].setName(pivotIt->cache().fields()[static_cast<std::size_t>(*patch.fieldIndex)]);
        }
    }
    if (patch.item) setting.setItem(*patch.item);
    if (patch.hierarchy) setting.setHierarchy(*patch.hierarchy);
    if (patch.name) setting.setName(*patch.name);
    return true;
}

bool Workbook::updateImportedPivotCacheRecordValue(const std::string& worksheetName,
                                                   const std::string& pivotTableName,
                                                   std::size_t recordIndex,
                                                   std::size_t fieldIndex,
                                                   const PivotCacheRecordValuePatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
        return sheet.name() == worksheetName;
    });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) {
        return pivot.name() == pivotTableName;
    });
    if (pivotIt == sheetIt->pivotTables_.end()) return false;
    const auto& cache = pivotIt->cache();
    if (recordIndex >= cache.records().size() || fieldIndex >= cache.fields().size()
        || fieldIndex >= cache.records()[recordIndex].size()) return false;
    const auto cachePart = cache.sharedCacheKey();
    if (cachePart.empty() || cachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) return false;

    std::string recordsPart;
    for (const auto& relationship : preservedRelationships_) {
        if (relationship.sourcePart != cachePart || relationship.targetMode == "External"
            || relationship.type.find("/pivotCacheRecords") == std::string::npos) continue;
        recordsPart = resolvePackagePart(cachePart, relationship.target);
        break;
    }
    if (recordsPart.empty()) {
        recordsPart = cachePart;
        const auto marker = recordsPart.find("pivotCacheDefinition");
        if (marker != std::string::npos) recordsPart.replace(marker, std::string("pivotCacheDefinition").size(), "pivotCacheRecords");
    }
    auto partIt = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& part) {
        return part.name == recordsPart;
    });
    if (partIt == preservedParts_.end()) return false;

    auto xml = partIt->data;
    const auto root = pivotBalancedElementRange(xml, "pivotCacheRecords");
    const auto record = pivotNthChildRange(xml, root, "r", recordIndex);
    const auto value = pivotNthCacheValueRange(xml, record, fieldIndex);
    if (!value) return false;

    std::string replacement;
    std::string modelValue = patch.value;
    switch (patch.type) {
        case PivotCacheRecordValueType::Missing:
            replacement = "<m/>";
            modelValue.clear();
            break;
        case PivotCacheRecordValueType::Number: {
            if (patch.value.empty()) throw std::invalid_argument("Pivot cache numeric value cannot be empty");
            std::size_t parsed = 0;
            try { (void)std::stod(patch.value, &parsed); }
            catch (...) { throw std::invalid_argument("Pivot cache numeric value is not a valid number"); }
            if (parsed != patch.value.size()) throw std::invalid_argument("Pivot cache numeric value is not a valid number");
            replacement = "<n v=\"" + xmlEscape(patch.value) + "\"/>";
            break;
        }
        case PivotCacheRecordValueType::String:
            replacement = "<s v=\"" + xmlEscape(patch.value) + "\"/>";
            break;
        case PivotCacheRecordValueType::Boolean:
            if (patch.value == "true" || patch.value == "1") {
                replacement = "<b v=\"1\"/>";
                modelValue = "true";
            } else if (patch.value == "false" || patch.value == "0") {
                replacement = "<b v=\"0\"/>";
                modelValue = "false";
            } else {
                throw std::invalid_argument("Pivot cache Boolean value must be true, false, 1, or 0");
            }
            break;
        case PivotCacheRecordValueType::Error:
            if (patch.value.empty()) throw std::invalid_argument("Pivot cache error value cannot be empty");
            replacement = "<e v=\"" + xmlEscape(patch.value) + "\"/>";
            break;
        case PivotCacheRecordValueType::DateTime:
            if (patch.value.empty()) throw std::invalid_argument("Pivot cache date/time value cannot be empty");
            replacement = "<d v=\"" + xmlEscape(patch.value) + "\"/>";
            break;
        case PivotCacheRecordValueType::SharedItem:
            if (patch.sharedItemIndex < 0) throw std::invalid_argument("Pivot cache shared-item index cannot be negative");
            replacement = "<x v=\"" + std::to_string(patch.sharedItemIndex) + "\"/>";
            break;
    }

    xml.replace(value.begin, value.end - value.begin, replacement);
    partIt->data = std::move(xml);

    for (auto& sheet : sheets_) {
        for (auto& pivot : sheet.pivotTables_) {
            auto& modelCache = pivot.cache();
            const auto& readOnlyCache = static_cast<const PivotCache&>(modelCache);
            if (modelCache.sharedCacheKey() != cachePart
                || recordIndex >= readOnlyCache.records().size()
                || fieldIndex >= readOnlyCache.records()[recordIndex].size()) continue;
            if (patch.type == PivotCacheRecordValueType::SharedItem) {
                // A shared-item index refers to the physical cache definition.
                // Keep the legacy string model update but invalidate exact type
                // metadata until the referenced shared item is semantically
                // resolved by a future cache-definition edit/load pass.
                modelCache.records()[recordIndex][fieldIndex] = modelValue;
                continue;
            }
            PivotCacheValueKind kind = PivotCacheValueKind::String;
            switch (patch.type) {
                case PivotCacheRecordValueType::Missing: kind = PivotCacheValueKind::Missing; break;
                case PivotCacheRecordValueType::Number: kind = PivotCacheValueKind::Number; break;
                case PivotCacheRecordValueType::String: kind = PivotCacheValueKind::String; break;
                case PivotCacheRecordValueType::Boolean: kind = PivotCacheValueKind::Boolean; break;
                case PivotCacheRecordValueType::Error: kind = PivotCacheValueKind::Error; break;
                case PivotCacheRecordValueType::DateTime: kind = PivotCacheValueKind::DateTime; break;
                case PivotCacheRecordValueType::SharedItem: break;
            }
            modelCache.setRecordValue(recordIndex, fieldIndex, modelValue, kind);
        }
    }
    return true;
}

PivotChartLinkValidationReport Workbook::validatePivotChartLinks() const {
    PivotChartLinkValidationReport report;
    struct PivotOwner { const Worksheet* sheet; const PivotTable* pivot; };
    std::vector<PivotOwner> pivots;
    for (const auto& sheet : sheets_) for (const auto& pivot : sheet.pivotTables()) pivots.push_back({&sheet, &pivot});

    const auto stripQualification = [](std::string value) {
        const auto bang = value.rfind('!');
        if (bang != std::string::npos) value = value.substr(bang + 1);
        if (value.size() >= 2 && value.front() == '\'' && value.back() == '\'') value = value.substr(1, value.size() - 2);
        return value;
    };
    const auto sourceSheetName = [](const std::string& value) -> std::string {
        const auto bang = value.rfind('!');
        if (bang == std::string::npos) return {};
        auto sheet = value.substr(0, bang);
        if (sheet.size() >= 2 && sheet.front() == '\'' && sheet.back() == '\'') {
            sheet = sheet.substr(1, sheet.size() - 2);
            std::size_t pos = 0;
            while ((pos = sheet.find("''", pos)) != std::string::npos) sheet.replace(pos, 2, "'");
        }
        return sheet;
    };

    for (const auto& sheet : sheets_) {
        for (const auto& chart : sheet.charts()) {
            if (!chart.pivotSource().present) continue;
            ++report.pivotChartsVisited;
            const auto pivotName = stripQualification(chart.pivotSource().pivotTableName);
            const auto qualifiedSheet = sourceSheetName(chart.pivotSource().pivotTableName);
            std::vector<PivotOwner> matches;
            for (const auto& owner : pivots) {
                if (owner.pivot->name() != pivotName) continue;
                if (!qualifiedSheet.empty() && owner.sheet->name() != qualifiedSheet) continue;
                matches.push_back(owner);
            }
            const auto chartId = chart.stableId().empty() ? chart.drawingObjectName() : chart.stableId();
            if (matches.empty()) {
                report.issues.push_back({sheet.name(), chartId, chart.pivotSource().pivotTableName,
                                         "PivotChart source does not resolve to a loaded PivotTable"});
                continue;
            }
            if (matches.size() > 1) {
                report.issues.push_back({sheet.name(), chartId, chart.pivotSource().pivotTableName,
                                         "PivotChart source is ambiguous; qualify the PivotTable with its worksheet"});
                continue;
            }
            const auto& pivot = *matches.front().pivot;
            const auto formatId = static_cast<std::uint32_t>(std::max(0, chart.pivotSource().formatId));
            if (pivot.chartFormatIndex() && *pivot.chartFormatIndex() != formatId) {
                report.issues.push_back({sheet.name(), chartId, pivot.name(),
                                         "PivotChart fmtId does not match PivotTable chartFormat"});
                continue;
            }
            if (!pivot.chartFormats().empty() && std::none_of(pivot.chartFormats().begin(), pivot.chartFormats().end(), [&](const auto& format) {
                    return format.chartIndex == formatId;
                })) {
                report.issues.push_back({sheet.name(), chartId, pivot.name(),
                                         "PivotChart fmtId is not represented by PivotTable chartFormats@chart"});
                continue;
            }
            ++report.validLinks;
        }
    }
    return report;
}

bool Workbook::updateImportedPivotOlapSource(const std::string& worksheetName,
                                             const std::string& pivotTableName,
                                             const PivotOlapSourcePatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
        return sheet.name() == worksheetName;
    });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) {
        return pivot.name() == pivotTableName;
    });
    if (pivotIt == sheetIt->pivotTables_.end()) return false;
    const auto cachePart = pivotIt->cache().sharedCacheKey();
    if (cachePart.empty() || cachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) return false;
    const auto* olap = static_cast<const xlpp::PivotCache&>(pivotIt->cache()).olap();
    if (!olap) return false; // OLAP patch is only valid for OLAP-backed caches.

    auto partIt = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& part) {
        return part.name == cachePart;
    });
    if (partIt == preservedParts_.end()) return false;
    auto xml = partIt->data;

    auto patchOlapAttribute = [&](const std::string& name, const std::optional<std::string>& value) {
        if (!value) return;
        const auto olapBegin = xml.find("<olapInfo");
        if (olapBegin == std::string::npos) return;
        const auto olapEnd = xml.find('>', olapBegin);
        if (olapEnd == std::string::npos) throw std::runtime_error("Malformed OLAP info start tag");
        std::string olapTag = xml.substr(olapBegin, olapEnd - olapBegin + 1);
        const auto key = name + "=\"";
        const auto begin = olapTag.find(key);
        if (value->empty()) {
            if (begin != std::string::npos) {
                auto eraseBegin = begin;
                while (eraseBegin > 0 && std::isspace(static_cast<unsigned char>(olapTag[eraseBegin - 1]))) --eraseBegin;
                const auto valueBegin = begin + key.size();
                const auto valueEnd = olapTag.find('"', valueBegin);
                if (valueEnd == std::string::npos) throw std::runtime_error("Malformed OLAP info attribute");
                olapTag.erase(eraseBegin, valueEnd - eraseBegin + 1);
            }
        } else if (begin != std::string::npos) {
            const auto valueBegin = begin + key.size();
            const auto valueEnd = olapTag.find('"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed OLAP info attribute");
            olapTag.replace(valueBegin, valueEnd - valueBegin, *value);
        } else {
            auto insertPos = olapTag.size() - 1;
            if (insertPos > 0 && olapTag[insertPos - 1] == '/') --insertPos;
            olapTag.insert(insertPos, " " + name + "=\"" + *value + "\"");
        }
        xml.replace(olapBegin, olapEnd - olapBegin + 1, olapTag);
    };

    if (patch.preserveFormatting) patchOlapAttribute("preserveFormatting",
        std::optional<std::string>(*patch.preserveFormatting ? "1" : "0"));
    if (patch.localCube) patchOlapAttribute("localCube", *patch.localCube);
    if (patch.localConnection) patchOlapAttribute("localConnection", *patch.localConnection);
    // connectionId is a cacheSource attribute in SpreadsheetML, not a
    // pivotCacheDefinition root attribute.
    if (patch.connectionId) {
        const auto sourceBegin = xml.find("<cacheSource");
        if (sourceBegin == std::string::npos) return false;
        const auto sourceEnd = xml.find('>', sourceBegin);
        if (sourceEnd == std::string::npos) throw std::runtime_error("Malformed pivot cacheSource start tag");
        std::string sourceTag = xml.substr(sourceBegin, sourceEnd - sourceBegin + 1);
        const auto key = std::string("connectionId=\"");
        const auto begin = sourceTag.find(key);
        if (begin != std::string::npos) {
            const auto valueBegin = begin + key.size();
            const auto valueEnd = sourceTag.find('"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed cacheSource connectionId");
            sourceTag.replace(valueBegin, valueEnd - valueBegin, std::to_string(*patch.connectionId));
        } else {
            auto insertPos = sourceTag.size() - 1;
            if (insertPos > 0 && sourceTag[insertPos - 1] == '/') --insertPos;
            sourceTag.insert(insertPos, " connectionId=\"" + std::to_string(*patch.connectionId) + "\"");
        }
        xml.replace(sourceBegin, sourceEnd - sourceBegin + 1, sourceTag);
    }

    // Refresh raw olapInfo bytes for the in-memory model BEFORE moving xml.
    const auto olapInfoBegin = xml.find("<olapInfo");
    std::string refreshedOlapInfo;
    if (olapInfoBegin != std::string::npos) {
        const auto openEnd = xml.find('>', olapInfoBegin);
        if (openEnd != std::string::npos) {
            if (openEnd > olapInfoBegin && xml[openEnd - 1] == '/') {
                refreshedOlapInfo = xml.substr(olapInfoBegin, openEnd - olapInfoBegin + 1);
            } else {
                const auto close = xml.find("</olapInfo>", openEnd);
                if (close != std::string::npos)
                    refreshedOlapInfo = xml.substr(olapInfoBegin, close + std::string("</olapInfo>").size() - olapInfoBegin);
            }
        }
    }
    partIt->data = std::move(xml);
    for (auto& sheet : sheets_) {
        for (auto& pivot : sheet.pivotTables_) {
            if (pivot.cache().sharedCacheKey() != cachePart) continue;
            auto& model = pivot.cache().olap();
            if (patch.preserveFormatting) model.preserveFormatting = *patch.preserveFormatting;
            if (patch.localCube) model.localCube = *patch.localCube;
            if (patch.localConnection) model.localConnection = *patch.localConnection;
            if (patch.connectionId) model.connectionId = *patch.connectionId;
            if (!refreshedOlapInfo.empty()) model.rawOlapInfoXml = refreshedOlapInfo;
        }
    }
    return true;
}

bool Workbook::updateImportedPivotCalculatedMember(const std::string& worksheetName,
                                                   const std::string& pivotTableName,
                                                   std::size_t memberIndex,
                                                   const PivotCalculatedMemberPatch& patch) {
    auto sheetIt = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& sheet) {
        return sheet.name() == worksheetName;
    });
    if (sheetIt == sheets_.end()) return false;
    auto pivotIt = std::find_if(sheetIt->pivotTables_.begin(), sheetIt->pivotTables_.end(), [&](const auto& pivot) {
        return pivot.name() == pivotTableName;
    });
    if (pivotIt == sheetIt->pivotTables_.end()) return false;
    const auto cachePart = pivotIt->cache().sharedCacheKey();
    if (cachePart.empty() || cachePart.rfind("xl/pivotCache/pivotCacheDefinition", 0) != 0) return false;
    if (memberIndex >= pivotIt->cache().calculatedMembers().size()) return false;

    auto partIt = std::find_if(preservedParts_.begin(), preservedParts_.end(), [&](const auto& part) {
        return part.name == cachePart;
    });
    if (partIt == preservedParts_.end()) return false;
    auto xml = partIt->data;

    std::size_t search = 0;
    std::size_t begin = std::string::npos;
    for (std::size_t i = 0; i <= memberIndex; ++i) {
        while (true) {
            begin = xml.find("<calculatedMember", search);
            if (begin == std::string::npos) return false;
            const auto after = begin + std::string("<calculatedMember").size();
            if (after >= xml.size() || std::isspace(static_cast<unsigned char>(xml[after])) || xml[after] == '>' || xml[after] == '/') break;
            search = after;
        }
        if (i != memberIndex) search = begin + 1;
    }
    const auto end = xml.find('>', begin);
    if (end == std::string::npos) return false;
    auto head = xml.substr(begin, end - begin + 1);
    const auto setAttribute = [&](const std::string& name, const std::optional<std::string>& value) {
        const auto key = name + "=\"";
        const auto attrBegin = head.find(key);
        if (!value || value->empty()) {
            if (attrBegin != std::string::npos) {
                auto eraseBegin = attrBegin;
                while (eraseBegin > 0 && std::isspace(static_cast<unsigned char>(head[eraseBegin - 1]))) --eraseBegin;
                const auto valueBegin = attrBegin + key.size();
                const auto valueEnd = head.find('"', valueBegin);
                if (valueEnd == std::string::npos) throw std::runtime_error("Malformed calculatedMember attribute");
                head.erase(eraseBegin, valueEnd - eraseBegin + 1);
            }
            return;
        }
        if (attrBegin != std::string::npos) {
            const auto valueBegin = attrBegin + key.size();
            const auto valueEnd = head.find('"', valueBegin);
            if (valueEnd == std::string::npos) throw std::runtime_error("Malformed calculatedMember attribute");
            head.replace(valueBegin, valueEnd - valueBegin, *value);
            return;
        }
        auto insertPos = head.size() - 1;
        if (insertPos > 0 && head[insertPos - 1] == '/') --insertPos;
        head.insert(insertPos, " " + name + "=\"" + *value + "\"");
    };

    if (patch.mdx) setAttribute("mdx", *patch.mdx);
    if (patch.memberName) setAttribute("memberName", *patch.memberName);
    if (patch.hierarchy) setAttribute("hierarchy", std::optional<std::string>(std::to_string(*patch.hierarchy)));
    if (patch.solveOrder) setAttribute("solveOrder", *patch.solveOrder);
    if (patch.set) setAttribute("set", *patch.set);
    xml.replace(begin, end - begin + 1, head);
    // Refresh the patched member's raw XML so a later save re-emits the edited
    // tag instead of the stale pre-patch subtree.
    const auto patchedBegin = xml.find("<calculatedMember", begin);
    const auto patchedEnd = xml.find('>', patchedBegin == std::string::npos ? 0 : patchedBegin);
    std::string refreshedRawXml;
    if (patchedBegin != std::string::npos && patchedEnd != std::string::npos) {
        refreshedRawXml = xml.substr(patchedBegin, patchedEnd - patchedBegin + 1);
        if (!refreshedRawXml.empty() && refreshedRawXml.back() == '>' && refreshedRawXml[refreshedRawXml.size() - 2] != '/') {
            const auto close = xml.find("</calculatedMember>", patchedEnd);
            if (close != std::string::npos)
                refreshedRawXml = xml.substr(patchedBegin, close + std::string("</calculatedMember>").size() - patchedBegin);
        }
    }
    partIt->data = std::move(xml);

    // Mirror scalar edits into the loaded model so repeated edits stay coherent.
    for (auto& sheet : sheets_) {
        for (auto& pivot : sheet.pivotTables_) {
            if (pivot.cache().sharedCacheKey() != cachePart) continue;
            auto& members = pivot.cache().calculatedMembers();
            if (memberIndex >= members.size()) continue;
            auto& member = members[memberIndex];
            if (patch.mdx) member.mdx = *patch.mdx;
            if (patch.memberName) member.memberName = *patch.memberName;
            if (patch.hierarchy) member.hierarchy = *patch.hierarchy;
            if (patch.solveOrder) member.solveOrder = *patch.solveOrder;
            if (patch.set) member.set = *patch.set;
            if (!refreshedRawXml.empty()) member.rawXml = refreshedRawXml;
        }
    }
    return true;
}


} // namespace xlpp
