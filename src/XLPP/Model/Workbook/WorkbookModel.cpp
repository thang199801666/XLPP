#include <XLPP/Workbook/Workbook.h>
#include <XLPP/Worksheet/WorksheetName.h>
#include <XLPP/Formula/ReferenceTranslator.h>
#include <algorithm>
#include <cctype>
#include <stdexcept>
#include <utility>

namespace {

bool asciiCaseEqual(std::string_view a, std::string_view b) noexcept {
    if (a.size() != b.size()) return false;
    for (std::size_t i = 0; i < a.size(); ++i) {
        const auto ca = static_cast<unsigned char>(a[i]);
        const auto cb = static_cast<unsigned char>(b[i]);
        if (ca < 0x80u && cb < 0x80u) {
            if (std::tolower(ca) != std::tolower(cb)) return false;
        } else if (ca != cb) {
            return false;
        }
    }
    return true;
}

std::string uniqueWorkbookObjectName(const std::string& base,
                                     const std::vector<std::string>& occupied) {
    const auto isOccupied = [&](std::string_view candidate) {
        return std::any_of(occupied.begin(), occupied.end(), [&](const auto& name) {
            return asciiCaseEqual(name, candidate);
        });
    };
    if (!isOccupied(base)) return base;
    for (std::size_t suffix = 2;; ++suffix) {
        const auto candidate = base + "_" + std::to_string(suffix);
        if (!isOccupied(candidate)) return candidate;
    }
}

} // namespace

namespace xlpp {

NamedStyle& Workbook::addNamedStyle(NamedStyle style) {
    if (style.name().empty()) throw std::invalid_argument("Named style name cannot be empty");
    if (namedStyle(style.name())) throw std::invalid_argument("Named style already exists: " + style.name());
    namedStyles_.push_back(std::move(style));
    return namedStyles_.back();
}

NamedStyle* Workbook::namedStyle(const std::string& name) noexcept {
    const auto it = std::find_if(namedStyles_.begin(), namedStyles_.end(),
                                 [&](const auto& style) { return asciiCaseEqual(style.name(), name); });
    return it == namedStyles_.end() ? nullptr : &*it;
}

const NamedStyle* Workbook::namedStyle(const std::string& name) const noexcept {
    const auto it = std::find_if(namedStyles_.begin(), namedStyles_.end(),
                                 [&](const auto& style) { return asciiCaseEqual(style.name(), name); });
    return it == namedStyles_.end() ? nullptr : &*it;
}

void Workbook::applyNamedStyle(Cell& cell, const std::string& name) const {
    const auto* style = namedStyle(name);
    if (!style) throw std::out_of_range("Unknown named style: " + name);
    cell.style() = style->style();
    cell.setNamedStyle(style->name());
}

DefinedName& Workbook::addDefinedName(DefinedName name) {
    if (name.localSheetId() && *name.localSheetId() >= sheets_.size())
        throw std::out_of_range("Defined-name localSheetId is outside the workbook");
    if (definedName(name.name(), name.localSheetId()))
        throw std::invalid_argument("Defined name already exists in the same scope: " + name.name());
    definedNames_.push_back(std::move(name));
    return definedNames_.back();
}

DefinedName* Workbook::definedName(const std::string& name,
                                   std::optional<std::size_t> localSheetId) noexcept {
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(), [&](auto& item) {
        return item.localSheetId() == localSheetId && asciiCaseEqual(item.name(), name);
    });
    return it == definedNames_.end() ? nullptr : &*it;
}

const DefinedName* Workbook::definedName(const std::string& name,
                                         std::optional<std::size_t> localSheetId) const noexcept {
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(), [&](const auto& item) {
        return item.localSheetId() == localSheetId && asciiCaseEqual(item.name(), name);
    });
    return it == definedNames_.end() ? nullptr : &*it;
}

DefinedName* Workbook::definedName(const std::string& name) noexcept {
    if (auto* global = definedName(name, std::nullopt)) return global;
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(),
                                 [&](auto& item) { return asciiCaseEqual(item.name(), name); });
    return it == definedNames_.end() ? nullptr : &*it;
}

const DefinedName* Workbook::definedName(const std::string& name) const noexcept {
    if (const auto* global = definedName(name, std::nullopt)) return global;
    const auto it = std::find_if(definedNames_.begin(), definedNames_.end(),
                                 [&](const auto& item) { return asciiCaseEqual(item.name(), name); });
    return it == definedNames_.end() ? nullptr : &*it;
}

Worksheet& Workbook::addWorksheet(std::string name) {
    validateWorksheetName(name);
    if (worksheet(name)) throw std::invalid_argument("Duplicate worksheet name (case-insensitive): " + name);
    sheets_.emplace_back(std::move(name));
    auto& sheet = sheets_.back();
    // VBA code names are independent from visible worksheet names. Allocate a
    // stable, workbook-unique code name at insertion time so deleting/reordering
    // sheets does not silently retarget worksheet event procedures.
    for (std::size_t candidate = 1;; ++candidate) {
        const auto codeName = "Sheet" + std::to_string(candidate);
        bool used = false;
        for (const auto& existing : sheets_) {
            if (&existing == &sheet) continue;
            if (asciiCaseEqual(existing.vbaCodeName(), codeName)) { used = true; break; }
        }
        if (!used) { sheet.setVbaCodeName(codeName); break; }
    }
    return sheet;
}

bool Workbook::removeWorksheet(const std::string& name) {
    const auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                 [&](const auto& sheet) { return worksheetNamesEqual(sheet.name(), name); });
    if (it == sheets_.end()) return false;
    const auto removedIndex = static_cast<std::size_t>(std::distance(sheets_.begin(), it));
    const auto removedName = it->name();

    // Excel turns explicit dependencies on a deleted sheet into #REF! rather
    // than leaving a syntactically valid reference to a sheet that no longer
    // exists. Rewrite all surviving worksheet-owned dependency surfaces before
    // erasing the target object.
    for (auto sheetIt = sheets_.begin(); sheetIt != sheets_.end(); ++sheetIt) {
        if (sheetIt != it) (void)sheetIt->invalidateWorksheetReferencesForRemoval(removedName);
    }
    for (auto& definedName : definedNames_) {
        if (definedName.localSheetId() && *definedName.localSheetId() == removedIndex) continue;
        auto translated = invalidateWorksheetReferences(definedName.value(), removedName);
        if (translated.changed()) definedName.setValue(std::move(translated.value));
    }

    sheets_.erase(it);

    // Excel localSheetId is positional. Remove names owned by the deleted
    // sheet and compact scopes after it so names never silently retarget.
    for (auto nameIt = definedNames_.begin(); nameIt != definedNames_.end();) {
        if (nameIt->localSheetId() && *nameIt->localSheetId() == removedIndex)
            nameIt = definedNames_.erase(nameIt);
        else
            ++nameIt;
    }
    for (auto& item : definedNames_) {
        if (item.localSheetId() && *item.localSheetId() > removedIndex)
            item.setLocalSheetId(*item.localSheetId() - 1u);
    }
    resetChartCacheDependencyTracking();
    calcProperties().setFullCalcOnLoad(true);
    calcProperties().setCalcOnSave(true);
    return true;
}

Worksheet* Workbook::worksheet(const std::string& name) noexcept {
    const auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                 [&](auto& sheet) { return worksheetNamesEqual(sheet.name(), name); });
    return it == sheets_.end() ? nullptr : &*it;
}

const Worksheet* Workbook::worksheet(const std::string& name) const noexcept {
    const auto it = std::find_if(sheets_.begin(), sheets_.end(),
                                 [&](const auto& sheet) { return worksheetNamesEqual(sheet.name(), name); });
    return it == sheets_.end() ? nullptr : &*it;
}

Worksheet& Workbook::operator[](std::size_t index) { return sheets_.at(index); }
const Worksheet& Workbook::operator[](std::size_t index) const { return sheets_.at(index); }

std::size_t Workbook::index(const Worksheet& sheet) const {
    const auto it = std::find_if(sheets_.begin(), sheets_.end(), [&](const auto& candidate) {
        return &candidate == &sheet;
    });
    if (it == sheets_.end()) throw std::out_of_range("Worksheet not in this workbook");
    return static_cast<std::size_t>(std::distance(sheets_.begin(), it));
}

std::vector<std::string> Workbook::sheetNames() const {
    std::vector<std::string> names;
    names.reserve(sheets_.size());
    for (const auto& sheet : sheets_) names.push_back(sheet.name());
    return names;
}

Worksheet& Workbook::copyWorksheet(const Worksheet& source, std::string newName) {
    validateWorksheetName(newName);
    if (worksheet(newName)) throw std::invalid_argument("Duplicate worksheet name (case-insensitive): " + newName);

    // Capture ownership before mutating the deque. A worksheet can also be
    // copied into a different workbook, in which case there are no source-local
    // defined names to clone.
    std::optional<std::size_t> sourceIndex;
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
        if (&sheets_[i] == &source) {
            sourceIndex = i;
            break;
        }
    }

    const auto sourceName = source.name();
    const auto destinationName = newName;
    std::optional<VbaModule> sourceDocumentModule;
    if (sourceIndex && generatedVbaProject_) {
        for (const auto& module : vbaModules()) {
            if (module.type == VbaModuleType::Document && asciiCaseEqual(module.name, source.vbaCodeName())) {
                sourceDocumentModule = module;
                break;
            }
        }
    }
    Worksheet copy = source;
    for (std::size_t candidate = 1;; ++candidate) {
        const auto codeName = "Sheet" + std::to_string(candidate);
        bool used = false;
        for (const auto& existing : sheets_) {
            if (asciiCaseEqual(existing.vbaCodeName(), codeName)) { used = true; break; }
        }
        if (!used) { copy.setVbaCodeName(codeName); break; }
    }

    // Excel's Copy Sheet semantics are a deep topology copy: explicit
    // references from objects owned by the source sheet to that same source
    // sheet must follow the clone. Cross-sheet references remain untouched.
    (void)copy.translateWorksheetRenameReferences(sourceName, destinationName);

    // Table and PivotTable names are workbook-global identifiers. A naive
    // value copy creates an invalid package with duplicate names, so allocate
    // deterministic unique identifiers while retaining object formatting and
    // behavior.
    std::vector<std::string> tableNames;
    std::vector<std::string> pivotNames;
    for (const auto& sheet : sheets_) {
        for (const auto& table : static_cast<const Worksheet&>(sheet).tables())
            tableNames.push_back(table.name());
        for (const auto& pivot : static_cast<const Worksheet&>(sheet).pivotTables())
            pivotNames.push_back(pivot.name());
    }

    for (auto& table : copy.tables()) {
        const auto oldName = table.name();
        const auto uniqueName = uniqueWorkbookObjectName(oldName, tableNames);
        if (!asciiCaseEqual(uniqueName, oldName)) {
            Table replacement(uniqueName, table.reference());
            replacement.setDisplayName(uniqueName);
            replacement.setShowHeaderRow(table.showHeaderRow());
            replacement.setShowTotalsRow(table.showTotalsRow());
            replacement.columns() = table.columns();
            replacement.styleInfo() = table.styleInfo();
            table = std::move(replacement);
        }
        tableNames.push_back(table.name());
    }
    for (auto& pivot : copy.pivotTables()) {
        const auto uniqueName = uniqueWorkbookObjectName(pivot.name(), pivotNames);
        pivot.setName(uniqueName);
        pivotNames.push_back(uniqueName);
    }

    copy.rename(destinationName);
    const auto destinationIndex = sheets_.size();
    sheets_.push_back(std::move(copy));

    // Sheet-local defined names are positional workbook metadata rather than a
    // Worksheet member. Clone the source scope and retarget explicit self
    // qualifiers to the copied worksheet. Global names intentionally remain
    // shared and are not duplicated.
    if (sourceIndex) {
        std::vector<DefinedName> localNames;
        for (const auto& definedName : definedNames_) {
            if (!definedName.localSheetId() || *definedName.localSheetId() != *sourceIndex) continue;
            auto cloned = definedName;
            cloned.setLocalSheetId(destinationIndex);
            auto translated = renameWorksheetReferences(cloned.value(), sourceName, destinationName);
            if (translated.changed()) cloned.setValue(std::move(translated.value));
            localNames.push_back(std::move(cloned));
        }
        for (auto& definedName : localNames) definedNames_.push_back(std::move(definedName));
    }

    if (sourceDocumentModule && !sourceDocumentModule->source.empty()) {
        sourceDocumentModule->name = sheets_.back().vbaCodeName();
        setVbaModule(std::move(*sourceDocumentModule));
    }

    resetChartCacheDependencyTracking();
    calcProperties().setFullCalcOnLoad(true);
    calcProperties().setCalcOnSave(true);
    return sheets_.back();
}

} // namespace xlpp
