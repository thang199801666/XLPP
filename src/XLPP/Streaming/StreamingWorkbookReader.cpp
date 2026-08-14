#include <XLPP/Streaming/StreamingWorkbookReader.h>
#include "../Packaging/ZipArchiveReader.h"
#include "../Streaming/SharedStringsReader.h"
#include "../XML/XmlPullReader.h"
#include "../XML/XmlScanner.h"
#include "../XML/XmlUtilities.h"
#include <optional>
#include <set>
#include <stdexcept>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

namespace xlpp::internal {

class WorksheetRowSource {
public:
    WorksheetRowSource(ZipArchiveReader archive, std::string entryName,
                       std::shared_ptr<const SharedStringsReader> sharedStrings,
                       std::size_t maxXmlElementBytes)
        : entry_(archive.openEntry(std::move(entryName))),
          sharedStrings_(std::move(sharedStrings)),
          xml_(std::in_place, [this](unsigned char* out, std::size_t capacity) {
              return entry_.read(out, capacity);
          }, maxXmlElementBytes) {}

    bool next(StreamingRow& row, std::size_t& rowNumber) {
        const auto rowTag = xml_->nextElement("row");
        if (rowTag.empty()) return false;
        const auto indexText = xmlAttribute(rowTag, "r");
        rowNumber = 0;
        if (!indexText.empty() && !parseSize(indexText, rowNumber))
            throw std::runtime_error("Invalid worksheet row index");
        row.clear();
        XmlScanner cells(rowTag);
        std::string_view cellTag;
        while (cells.nextElement("c", cellTag)) {
            StreamingCell cell;
            const auto address = xmlAttribute(cellTag, "r");
            cell.address.assign(address.data(), address.size());
            const auto styleValue = xmlAttribute(cellTag, "s");
            if (!styleValue.empty()) {
                std::size_t sid = 0;
                if (!parseSize(styleValue, sid))
                    throw std::runtime_error("Invalid worksheet cell style index");
                cell.styleIndex = sid;
            }
            const auto type = xmlAttribute(cellTag, "t");
            if (type == "inlineStr") {
                const auto text = xmlText(cellTag, "t");
                cell.value = containsEntity(text) ? xmlUnescape(text) : std::string(text);
            } else if (type == "b") {
                cell.value = xmlText(cellTag, "v") == "1";
            } else if (type == "e") {
                cell.value = cellErrorFromString(xmlText(cellTag, "v"));
            } else if (type == "s") {
                const auto value = xmlText(cellTag, "v");
                if (!value.empty() && sharedStrings_) {
                    std::size_t index = 0;
                    if (!parseSize(value, index))
                        throw std::runtime_error("Invalid shared-string index");
                    if (const auto* text = sharedStrings_->lookup(index)) cell.value = *text;
                    else throw std::runtime_error("Shared-string index is out of range");
                }
            } else {
                const auto value = xmlText(cellTag, "v");
                if (!value.empty()) {
                    double number = 0.0;
                    if (!parseDouble(value, number))
                        throw std::runtime_error("Invalid numeric cell value");
                    cell.value = number;
                }
            }
            const auto formula = xmlText(cellTag, "f");
            if (!formula.empty())
                cell.formula = containsEntity(formula) ? xmlUnescape(formula)
                                                       : std::string(formula);
            row.push_back(std::move(cell));
        }
        return true;
    }

private:
    ZipEntrySource entry_;
    std::shared_ptr<const SharedStringsReader> sharedStrings_;
    std::optional<XmlPullReader> xml_;
};

} // namespace xlpp::internal

namespace {

struct SheetBinding {
    std::string name;
    std::string entry;
};

bool hasUriScheme(std::string_view target) {
    const auto colon = target.find(':');
    if (colon == std::string_view::npos) return false;
    const auto slash = target.find('/');
    if (slash != std::string_view::npos && slash < colon) return false;
    if (colon == 0) return false;
    const auto alpha = [](char c) { return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'); };
    if (!alpha(target.front())) return false;
    for (std::size_t i = 1; i < colon; ++i) {
        const char c = target[i];
        if (!(alpha(c) || (c >= '0' && c <= '9') || c == '+' || c == '-' || c == '.')) return false;
    }
    return true;
}

std::string resolveInternalPackageTarget(std::string_view sourcePart, std::string_view target) {
    if (target.empty()) throw std::runtime_error("OOXML relationship target is empty");
    if (target.find('\0') != std::string_view::npos || target.find('\\') != std::string_view::npos)
        throw std::runtime_error("OOXML relationship target contains an invalid path character");
    if (hasUriScheme(target))
        throw std::runtime_error("Internal OOXML relationship target must not contain a URI scheme");

    // OPC targets may contain a fragment. It does not participate in package
    // part lookup. Query components are likewise not package-name bytes.
    const auto suffix = target.find_first_of("?#");
    if (suffix != std::string_view::npos) target = target.substr(0, suffix);
    if (target.empty()) throw std::runtime_error("OOXML relationship target does not name a package part");

    std::string combined;
    if (target.front() == '/') {
        combined.assign(target.substr(1));
    } else {
        const auto slash = sourcePart.rfind('/');
        if (slash != std::string_view::npos) {
            combined.assign(sourcePart.substr(0, slash + 1));
        }
        combined.append(target);
    }

    std::vector<std::string_view> segments;
    std::size_t begin = 0;
    while (begin <= combined.size()) {
        const auto end = combined.find('/', begin);
        const auto segment = std::string_view(combined).substr(
            begin, (end == std::string::npos ? combined.size() : end) - begin);
        if (segment.empty() || segment == ".") {
            // Ignore duplicate separators and current-directory segments.
        } else if (segment == "..") {
            if (segments.empty())
                throw std::runtime_error("OOXML relationship target escapes the package root");
            segments.pop_back();
        } else {
            segments.push_back(segment);
        }
        if (end == std::string::npos) break;
        begin = end + 1;
    }
    if (segments.empty()) throw std::runtime_error("OOXML relationship target resolves to the package root");

    std::string result;
    for (const auto segment : segments) {
        if (!result.empty()) result.push_back('/');
        result.append(segment.data(), segment.size());
    }
    return result;
}

std::vector<SheetBinding> readSheetBindings(const xlpp::internal::ZipArchiveReader& archive,
                                           std::vector<std::string>* chartsheetNames,
                                           std::vector<std::string>* workbookSheetNames) {
    const auto workbook = archive.readEntry("xl/workbook.xml");
    std::unordered_map<std::string, std::string> targets;
    std::set<std::string> chartsheetRelationshipIds;
    const bool hasRelationships = archive.contains("xl/_rels/workbook.xml.rels");
    if (hasRelationships) {
        const auto rels = archive.readEntry("xl/_rels/workbook.xml.rels");
        xlpp::internal::XmlScanner relsScanner(rels);
        std::string_view relationship;
        while (relsScanner.nextElement("Relationship", relationship)) {
            const auto id = xlpp::internal::xmlAttribute(relationship, "Id");
            const auto target = xlpp::internal::xmlAttribute(relationship, "Target");
            const auto targetMode = xlpp::internal::xmlAttribute(relationship, "TargetMode");
            const auto type = xlpp::internal::xmlAttribute(relationship, "Type");
            if (id.empty()) throw std::runtime_error("Workbook relationship is missing Id");
            if (targetMode == "External") continue;
            if (!type.empty() && type.ends_with("/chartsheet")) {
                if (!chartsheetRelationshipIds.insert(std::string(id)).second)
                    throw std::runtime_error("Duplicate workbook relationship Id: " + std::string(id));
                continue;
            }
            // Only worksheet relationships are row-streaming candidates.
            if (!type.empty() && !type.ends_with("/worksheet")) continue;
            const auto resolved = resolveInternalPackageTarget("xl/workbook.xml", target);
            auto [it, inserted] = targets.emplace(std::string(id), resolved);
            if (!inserted)
                throw std::runtime_error("Duplicate workbook relationship Id: " + std::string(id));
        }
    }

    std::vector<SheetBinding> bindings;
    std::set<std::string> worksheetNames;
    xlpp::internal::XmlScanner wbScanner(workbook);
    std::string_view sheet;
    while (wbScanner.nextElement("sheet", sheet)) {
        const auto nameView = xlpp::internal::xmlAttribute(sheet, "name");
        if (nameView.empty()) throw std::runtime_error("Worksheet is missing a name");
        std::string name = xlpp::internal::xmlUnescape(nameView);
        if (!worksheetNames.insert(name).second)
            throw std::runtime_error("Duplicate worksheet name: " + name);

        std::string rid = std::string(xlpp::internal::xmlAttribute(sheet, "r:id"));
        if (rid.empty()) rid = std::string(xlpp::internal::xmlAttribute(sheet, "id"));
        if (workbookSheetNames) workbookSheetNames->push_back(name);
        if (!rid.empty() && chartsheetRelationshipIds.contains(rid)) {
            if (chartsheetNames) chartsheetNames->push_back(std::move(name));
            continue;
        }
        std::string target;
        if (!rid.empty()) {
            const auto it = targets.find(rid);
            if (it != targets.end()) target = it->second;
            else if (hasRelationships)
                throw std::runtime_error("Worksheet relationship not found: " + rid);
        }
        // Retain compatibility with minimal packages that omit workbook rels
        // entirely; once a rels part exists, dangling r:id is considered malformed.
        if (target.empty()) target = "xl/worksheets/sheet" + std::to_string(bindings.size() + 1) + ".xml";
        if (!archive.contains(target))
            throw std::runtime_error("Worksheet part not found: " + target);
        bindings.push_back(SheetBinding{std::move(name), std::move(target)});
    }
    return bindings;
}

} // namespace

namespace xlpp {

struct StreamingWorkbookReader::SharedState {
    SharedState(const std::filesystem::path& path, const StreamingReadOptions& options)
        : archive(path, internal::ZipOpenLimits{options.maxEntries, options.maxEntryBytes,
                                               options.maxTotalBytes, options.maxFileBytes}),
          maxXmlElementBytes(options.maxXmlElementBytes) {}
    internal::ZipArchiveReader archive;
    std::vector<SheetBinding> bindings;
    std::vector<std::string> chartsheetNames;
    std::vector<std::string> workbookSheetNames;
    std::shared_ptr<internal::SharedStringsReader> sharedStrings;
    std::size_t maxXmlElementBytes{internal::XmlPullReader::kDefaultMaxBufferedBytes};
};

StreamingWorkbookReader::StreamingWorkbookReader(const std::filesystem::path& path)
    : StreamingWorkbookReader(path, StreamingReadOptions{}) {}

StreamingWorkbookReader::StreamingWorkbookReader(const std::filesystem::path& path,
                                                   const StreamingReadOptions& options) {
    auto state = std::make_shared<SharedState>(path, options);
    state->bindings = readSheetBindings(state->archive, &state->chartsheetNames, &state->workbookSheetNames);
    state_ = std::move(state);
}

std::vector<std::string> StreamingWorkbookReader::worksheetNames() const {
    std::vector<std::string> names;
    names.reserve(state_->bindings.size());
    for (const auto& binding : state_->bindings) names.push_back(binding.name);
    return names;
}

std::vector<std::string> StreamingWorkbookReader::chartsheetNames() const {
    return state_->chartsheetNames;
}

std::vector<std::string> StreamingWorkbookReader::workbookSheetNames() const {
    return state_->workbookSheetNames;
}

StreamingWorksheetReader StreamingWorkbookReader::worksheet(const std::string& worksheetName) const {
    for (const auto& binding : state_->bindings) {
        if (binding.name == worksheetName) {
            if (!state_->sharedStrings)
                state_->sharedStrings = std::make_shared<internal::SharedStringsReader>(state_->archive);
            auto state = state_;
            const auto entry = binding.entry;
            auto factory = [state, entry]() -> std::shared_ptr<internal::WorksheetRowSource> {
                return std::make_shared<internal::WorksheetRowSource>(state->archive, entry,
                                                                      state->sharedStrings,
                                                                      state->maxXmlElementBytes);
            };
            return StreamingWorksheetReader(std::move(factory));
        }
    }
    throw std::runtime_error("Worksheet not found: " + worksheetName);
}

void StreamingWorkbookReader::forEachRow(const std::string& worksheetName,
                                         const std::function<bool(std::size_t, const StreamingRow&)>& callback) const {
    worksheet(worksheetName).forEachRow(callback);
}

StreamingWorksheetReader::StreamingWorksheetReader(
    std::function<std::shared_ptr<internal::WorksheetRowSource>()> factory)
    : factory_(std::move(factory)) {}

StreamingRowIterator StreamingWorksheetReader::begin() {
    return StreamingRowIterator(factory_());
}

void StreamingWorksheetReader::forEachRow(const std::function<bool(std::size_t, const StreamingRow&)>& callback) {
    for (auto it = begin(); it != end(); ++it) {
        if (!callback(it.rowNumber(), *it)) break;
    }
}

StreamingRowIterator::StreamingRowIterator(std::shared_ptr<internal::WorksheetRowSource> source)
    : source_(std::move(source)) {
    advance();
}

void StreamingRowIterator::advance() {
    valid_ = source_ && source_->next(current_, rowNumber_);
}

StreamingRowIterator& StreamingRowIterator::operator++() {
    advance();
    return *this;
}

bool StreamingRowIterator::operator==(const StreamingRowIterator& other) const noexcept {
    if (valid_ != other.valid_) return false;
    return !valid_ || source_ == other.source_;
}

} // namespace xlpp
