#include <XLPP/Streaming/StreamingWorkbookReader.h>
#include "../Packaging/ZipArchiveReader.h"
#include "../Streaming/SharedStringsReader.h"
#include "../XML/XmlPullReader.h"
#include "../XML/XmlScanner.h"
#include "../XML/XmlUtilities.h"
#include <optional>
#include <stdexcept>
#include <string>
#include <unordered_map>
#include <utility>

namespace xlpp::internal {

class WorksheetRowSource {
public:
    WorksheetRowSource(ZipArchiveReader archive, std::string entryName,
                       std::shared_ptr<const SharedStringsReader> sharedStrings)
        : entry_(archive.openEntry(std::move(entryName))),
          sharedStrings_(std::move(sharedStrings)),
          xml_(std::in_place, [this](unsigned char* out, std::size_t capacity) {
              return entry_.read(out, capacity);
          }) {}

    bool next(StreamingRow& row, std::size_t& rowNumber) {
        const auto rowTag = xml_->nextElement("row");
        if (rowTag.empty()) return false;
        const auto indexText = xmlAttribute(rowTag, "r");
        rowNumber = 0;
        if (!indexText.empty() && !parseSize(indexText, rowNumber))
            rowNumber = static_cast<std::size_t>(std::stoull(std::string(indexText)));
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
                if (parseSize(styleValue, sid))
                    cell.styleIndex = sid;
                else
                    cell.styleIndex = static_cast<std::size_t>(std::stoull(std::string(styleValue)));
            }
            const auto type = xmlAttribute(cellTag, "t");
            if (type == "inlineStr") {
                const auto text = xmlText(cellTag, "t");
                cell.value = containsEntity(text) ? xmlUnescape(text) : std::string(text);
            } else if (type == "b") {
                cell.value = xmlText(cellTag, "v") == "1";
            } else if (type == "e") {
                cell.value = cellErrorFromString(std::string(xmlText(cellTag, "v")));
            } else if (type == "s") {
                const auto value = xmlText(cellTag, "v");
                if (!value.empty() && sharedStrings_) {
                    std::size_t index = 0;
                    if (!parseSize(value, index))
                        index = static_cast<std::size_t>(std::stoull(std::string(value)));
                    if (const auto* text = sharedStrings_->lookup(index)) cell.value = *text;
                }
            } else {
                const auto value = xmlText(cellTag, "v");
                if (!value.empty()) {
                    double number = 0.0;
                    if (parseDouble(value, number))
                        cell.value = number;
                    else
                        cell.value = std::stod(std::string(value));
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

std::vector<SheetBinding> readSheetBindings(const xlpp::internal::ZipArchiveReader& archive) {
    const auto workbook = archive.readEntry("xl/workbook.xml");
    std::unordered_map<std::string, std::string> targets;
    if (archive.contains("xl/_rels/workbook.xml.rels")) {
        const auto rels = archive.readEntry("xl/_rels/workbook.xml.rels");
        xlpp::internal::XmlScanner relsScanner(rels);
        std::string_view relationship;
        while (relsScanner.nextElement("Relationship", relationship)) {
            const auto id = xlpp::internal::xmlAttribute(relationship, "Id");
            const auto target = xlpp::internal::xmlAttribute(relationship, "Target");
            targets[std::string(id)] = std::string(target);
        }
    }
    std::vector<SheetBinding> bindings;
    xlpp::internal::XmlScanner wbScanner(workbook);
    std::string_view sheet;
    while (wbScanner.nextElement("sheet", sheet)) {
        const auto nameView = xlpp::internal::xmlAttribute(sheet, "name");
        std::string name = xlpp::internal::xmlUnescape(nameView);
        std::string rid = std::string(xlpp::internal::xmlAttribute(sheet, "r:id"));
        if (rid.empty()) rid = std::string(xlpp::internal::xmlAttribute(sheet, "id"));
        const auto it = targets.find(rid);
        std::string target = it == targets.end() ? std::string{} : it->second;
        if (target.empty()) target = "worksheets/sheet" + std::to_string(bindings.size() + 1) + ".xml";
        if (!target.empty() && target.front() == '/') target.erase(0, 1);
        else if (!target.empty()) target = "xl/" + target;
        bindings.push_back(SheetBinding{std::move(name), std::move(target)});
    }
    return bindings;
}

} // namespace

namespace xlpp {

struct StreamingWorkbookReader::SharedState {
    explicit SharedState(const std::filesystem::path& path) : archive(path) {}
    internal::ZipArchiveReader archive;
    std::vector<SheetBinding> bindings;
    std::shared_ptr<internal::SharedStringsReader> sharedStrings;
};

StreamingWorkbookReader::StreamingWorkbookReader(const std::filesystem::path& path) {
    auto state = std::make_shared<SharedState>(path);
    state->bindings = readSheetBindings(state->archive);
    state_ = std::move(state);
}

std::vector<std::string> StreamingWorkbookReader::worksheetNames() const {
    std::vector<std::string> names;
    names.reserve(state_->bindings.size());
    for (const auto& binding : state_->bindings) names.push_back(binding.name);
    return names;
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
                                                                      state->sharedStrings);
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
