#include <XLPP/Streaming/StreamingWorkbookWriter.h>
#include "../Packaging/ZipArchive.h"
#include "../XML/XmlUtilities.h"
#include <chrono>
#include <iomanip>
#include <list>
#include <sstream>
#include <stdexcept>
#include <unordered_map>
#include <utility>

namespace xlpp {

class SharedStringTable {
public:
    SharedStringTable(SharedStringMode mode, std::size_t capacity)
        : mode_(mode), capacity_(capacity ? capacity : 1) {}

    // Returns the shared-string index for `text`, assigning a fresh index on a
    // cache miss. Counts every occurrence for the sst `count` attribute.
    std::size_t index(std::string_view text) {
        ++occurrences_;
        const std::string key(text);
        if (mode_ == SharedStringMode::Hash) {
            const auto found = hash_.find(key);
            if (found != hash_.end()) return found->second;
            const std::size_t id = strings_.size();
            strings_.push_back(key);
            hash_.emplace(key, id);
            return id;
        }
        const auto found = cache_.find(key);
        if (found != cache_.end()) {
            order_.splice(order_.end(), order_, found->second.position);
            return found->second.index;
        }
        const std::size_t id = strings_.size();
        strings_.push_back(key);
        order_.push_back(key);
        cache_.emplace(key, Entry{id, std::prev(order_.end())});
        if (cache_.size() > capacity_) {
            const std::string& oldest = order_.front();
            cache_.erase(oldest);
            order_.pop_front();
        }
        return id;
    }

    bool enabled() const noexcept { return mode_ != SharedStringMode::Disabled; }
    std::size_t size() const noexcept { return strings_.size(); }
    std::size_t occurrences() const noexcept { return occurrences_; }
    const std::vector<std::string>& strings() const noexcept { return strings_; }

private:
    struct Entry {
        std::size_t index;
        std::list<std::string>::iterator position;
    };
    SharedStringMode mode_;
    std::size_t capacity_;
    std::vector<std::string> strings_;
    std::unordered_map<std::string, std::size_t> hash_;
    std::unordered_map<std::string, Entry> cache_;
    std::list<std::string> order_;
    std::size_t occurrences_{0};
};

std::string sharedStringsXml(const SharedStringTable& table) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\""
        << table.occurrences() << "\" uniqueCount=\"" << table.size() << "\">";
    for (const auto& text : table.strings())
        xml << "<si><t xml:space=\"preserve\">" << internal::xmlEscape(text) << "</t></si>";
    xml << "</sst>";
    return xml.str();
}

} // namespace xlpp

namespace {
std::string cellXml(const xlpp::CellValue& value, std::size_t row, std::size_t column,
                    xlpp::SharedStringTable* sharedStrings, bool date1904) {
    const auto address = xlpp::CellReference{row, column}.address();
    std::ostringstream xml;
    if (std::holds_alternative<std::monostate>(value)) return {};
    if (const auto* text = std::get_if<std::string>(&value)) {
        if (sharedStrings && sharedStrings->enabled()) {
            xml << "<c r=\"" << address << "\" t=\"s\"><v>" << sharedStrings->index(*text)
                << "</v></c>";
        } else {
            xml << "<c r=\"" << address << "\" t=\"inlineStr\"><is><t xml:space=\"preserve\">"
                << xlpp::internal::xmlEscape(*text) << "</t></is></c>";
        }
    } else if (const auto* number = std::get_if<double>(&value)) {
        xml << std::setprecision(17);
        xml << "<c r=\"" << address << "\"><v>" << *number << "</v></c>";
    } else if (const auto* boolean = std::get_if<bool>(&value)) {
        xml << "<c r=\"" << address << "\" t=\"b\"><v>" << (*boolean ? 1 : 0) << "</v></c>";
    } else if (const auto* error = std::get_if<xlpp::CellError>(&value)) {
        xml << "<c r=\"" << address << "\" t=\"e\"><v>" << xlpp::toString(*error) << "</v></c>";
    } else if (const auto* date = std::get_if<xlpp::DateTime>(&value)) {
        xml << std::setprecision(17);
        xml << "<c r=\"" << address << "\"><v>" << xlpp::toExcelSerial(*date, date1904)
            << "</v></c>";
    }
    return xml.str();
}
std::string workbookXml(const std::vector<xlpp::StreamingWorksheetWriter>& sheets, bool date1904) {
    std::ostringstream xml;
    xml << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">";
    if (date1904) xml << "<workbookPr date1904=\"1\"/>";
    xml << "<sheets>";
    for (std::size_t i = 0; i < sheets.size(); ++i)
        xml << "<sheet name=\"" << xlpp::internal::xmlEscape(sheets[i].name()) << "\" sheetId=\""
            << i + 1 << "\" r:id=\"rId" << i + 1 << "\"/>";
    xml << "</sheets></workbook>";
    return xml.str();
}
} // namespace

namespace xlpp {

StreamingWorksheetWriter::StreamingWorksheetWriter(std::string name, std::filesystem::path spoolPath,
                                                   SharedStringTable* sharedStrings, bool date1904)
    : name_(std::move(name)),
      spoolPath_(std::move(spoolPath)),
      stream_(std::make_unique<std::ofstream>(spoolPath_, std::ios::binary)),
      sharedStrings_(sharedStrings),
      date1904_(date1904) {
    if (!*stream_) throw std::runtime_error("Cannot create streaming worksheet spool");
    *stream_ << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>";
}
StreamingWorksheetWriter::StreamingWorksheetWriter(StreamingWorksheetWriter&& other) noexcept
    : name_(std::move(other.name_)), spoolPath_(std::move(other.spoolPath_)), stream_(std::move(other.stream_)),
      sharedStrings_(other.sharedStrings_), date1904_(other.date1904_), rowCount_(other.rowCount_), finished_(other.finished_) { other.finished_ = true; }
StreamingWorksheetWriter& StreamingWorksheetWriter::operator=(StreamingWorksheetWriter&& other) noexcept {
    if (this != &other) {
        try { finish(); } catch (...) {}
        name_ = std::move(other.name_); spoolPath_ = std::move(other.spoolPath_); stream_ = std::move(other.stream_);
        sharedStrings_ = other.sharedStrings_; date1904_ = other.date1904_; rowCount_ = other.rowCount_; finished_ = other.finished_; other.finished_ = true;
    }
    return *this;
}
StreamingWorksheetWriter::~StreamingWorksheetWriter(){try{finish();}catch(...){}}
void StreamingWorksheetWriter::append(const std::vector<CellValue>& row){
    if(finished_)throw std::logic_error("Streaming worksheet is closed");
    ++rowCount_;
    *stream_ << "<row r=\"" << rowCount_ << "\">";
    for(std::size_t i=0;i<row.size();++i)*stream_ << cellXml(row[i],rowCount_,i+1,sharedStrings_,date1904_);
    *stream_ << "</row>";
}
void StreamingWorksheetWriter::finish(){if(finished_)return;if(!stream_){finished_=true;return;}*stream_ << "</sheetData></worksheet>";stream_->flush();stream_->close();finished_=true;}
StreamingWorkbookWriter::StreamingWorkbookWriter(std::filesystem::path outputPath, SharedStringMode sharedStrings, std::size_t lruCapacity)
    : outputPath_(std::move(outputPath)) {
    tempDirectory_ = std::filesystem::temp_directory_path()/("xlpp_stream_"+std::to_string(std::chrono::steady_clock::now().time_since_epoch().count()));
    std::filesystem::create_directories(tempDirectory_);
    if (sharedStrings != SharedStringMode::Disabled)
        sharedStrings_ = std::make_unique<SharedStringTable>(sharedStrings, lruCapacity);
}
StreamingWorkbookWriter::~StreamingWorkbookWriter(){try{if(!closed_)close();}catch(...){}std::error_code ec;std::filesystem::remove_all(tempDirectory_,ec);}
StreamingWorksheetWriter& StreamingWorkbookWriter::addWorksheet(std::string name){
    if(closed_)throw std::logic_error("Streaming workbook is closed");
    if(name.empty())throw std::invalid_argument("Worksheet name cannot be empty");
    const auto path=tempDirectory_/("sheet"+std::to_string(sheets_.size()+1)+".xml");
    sheets_.push_back(StreamingWorksheetWriter(std::move(name),path,sharedStrings_.get(),date1904_));
    return sheets_.back();
}
void StreamingWorkbookWriter::close(){
    if(closed_)return;
    if(sheets_.empty())addWorksheet("Sheet1");
    for(auto& s:sheets_)s.finish();
    const bool hasSharedStrings = sharedStrings_ && sharedStrings_->enabled() && sharedStrings_->size() > 0;
    std::ostringstream contentTypes;
    contentTypes << "<?xml version=\"1.0\" encoding=\"UTF-8\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                 << "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                 << "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                 << "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>";
    for(std::size_t i=0;i<sheets_.size();++i)
        contentTypes << "<Override PartName=\"/xl/worksheets/sheet"+std::to_string(i+1)+".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
    if(hasSharedStrings)
        contentTypes << "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
    contentTypes << "</Types>";
    internal::ZipArchive zip;
    zip.setCompressionLevel(compressionLevel_ == CompressionLevel::Store ? 0
                            : compressionLevel_ == CompressionLevel::Fastest ? 1
                            : compressionLevel_ == CompressionLevel::Best ? 9 : -1);
    zip.setCompressionStrategy(compressionStrategy_ == CompressionStrategy::Filtered ? 1
                               : compressionStrategy_ == CompressionStrategy::HuffmanOnly ? 2
                               : compressionStrategy_ == CompressionStrategy::Rle ? 3
                               : compressionStrategy_ == CompressionStrategy::Fixed ? 4 : 0);
    zip.setParallelWorkers(parallelWorkers_);
    zip.add("[Content_Types].xml", contentTypes.str());
    zip.add("_rels/.rels","<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");
    zip.add("xl/workbook.xml",workbookXml(sheets_,date1904_));
    std::string rels="<?xml version=\"1.0\" encoding=\"UTF-8\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    for(std::size_t i=0;i<sheets_.size();++i)
        rels+="<Relationship Id=\"rId"+std::to_string(i+1)+"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"+std::to_string(i+1)+".xml\"/>";
    if(hasSharedStrings) {
        rels += "<Relationship Id=\"rId"+std::to_string(sheets_.size()+1)+"\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>";
        zip.add("xl/sharedStrings.xml",sharedStringsXml(*sharedStrings_));
    }
    rels+="</Relationships>";
    zip.add("xl/_rels/workbook.xml.rels",std::move(rels));
    for(std::size_t i=0;i<sheets_.size();++i)zip.addFile("xl/worksheets/sheet"+std::to_string(i+1)+".xml",sheets_[i].spoolPath_);
    zip.save(outputPath_);
    closed_=true;
}

} // namespace xlpp
