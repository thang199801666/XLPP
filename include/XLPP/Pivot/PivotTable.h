#pragma once
#include <algorithm>
#include <cstdint>
#include <cstddef>
#include <optional>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

namespace xlpp {

// Physical value type retained by PivotCache records. Unlike the legacy
// string-only cache surface, this metadata distinguishes values that have the
// same display text (for example numeric 123 vs string "123") and preserves
// date/time records as SpreadsheetML <d> values.
enum class PivotCacheValueKind { Missing, Number, String, Boolean, Error, DateTime };

// OLAP pivot source metadata. XL++ models the cache-source identity and the
// common olapInfo fields so callers can inspect and selectively patch an OLAP
// pivot without regenerating the physical cache. Deeper OLAP semantics
// (olapKpi, MDX calculated members, slicer/timeline parts) remain
// preservation-backed: untouched XML is carried through byte-for-byte.
struct PivotOlapSourceInfo {
    // cacheSource@type; empty for worksheet-backed caches.
    std::string sourceType;
    // olapInfo@preserveFormatting (defaults true for OLAP per SpreadsheetML).
    bool preserveFormatting{true};
    // olapInfo@localCube / @localConnection for local OLAP caches.
    std::string localCube;
    std::string localConnection;
    // p:connection ID when the cache binds to a workbook data connection.
    int connectionId{-1};
    // Raw <olapInfo> subtree for lossless carry-through when it is not edited.
    std::string rawOlapInfoXml;
    bool present() const noexcept { return !sourceType.empty(); }
};

// One SpreadsheetML <calculatedMember> node (OLAP calculated members or
// measure sets). Kept lossless so the surrounding cache definition can be
// patched without re-serializing unmodeled MDX content.
struct PivotCalculatedMember {
    std::string name;
    std::string mdx;
    std::string memberName;
    int hierarchy{-1};
    std::string solveOrder;
    std::string set;
    // Raw <calculatedMember> subtree used for byte-preserving writes.
    std::string rawXml;
};

struct PivotGroupItem {
    // Physical SpreadsheetML group-item kind. groupItems may mix number (<n>),
    // date (<d>), string (<s>) and missing (<m>) entries; preserving the kind
    // keeps date/number bins distinct from plain labels during round-trips.
    PivotCacheValueKind kind{PivotCacheValueKind::String};
    std::string value;
};

struct PivotFieldGroup {
    int parentField{-1};
    int baseField{-1};
    std::string groupBy;
    bool autoStart{true};
    bool autoEnd{true};
    std::optional<double> startNumber;
    std::optional<double> endNumber;
    std::optional<double> interval;
    std::string startDate;
    std::string endDate;
    // Legacy text-only view of groupItems; a new item appended here is treated
    // as a string kind. Prefer typedGroupItems()/addTypedGroupItem() when exact
    // date/number bin preservation matters.
    std::vector<std::string> items;
    std::vector<PivotGroupItem> typedItems;
    const PivotGroupItem* typedGroupItem(std::size_t index) const noexcept {
        return index < typedItems.size() ? &typedItems[index] : nullptr;
    }
    void addTypedGroupItem(PivotCacheValueKind kind, std::string value) {
        typedItems.push_back({kind, std::move(value)});
        items.push_back(typedItems.back().value);
    }
};

class PivotCache {
public:
    int cacheId() const noexcept { return cacheId_; }
    void setCacheId(int value) {
        if (value <= 0) throw std::invalid_argument("Pivot cache ID must be positive");
        cacheId_ = value;
    }

    const std::string& sourceData() const noexcept { return sourceData_; }
    void setSourceData(std::string value) { sourceData_ = std::move(value); }
    const std::string& sourceName() const noexcept { return sourceName_; }
    void setSourceName(std::string value) { sourceName_ = std::move(value); }

    // Optional identity used when multiple PivotTables intentionally share one
    // physical PivotCache. Imported workbooks populate this from the cache part
    // path; generated workbooks only share caches when callers opt in by using
    // the same non-empty key on compatible cache models.
    const std::string& sharedCacheKey() const noexcept { return sharedCacheKey_; }
    void setSharedCacheKey(std::string value) { sharedCacheKey_ = std::move(value); }

    bool refreshOnLoad() const noexcept { return refreshOnLoad_; }
    void setRefreshOnLoad(bool value) noexcept { refreshOnLoad_ = value; }
    bool saveData() const noexcept { return saveData_; }
    void setSaveData(bool value) noexcept { saveData_ = value; }
    bool enableRefresh() const noexcept { return enableRefresh_; }
    void setEnableRefresh(bool value) noexcept { enableRefresh_ = value; }
    int missingItemsLimit() const noexcept { return missingItemsLimit_; }
    void setMissingItemsLimit(int value) noexcept { missingItemsLimit_ = value; }

    const std::vector<std::string>& fields() const noexcept { return fields_; }
    std::vector<std::string>& fields() noexcept { return fields_; }
    void setFields(std::vector<std::string> value) {
        if (!records_.empty()) {
            for (const auto& record : records_) {
                if (record.size() != value.size())
                    throw std::invalid_argument("Pivot cache record width must match field count");
            }
        }
        fields_ = std::move(value);
        fieldFormulas_.resize(fields_.size());
        fieldGroups_.resize(fields_.size());
        fieldCaptions_.resize(fields_.size());
        fieldNumberFormatIds_.resize(fields_.size());
        fieldDatabaseFields_.resize(fields_.size(), true);
    }
    void addField(std::string value) {
        if (!records_.empty())
            throw std::logic_error("Cannot add a pivot cache field after records have been added");
        fields_.push_back(std::move(value));
        fieldFormulas_.emplace_back();
        fieldGroups_.emplace_back();
        fieldCaptions_.emplace_back();
        fieldNumberFormatIds_.push_back(0);
        fieldDatabaseFields_.push_back(true);
    }

    // Calculated Pivot fields are represented by a cacheField formula. XL++
    // keeps record geometry rectangular by appending a blank cached value to
    // existing records when a calculated field is introduced after data load.
    int addCalculatedField(std::string name, std::string formula) {
        if (name.empty()) throw std::invalid_argument("Calculated pivot field name cannot be empty");
        if (formula.empty()) throw std::invalid_argument("Calculated pivot field formula cannot be empty");
        if (fieldIndex(name) >= 0) throw std::invalid_argument("Duplicate pivot cache field: " + name);
        fields_.push_back(std::move(name));
        fieldFormulas_.resize(fields_.size());
        fieldGroups_.resize(fields_.size());
        fieldCaptions_.resize(fields_.size());
        fieldNumberFormatIds_.resize(fields_.size());
        fieldDatabaseFields_.resize(fields_.size(), true);
        fieldFormulas_.back() = std::move(formula);
        fieldDatabaseFields_.back() = false;
        const bool typedRecords = hasTypedRecordKinds();
        for (auto& record : records_) record.emplace_back();
        if (typedRecords)
            for (auto& kinds : recordKinds_) kinds.push_back(PivotCacheValueKind::Missing);
        return static_cast<int>(fields_.size() - 1);
    }

    const std::string& fieldFormula(std::size_t index) const {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        static const std::string empty;
        return index < fieldFormulas_.size() ? fieldFormulas_[index] : empty;
    }
    void setFieldFormula(std::size_t index, std::string formula) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        fieldFormulas_.resize(fields_.size());
        fieldDatabaseFields_.resize(fields_.size(), true);
        fieldFormulas_[index] = std::move(formula);
        if (!fieldFormulas_[index].empty()) fieldDatabaseFields_[index] = false;
    }
    bool isCalculatedField(std::size_t index) const { return !fieldFormula(index).empty(); }

    const PivotFieldGroup* tryFieldGroup(std::size_t index) const noexcept {
        return index < fieldGroups_.size() && fieldGroups_[index].has_value() ? &*fieldGroups_[index] : nullptr;
    }
    PivotFieldGroup& fieldGroup(std::size_t index) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        fieldGroups_.resize(fields_.size());
        if (!fieldGroups_[index]) fieldGroups_[index].emplace();
        return *fieldGroups_[index];
    }
    void setFieldGroup(std::size_t index, PivotFieldGroup group) { fieldGroup(index) = std::move(group); }
    void clearFieldGroup(std::size_t index) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        fieldGroups_.resize(fields_.size());
        fieldGroups_[index].reset();
    }

    void setNumericFieldGrouping(std::size_t index, double start, double end, double interval,
                                 bool autoStart = false, bool autoEnd = false) {
        if (interval <= 0.0) throw std::invalid_argument("Pivot numeric grouping interval must be positive");
        if (end < start) throw std::invalid_argument("Pivot numeric grouping end must not precede start");
        PivotFieldGroup group;
        group.baseField = static_cast<int>(index);
        group.autoStart = autoStart;
        group.autoEnd = autoEnd;
        group.startNumber = start;
        group.endNumber = end;
        group.interval = interval;
        setFieldGroup(index, std::move(group));
    }

    void setDateFieldGrouping(std::size_t index, std::string groupBy,
                              std::string startDate = {}, std::string endDate = {},
                              bool autoStart = true, bool autoEnd = true) {
        static const std::vector<std::string> supported{
            "seconds", "minutes", "hours", "days", "months", "quarters", "years"
        };
        if (std::find(supported.begin(), supported.end(), groupBy) == supported.end())
            throw std::invalid_argument("Unsupported pivot date grouping unit: " + groupBy);
        PivotFieldGroup group;
        group.baseField = static_cast<int>(index);
        group.groupBy = std::move(groupBy);
        group.autoStart = autoStart;
        group.autoEnd = autoEnd;
        group.startDate = std::move(startDate);
        group.endDate = std::move(endDate);
        setFieldGroup(index, std::move(group));
    }

    const std::string& fieldCaption(std::size_t index) const {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        static const std::string empty;
        return index < fieldCaptions_.size() ? fieldCaptions_[index] : empty;
    }
    void setFieldCaption(std::size_t index, std::string value) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        fieldCaptions_.resize(fields_.size());
        fieldCaptions_[index] = std::move(value);
    }
    int fieldNumberFormatId(std::size_t index) const {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        return index < fieldNumberFormatIds_.size() ? fieldNumberFormatIds_[index] : 0;
    }
    void setFieldNumberFormatId(std::size_t index, int value) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        if (value < 0) throw std::invalid_argument("Pivot cache field number format ID cannot be negative");
        fieldNumberFormatIds_.resize(fields_.size());
        fieldNumberFormatIds_[index] = value;
    }
    bool fieldDatabaseField(std::size_t index) const {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        return index < fieldDatabaseFields_.size() ? fieldDatabaseFields_[index] : true;
    }
    void setFieldDatabaseField(std::size_t index, bool value) {
        if (index >= fields_.size()) throw std::out_of_range("Pivot cache field index is out of range");
        fieldDatabaseFields_.resize(fields_.size(), true);
        fieldDatabaseFields_[index] = value;
    }

    int fieldIndex(const std::string& name) const noexcept {
        const auto it = std::find(fields_.begin(), fields_.end(), name);
        return it == fields_.end() ? -1 : static_cast<int>(std::distance(fields_.begin(), it));
    }

    const std::vector<std::vector<std::string>>& records() const noexcept { return records_; }
    // Direct mutable access keeps the legacy API but invalidates explicit type
    // metadata because arbitrary string replacement cannot communicate the
    // intended physical Pivot value type. Call setRecordValue()/addTypedRecord()
    // when exact type preservation matters.
    std::vector<std::vector<std::string>>& records() noexcept { recordKindsValid_ = false; return records_; }
    const std::vector<std::vector<PivotCacheValueKind>>& recordKinds() const noexcept { return recordKinds_; }
    bool hasTypedRecordKinds() const noexcept {
        if (!recordKindsValid_ || recordKinds_.size() != records_.size()) return false;
        for (std::size_t i = 0; i < records_.size(); ++i)
            if (recordKinds_[i].size() != records_[i].size()) return false;
        return true;
    }
    PivotCacheValueKind recordKind(std::size_t row, std::size_t column) const {
        if (!hasTypedRecordKinds() || row >= recordKinds_.size() || column >= recordKinds_[row].size())
            throw std::out_of_range("Pivot cache record kind is not available");
        return recordKinds_[row][column];
    }
    void setRecords(std::vector<std::vector<std::string>> value) {
        if (!fields_.empty()) {
            for (const auto& record : value) {
                if (record.size() != fields_.size())
                    throw std::invalid_argument("Pivot cache record width must match field count");
            }
        }
        records_ = std::move(value);
        recordKinds_.clear();
        recordKindsValid_ = false;
    }
    void setTypedRecords(std::vector<std::vector<std::string>> values,
                         std::vector<std::vector<PivotCacheValueKind>> kinds) {
        if (values.size() != kinds.size())
            throw std::invalid_argument("Pivot typed record row count must match record row count");
        for (std::size_t row = 0; row < values.size(); ++row) {
            if ((!fields_.empty() && values[row].size() != fields_.size()) || kinds[row].size() != values[row].size())
                throw std::invalid_argument("Pivot typed record geometry must match cache fields");
        }
        records_ = std::move(values);
        recordKinds_ = std::move(kinds);
        recordKindsValid_ = true;
    }
    void addRecord(std::vector<std::string> value) {
        if (!fields_.empty() && value.size() != fields_.size())
            throw std::invalid_argument("Pivot cache record width must match field count");
        records_.push_back(std::move(value));
        recordKinds_.clear();
        recordKindsValid_ = false;
    }
    void addTypedRecord(std::vector<std::string> value, std::vector<PivotCacheValueKind> kinds) {
        if ((!fields_.empty() && value.size() != fields_.size()) || kinds.size() != value.size())
            throw std::invalid_argument("Pivot typed record geometry must match cache fields");
        if (!recordKindsValid_ && !records_.empty())
            throw std::logic_error("Cannot append an explicitly typed Pivot record after untyped records");
        records_.push_back(std::move(value));
        recordKinds_.push_back(std::move(kinds));
        recordKindsValid_ = true;
    }
    void setRecordValue(std::size_t row, std::size_t column, std::string value, PivotCacheValueKind kind) {
        if (row >= records_.size() || column >= records_[row].size())
            throw std::out_of_range("Pivot cache record coordinate is out of range");
        if (!hasTypedRecordKinds()) {
            recordKinds_.assign(records_.size(), {});
            for (std::size_t r = 0; r < records_.size(); ++r)
                recordKinds_[r].assign(records_[r].size(), PivotCacheValueKind::String);
            recordKindsValid_ = true;
        }
        records_[row][column] = std::move(value);
        recordKinds_[row][column] = kind;
    }
    void clearRecords() noexcept { records_.clear(); recordKinds_.clear(); recordKindsValid_ = false; }

    // OLAP source identity and calculated members. Worksheet-backed caches keep
    // olap() absent; OLAP caches retain the olapInfo/calculatedMember XML so a
    // selective patch does not truncate unmodeled MDX or connection metadata.
    const PivotOlapSourceInfo* olap() const noexcept { return olap_ ? &*olap_ : nullptr; }
    PivotOlapSourceInfo& olap() {
        if (!olap_) olap_.emplace();
        return *olap_;
    }
    void setOlap(PivotOlapSourceInfo info) { olap_ = std::move(info); }
    void clearOlap() noexcept { olap_.reset(); }
    const std::vector<PivotCalculatedMember>& calculatedMembers() const noexcept { return calculatedMembers_; }
    std::vector<PivotCalculatedMember>& calculatedMembers() noexcept { return calculatedMembers_; }
    void setCalculatedMembers(std::vector<PivotCalculatedMember> members) { calculatedMembers_ = std::move(members); }
    void clearCalculatedMembers() noexcept { calculatedMembers_.clear(); }

    // Structural worksheet edits use these helpers to keep a worksheet-backed
    // PivotCache rectangular when source columns are inserted or deleted.
    // They preserve per-field metadata, update fieldGroup field indices, and
    // insert/erase the corresponding value in every cached record.
    void insertSourceFields(std::size_t index, std::size_t count,
                            const std::string& namePrefix = "Column") {
        if (count == 0) return;
        if (index > fields_.size()) throw std::out_of_range("Pivot cache field insertion index is out of range");
        fieldFormulas_.resize(fields_.size());
        fieldGroups_.resize(fields_.size());
        fieldCaptions_.resize(fields_.size());
        fieldNumberFormatIds_.resize(fields_.size());
        fieldDatabaseFields_.resize(fields_.size(), true);

        std::vector<std::string> names;
        names.reserve(count);
        for (std::size_t i = 0; i < count; ++i) {
            std::string candidate = namePrefix + std::to_string(index + i + 1);
            if (candidate.empty()) candidate = "Column" + std::to_string(index + i + 1);
            std::string unique = candidate;
            std::size_t suffix = 2;
            while (std::find(fields_.begin(), fields_.end(), unique) != fields_.end() ||
                   std::find(names.begin(), names.end(), unique) != names.end())
                unique = candidate + "_" + std::to_string(suffix++);
            names.push_back(std::move(unique));
        }

        fields_.insert(fields_.begin() + static_cast<std::ptrdiff_t>(index), names.begin(), names.end());
        fieldFormulas_.insert(fieldFormulas_.begin() + static_cast<std::ptrdiff_t>(index), count, {});
        fieldGroups_.insert(fieldGroups_.begin() + static_cast<std::ptrdiff_t>(index), count, std::nullopt);
        fieldCaptions_.insert(fieldCaptions_.begin() + static_cast<std::ptrdiff_t>(index), count, {});
        fieldNumberFormatIds_.insert(fieldNumberFormatIds_.begin() + static_cast<std::ptrdiff_t>(index), count, 0);
        fieldDatabaseFields_.insert(fieldDatabaseFields_.begin() + static_cast<std::ptrdiff_t>(index), count, true);
        const bool typedRecords = hasTypedRecordKinds();
        for (auto& record : records_)
            record.insert(record.begin() + static_cast<std::ptrdiff_t>(index), count, {});
        if (typedRecords)
            for (auto& kinds : recordKinds_)
                kinds.insert(kinds.begin() + static_cast<std::ptrdiff_t>(index), count, PivotCacheValueKind::Missing);

        for (auto& group : fieldGroups_) {
            if (!group) continue;
            if (group->baseField >= static_cast<int>(index)) group->baseField += static_cast<int>(count);
            if (group->parentField >= static_cast<int>(index)) group->parentField += static_cast<int>(count);
        }
    }

    void eraseSourceFields(std::size_t index, std::size_t count) {
        if (count == 0) return;
        if (index >= fields_.size() || count > fields_.size() - index)
            throw std::out_of_range("Pivot cache field erase range is out of range");
        fieldFormulas_.resize(fields_.size());
        fieldGroups_.resize(fields_.size());
        fieldCaptions_.resize(fields_.size());
        fieldNumberFormatIds_.resize(fields_.size());
        fieldDatabaseFields_.resize(fields_.size(), true);
        const auto first = static_cast<std::ptrdiff_t>(index);
        const auto last = static_cast<std::ptrdiff_t>(index + count);
        fields_.erase(fields_.begin() + first, fields_.begin() + last);
        fieldFormulas_.erase(fieldFormulas_.begin() + first, fieldFormulas_.begin() + last);
        fieldGroups_.erase(fieldGroups_.begin() + first, fieldGroups_.begin() + last);
        fieldCaptions_.erase(fieldCaptions_.begin() + first, fieldCaptions_.begin() + last);
        fieldNumberFormatIds_.erase(fieldNumberFormatIds_.begin() + first, fieldNumberFormatIds_.begin() + last);
        fieldDatabaseFields_.erase(fieldDatabaseFields_.begin() + first, fieldDatabaseFields_.begin() + last);
        const bool typedRecords = hasTypedRecordKinds();
        for (auto& record : records_) {
            if (record.size() < index + count)
                throw std::logic_error("Pivot cache record width is smaller than field erase range");
            record.erase(record.begin() + first, record.begin() + last);
        }
        if (typedRecords)
            for (auto& kinds : recordKinds_) kinds.erase(kinds.begin() + first, kinds.begin() + last);

        const int cutBegin = static_cast<int>(index);
        const int cutEnd = static_cast<int>(index + count);
        const int delta = static_cast<int>(count);
        auto repair = [&](int& field) {
            if (field < 0) return;
            if (field >= cutEnd) field -= delta;
            else if (field >= cutBegin) field = -1;
        };
        for (auto& group : fieldGroups_) {
            if (!group) continue;
            repair(group->baseField);
            repair(group->parentField);
        }
    }

private:
    int cacheId_{1};
    std::string sourceData_;
    std::string sourceName_;
    std::string sharedCacheKey_;
    bool refreshOnLoad_{true};
    bool saveData_{true};
    bool enableRefresh_{true};
    int missingItemsLimit_{-1};
    std::vector<std::string> fields_;
    std::vector<std::string> fieldFormulas_;
    std::vector<std::optional<PivotFieldGroup>> fieldGroups_;
    std::vector<std::string> fieldCaptions_;
    std::vector<int> fieldNumberFormatIds_;
    std::vector<bool> fieldDatabaseFields_;
    std::vector<std::vector<std::string>> records_;
    std::vector<std::vector<PivotCacheValueKind>> recordKinds_;
    bool recordKindsValid_{false};
    std::optional<PivotOlapSourceInfo> olap_;
    std::vector<PivotCalculatedMember> calculatedMembers_;
};

struct PivotFieldItem {
    int cacheIndex{-1};
    std::string type;
    std::string caption;
    bool hidden{false};
    bool showDetails{true};
    bool formula{false};
    bool missing{false};
    // P1T stable logical binding. Pivot shared-item indexes are physical and
    // can change when cache records are edited/compacted. When present, XL++
    // resolves cacheIndex again from (kind,value) at save time so hidden/item
    // semantics follow the same logical value instead of a stale ordinal.
    bool hasCacheValue{false};
    std::string cacheValue;
    PivotCacheValueKind cacheValueKind{PivotCacheValueKind::String};
    void bindCacheValue(std::string value, PivotCacheValueKind kind) {
        cacheValue = std::move(value);
        cacheValueKind = kind;
        hasCacheValue = true;
    }
    void clearCacheValueBinding() noexcept { hasCacheValue = false; cacheValue.clear(); }
    // Raw original <item> node retained for consumers that need attributes not
    // yet promoted to the high-level model. Generated XML uses the fields above.
    std::string rawXml;
};

class PivotField {
public:
    PivotField() = default;
    explicit PivotField(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& axis() const noexcept { return axis_; }
    void setAxis(std::string value) { axis_ = std::move(value); }
    bool showAll() const noexcept { return showAll_; }
    void setShowAll(bool value) noexcept { showAll_ = value; }
    int sortType() const noexcept { return sortType_; }
    void setSortType(int value) noexcept { sortType_ = value; }
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }
    bool compact() const noexcept { return compact_; }
    void setCompact(bool value) noexcept { compact_ = value; }
    bool outline() const noexcept { return outline_; }
    void setOutline(bool value) noexcept { outline_ = value; }
    bool defaultSubtotal() const noexcept { return defaultSubtotal_; }
    void setDefaultSubtotal(bool value) noexcept { defaultSubtotal_ = value; }
    const std::vector<std::string>& subtotals() const noexcept { return subtotals_; }
    std::vector<std::string>& subtotals() noexcept { return subtotals_; }
    void addSubtotal(std::string value) {
        static const std::vector<std::string> supported{
            "sum", "countA", "avg", "max", "min", "product", "count",
            "stdDev", "stdDevP", "var", "varP"
        };
        if (std::find(supported.begin(), supported.end(), value) == supported.end())
            throw std::invalid_argument("Unsupported pivot field subtotal: " + value);
        if (std::find(subtotals_.begin(), subtotals_.end(), value) == subtotals_.end())
            subtotals_.push_back(std::move(value));
    }
    void clearSubtotals() noexcept { subtotals_.clear(); }
    const std::vector<int>& hiddenItems() const noexcept { return hiddenItems_; }
    std::vector<int>& hiddenItems() noexcept { return hiddenItems_; }
    void hideItem(int index) {
        if (index < 0) return;
        if (std::find(hiddenItems_.begin(), hiddenItems_.end(), index) == hiddenItems_.end()) hiddenItems_.push_back(index);
        for (auto& item : items_) if (item.cacheIndex == index) item.hidden = true;
    }
    const std::vector<PivotFieldItem>& items() const noexcept { return items_; }
    std::vector<PivotFieldItem>& items() noexcept { return items_; }
    PivotFieldItem& addItem(PivotFieldItem item) {
        if (item.cacheIndex < -1) throw std::invalid_argument("Pivot field item cache index is invalid");
        if (item.hidden && item.cacheIndex >= 0
            && std::find(hiddenItems_.begin(), hiddenItems_.end(), item.cacheIndex) == hiddenItems_.end())
            hiddenItems_.push_back(item.cacheIndex);
        items_.push_back(std::move(item));
        return items_.back();
    }
    // P1T logical item API: bind field semantics to the typed cache value
    // rather than to a physical shared-item ordinal that may change when the
    // PivotCache is edited or compacted. The serializer resolves the current
    // physical x index at save time.
    PivotFieldItem& addCacheValueItem(std::string value, PivotCacheValueKind kind, bool hidden = false) {
        auto found = std::find_if(items_.begin(), items_.end(), [&](const PivotFieldItem& item) {
            return item.hasCacheValue && item.cacheValue == value && item.cacheValueKind == kind;
        });
        if (found != items_.end()) {
            found->hidden = hidden;
            return *found;
        }
        PivotFieldItem item;
        item.hidden = hidden;
        item.bindCacheValue(std::move(value), kind);
        items_.push_back(std::move(item));
        return items_.back();
    }
    PivotFieldItem& hideCacheValue(std::string value, PivotCacheValueKind kind = PivotCacheValueKind::String) {
        return addCacheValueItem(std::move(value), kind, true);
    }
    PivotFieldItem& showCacheValue(std::string value, PivotCacheValueKind kind = PivotCacheValueKind::String) {
        return addCacheValueItem(std::move(value), kind, false);
    }
    void clearItems() noexcept { items_.clear(); hiddenItems_.clear(); }

private:
    std::string name_;
    std::string axis_{"axisRow"};
    bool showAll_{false};
    int sortType_{0};
    int fieldIndex_{-1};
    bool compact_{true};
    bool outline_{true};
    bool defaultSubtotal_{false};
    std::vector<std::string> subtotals_;
    std::vector<int> hiddenItems_;
    std::vector<PivotFieldItem> items_;
};

class PivotFieldReference {
public:
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }
    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& displayName() const noexcept { return displayName_; }
    void setDisplayName(std::string value) { displayName_ = std::move(value); }
    const std::string& subtotal() const noexcept { return subtotal_; }
    void setSubtotal(std::string value) {
        static const std::vector<std::string> supported{
            "average", "count", "countNums", "max", "min", "product",
            "stdDev", "stdDevp", "sum", "var", "varp"
        };
        if (std::find(supported.begin(), supported.end(), value) == supported.end())
            throw std::invalid_argument("Unsupported pivot data aggregation: " + value);
        subtotal_ = std::move(value);
    }
    const std::string& showDataAs() const noexcept { return showDataAs_; }
    void setShowDataAs(std::string value) {
        static const std::vector<std::string> supported{
            "normal", "difference", "percent", "percentDiff", "runTotal",
            "percentOfRow", "percentOfCol", "percentOfTotal", "index"
        };
        if (std::find(supported.begin(), supported.end(), value) == supported.end())
            throw std::invalid_argument("Unsupported pivot showDataAs mode: " + value);
        showDataAs_ = std::move(value);
    }
    int baseField() const noexcept { return baseField_; }
    void setBaseField(int value) noexcept { baseField_ = value; }
    int baseItem() const noexcept { return baseItem_; }
    void setBaseItem(int value) noexcept { baseItem_ = value; }
    int numberFormatId() const noexcept { return numberFormatId_; }
    void setNumberFormatId(int value) noexcept { numberFormatId_ = value; }

private:
    int fieldIndex_{-1};
    std::string name_;
    std::string subtotal_{"sum"};
    std::string displayName_;
    std::string showDataAs_{"normal"};
    int baseField_{0};
    int baseItem_{0};
    int numberFormatId_{0};
};

class PivotPageField {
public:
    int fieldIndex() const noexcept { return fieldIndex_; }
    void setFieldIndex(int value) noexcept { fieldIndex_ = value; }
    int item() const noexcept { return item_; }
    void setItem(int value) noexcept { item_ = value; }
    int hierarchy() const noexcept { return hierarchy_; }
    void setHierarchy(int value) noexcept { hierarchy_ = value; }
    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
private:
    int fieldIndex_{-1};
    int item_{-1};
    int hierarchy_{-1};
    std::string name_;
};



struct PivotCacheFieldPatch {
    std::optional<std::string> name;
    std::optional<std::string> caption;
    std::optional<std::string> formula;
    std::optional<int> numberFormatId;
    std::optional<bool> databaseField;
};

struct PivotFieldItemPatch {
    std::optional<int> cacheIndex;
    std::optional<std::string> type;
    std::optional<std::string> caption;
    std::optional<bool> hidden;
    std::optional<bool> showDetails;
    std::optional<bool> formula;
    std::optional<bool> missing;
};

struct PivotFilterPatch {
    std::optional<int> fieldIndex;
    std::optional<std::string> type;
    std::optional<std::uint32_t> id;
    std::optional<int> evaluationOrder;
    std::optional<int> measureField;
    std::optional<int> measureHierarchy;
    std::optional<int> memberPropertyField;
    std::optional<std::string> name;
    std::optional<std::string> description;
    std::optional<std::string> stringValue1;
    std::optional<std::string> stringValue2;
    // Replaces the nested <autoFilter> subtree verbatim when supplied. Empty
    // removes it. Callers remain responsible for supplying valid SpreadsheetML.
    std::optional<std::string> autoFilterXml;
};

struct PivotDataFieldPatch {
    std::optional<int> fieldIndex;
    std::optional<std::string> name;
    std::optional<std::string> subtotal;
    std::optional<std::string> showDataAs;
    std::optional<int> baseField;
    std::optional<int> baseItem;
    std::optional<int> numberFormatId;
};

struct PivotPageFieldPatch {
    std::optional<int> fieldIndex;
    std::optional<int> item;
    std::optional<int> hierarchy;
    std::optional<std::string> name;
};

enum class PivotCacheRecordValueType {
    Missing,
    Number,
    String,
    Boolean,
    Error,
    DateTime,
    SharedItem
};

struct PivotCacheRecordValuePatch {
    PivotCacheRecordValueType type{PivotCacheRecordValueType::String};
    // Logical value exposed by PivotCache::records(). For SharedItem this is
    // the resolved shared-item text while sharedItemIndex controls OOXML x@v.
    std::string value;
    int sharedItemIndex{-1};
};

struct PivotChartLinkIssue {
    std::string worksheetName;
    std::string chartId;
    std::string pivotTableName;
    std::string message;
};

struct PivotChartLinkValidationReport {
    std::size_t pivotChartsVisited{0};
    std::size_t validLinks{0};
    std::vector<PivotChartLinkIssue> issues;
    bool ok() const noexcept { return issues.empty(); }
};

struct PivotCacheOptionsPatch {
    std::optional<bool> refreshOnLoad;
    std::optional<bool> saveData;
    std::optional<bool> enableRefresh;
    std::optional<int> missingItemsLimit;
};

// Selective OLAP cache-source/olapInfo patch. Fields not requested keep their
// original bytes; requested scalar attributes are rewritten in place and the
// unmodeled olapInfo children remain byte-preserved.
struct PivotOlapSourcePatch {
    std::optional<bool> preserveFormatting;
    std::optional<std::string> localCube;
    std::optional<std::string> localConnection;
    std::optional<int> connectionId;
};

struct PivotCalculatedMemberPatch {
    std::optional<std::string> mdx;
    std::optional<std::string> memberName;
    std::optional<int> hierarchy;
    std::optional<std::string> solveOrder;
    std::optional<std::string> set;
};

struct PivotChartFormat {
    // Matches SpreadsheetML chartFormat@chart. DrawingML c:pivotSource/c:fmtId
    // links back to this chart index for PivotChart-specific formatting.
    std::uint32_t chartIndex{0};
    std::uint32_t formatId{0};
    bool series{false};
    // Raw <pivotArea> subtree. PivotArea has a broad selector grammar; keeping
    // it lossless lets XL++ edit/link PivotCharts without truncating selectors.
    std::string pivotAreaXml;
};

struct PivotFilter {
    int fieldIndex{-1};
    std::string type;
    std::uint32_t id{0};
    int evaluationOrder{0};
    int measureField{-1};
    int measureHierarchy{-1};
    int memberPropertyField{-1};
    std::string name;
    std::string description;
    std::string stringValue1;
    std::string stringValue2;
    // Raw SpreadsheetML <autoFilter> subtree. This keeps advanced filter
    // operator payloads lossless while the high-level model focuses on the
    // PivotFilter attributes that control field/measure semantics.
    std::string autoFilterXml;
};

class PivotTable {
public:
    PivotTable() = default;
    explicit PivotTable(std::string name) : name_(std::move(name)) {}

    const std::string& name() const noexcept { return name_; }
    void setName(std::string value) { name_ = std::move(value); }
    const std::string& location() const noexcept { return location_; }
    void setLocation(std::string value) { location_ = std::move(value); }

    bool rowGrandTotals() const noexcept { return rowGrandTotals_; }
    void setRowGrandTotals(bool value) noexcept { rowGrandTotals_ = value; }
    bool columnGrandTotals() const noexcept { return columnGrandTotals_; }
    void setColumnGrandTotals(bool value) noexcept { columnGrandTotals_ = value; }
    bool compact() const noexcept { return compact_; }
    void setCompact(bool value) noexcept { compact_ = value; }
    bool outline() const noexcept { return outline_; }
    void setOutline(bool value) noexcept { outline_ = value; }
    bool preserveFormatting() const noexcept { return preserveFormatting_; }
    void setPreserveFormatting(bool value) noexcept { preserveFormatting_ = value; }
    bool useAutoFormatting() const noexcept { return useAutoFormatting_; }
    void setUseAutoFormatting(bool value) noexcept { useAutoFormatting_ = value; }
    bool showDrill() const noexcept { return showDrill_; }
    void setShowDrill(bool value) noexcept { showDrill_ = value; }
    bool multipleFieldFilters() const noexcept { return multipleFieldFilters_; }
    void setMultipleFieldFilters(bool value) noexcept { multipleFieldFilters_ = value; }
    const std::string& styleName() const noexcept { return styleName_; }
    void setStyleName(std::string value) { styleName_ = std::move(value); }
    bool showRowHeaders() const noexcept { return showRowHeaders_; }
    void setShowRowHeaders(bool value) noexcept { showRowHeaders_ = value; }
    bool showColumnHeaders() const noexcept { return showColumnHeaders_; }
    void setShowColumnHeaders(bool value) noexcept { showColumnHeaders_ = value; }
    bool showRowStripes() const noexcept { return showRowStripes_; }
    void setShowRowStripes(bool value) noexcept { showRowStripes_ = value; }
    bool showColumnStripes() const noexcept { return showColumnStripes_; }
    void setShowColumnStripes(bool value) noexcept { showColumnStripes_ = value; }
    const std::optional<std::uint32_t>& chartFormatIndex() const noexcept { return chartFormatIndex_; }
    void setChartFormatIndex(std::uint32_t value) noexcept { chartFormatIndex_ = value; }
    void clearChartFormatIndex() noexcept { chartFormatIndex_.reset(); }

    PivotCache& cache() noexcept { return cache_; }
    const PivotCache& cache() const noexcept { return cache_; }

    PivotField& addRowField(std::string name) {
        rowFields_.emplace_back(std::move(name));
        rowFields_.back().setAxis("axisRow");
        rowFields_.back().setFieldIndex(cache_.fieldIndex(rowFields_.back().name()));
        return rowFields_.back();
    }
    PivotField& addColumnField(std::string name) {
        columnFields_.emplace_back(std::move(name));
        columnFields_.back().setAxis("axisCol");
        columnFields_.back().setFieldIndex(cache_.fieldIndex(columnFields_.back().name()));
        return columnFields_.back();
    }
    PivotField& addPageField(std::string name) {
        pageFields_.emplace_back(std::move(name));
        pageFields_.back().setAxis("axisPage");
        pageFields_.back().setFieldIndex(cache_.fieldIndex(pageFields_.back().name()));
        PivotPageField page;
        page.setFieldIndex(pageFields_.back().fieldIndex());
        pageFieldSettings_.push_back(std::move(page));
        return pageFields_.back();
    }
    PivotFieldReference& addDataField(int fieldIndex = 0) {
        PivotFieldReference reference;
        reference.setFieldIndex(fieldIndex);
        if (fieldIndex >= 0 && static_cast<std::size_t>(fieldIndex) < cache_.fields().size())
            reference.setName(cache_.fields()[static_cast<std::size_t>(fieldIndex)]);
        dataFields_.push_back(std::move(reference));
        return dataFields_.back();
    }
    PivotFieldReference& addDataField(std::string name, std::string subtotal = "sum") {
        PivotFieldReference reference;
        reference.setName(std::move(name));
        reference.setFieldIndex(cache_.fieldIndex(reference.name()));
        reference.setSubtotal(std::move(subtotal));
        dataFields_.push_back(std::move(reference));
        return dataFields_.back();
    }

    const std::vector<PivotField>& rowFields() const noexcept { return rowFields_; }
    const std::vector<PivotField>& columnFields() const noexcept { return columnFields_; }
    const std::vector<PivotField>& pageFields() const noexcept { return pageFields_; }
    const std::vector<PivotFieldReference>& dataFields() const noexcept { return dataFields_; }
    const std::vector<PivotPageField>& pageFieldSettings() const noexcept { return pageFieldSettings_; }

    std::vector<PivotField>& rowFields() noexcept { return rowFields_; }
    std::vector<PivotField>& columnFields() noexcept { return columnFields_; }
    std::vector<PivotField>& pageFields() noexcept { return pageFields_; }
    std::vector<PivotFieldReference>& dataFields() noexcept { return dataFields_; }
    std::vector<PivotPageField>& pageFieldSettings() noexcept { return pageFieldSettings_; }

    const std::vector<PivotFilter>& filters() const noexcept { return filters_; }
    std::vector<PivotFilter>& filters() noexcept { return filters_; }
    PivotFilter& addFilter(PivotFilter filter) {
        if (filter.fieldIndex < 0) throw std::invalid_argument("Pivot filter field index must be non-negative");
        if (filter.type.empty()) throw std::invalid_argument("Pivot filter type cannot be empty");
        if (filter.id == 0) {
            std::uint32_t next = 1;
            for (const auto& existing : filters_) next = std::max(next, existing.id + 1);
            filter.id = next;
        }
        filters_.push_back(std::move(filter));
        return filters_.back();
    }
    void clearFilters() noexcept { filters_.clear(); }

    const std::vector<PivotChartFormat>& chartFormats() const noexcept { return chartFormats_; }
    std::vector<PivotChartFormat>& chartFormats() noexcept { return chartFormats_; }
    PivotChartFormat& addChartFormat(PivotChartFormat format) {
        chartFormats_.push_back(std::move(format));
        return chartFormats_.back();
    }
    void clearChartFormats() noexcept { chartFormats_.clear(); }

private:
    std::string name_;
    std::string location_;
    PivotCache cache_;
    std::vector<PivotField> rowFields_;
    std::vector<PivotField> columnFields_;
    std::vector<PivotField> pageFields_;
    std::vector<PivotFieldReference> dataFields_;
    std::vector<PivotPageField> pageFieldSettings_;
    std::vector<PivotFilter> filters_;
    std::vector<PivotChartFormat> chartFormats_;
    std::optional<std::uint32_t> chartFormatIndex_;
    bool rowGrandTotals_{true};
    bool columnGrandTotals_{true};
    bool compact_{true};
    bool outline_{true};
    bool preserveFormatting_{true};
    bool useAutoFormatting_{true};
    bool showDrill_{true};
    bool multipleFieldFilters_{false};
    std::string styleName_{"PivotStyleLight16"};
    bool showRowHeaders_{true};
    bool showColumnHeaders_{true};
    bool showRowStripes_{false};
    bool showColumnStripes_{false};
};

} // namespace xlpp
