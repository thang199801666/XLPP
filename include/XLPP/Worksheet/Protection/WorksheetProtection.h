#pragma once
#include <string>

namespace xlpp {
class WorksheetProtection {
public:
    bool enabled() const noexcept { return enabled_; }
    void setEnabled(bool v) noexcept { enabled_ = v; }
    const std::string& passwordHash() const noexcept { return passwordHash_; }
    void setPasswordHash(std::string v) { passwordHash_ = std::move(v); enabled_ = true; }
    bool selectLockedCells() const noexcept { return selectLockedCells_; }
    void setSelectLockedCells(bool v) noexcept { selectLockedCells_ = v; }
    bool selectUnlockedCells() const noexcept { return selectUnlockedCells_; }
    void setSelectUnlockedCells(bool v) noexcept { selectUnlockedCells_ = v; }
    bool formatCells() const noexcept { return formatCells_; }
    void setFormatCells(bool v) noexcept { formatCells_ = v; }
    bool formatColumns() const noexcept { return formatColumns_; }
    void setFormatColumns(bool v) noexcept { formatColumns_ = v; }
    bool formatRows() const noexcept { return formatRows_; }
    void setFormatRows(bool v) noexcept { formatRows_ = v; }
    bool insertRows() const noexcept { return insertRows_; }
    void setInsertRows(bool v) noexcept { insertRows_ = v; }
    bool insertColumns() const noexcept { return insertColumns_; }
    void setInsertColumns(bool v) noexcept { insertColumns_ = v; }
    bool deleteRows() const noexcept { return deleteRows_; }
    void setDeleteRows(bool v) noexcept { deleteRows_ = v; }
    bool deleteColumns() const noexcept { return deleteColumns_; }
    void setDeleteColumns(bool v) noexcept { deleteColumns_ = v; }
    bool sort() const noexcept { return sort_; }
    void setSort(bool v) noexcept { sort_ = v; }
    bool autoFilter() const noexcept { return autoFilter_; }
    void setAutoFilter(bool v) noexcept { autoFilter_ = v; }
private:
    bool enabled_{false};
    std::string passwordHash_;
    bool selectLockedCells_{true}, selectUnlockedCells_{true};
    bool formatCells_{false}, formatColumns_{false}, formatRows_{false};
    bool insertRows_{false}, insertColumns_{false}, deleteRows_{false}, deleteColumns_{false};
    bool sort_{false}, autoFilter_{false};
};
}
