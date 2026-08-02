#pragma once
#include <XLPP/Cell/Cell.h>
#include <cstddef>
#include <functional>
#include <string>
#include <vector>

namespace xlpp {
class Worksheet;

class CellRange {
public:
    CellRange(Worksheet& worksheet, std::size_t minRow, std::size_t minColumn,
              std::size_t maxRow, std::size_t maxColumn);

    std::size_t minRow() const noexcept { return minRow_; }
    std::size_t minColumn() const noexcept { return minColumn_; }
    std::size_t maxRow() const noexcept { return maxRow_; }
    std::size_t maxColumn() const noexcept { return maxColumn_; }
    std::size_t rowCount() const noexcept { return maxRow_ - minRow_ + 1; }
    std::size_t columnCount() const noexcept { return maxColumn_ - minColumn_ + 1; }
    std::string address() const;

    Cell& cell(std::size_t relativeRow, std::size_t relativeColumn);
    std::vector<Cell*> cells();
    std::vector<std::vector<Cell*>> rows();
    void setValue(const CellValue& value);
    void clear();
    void forEach(const std::function<void(Cell&)>& callback);
    std::vector<CellValue> values() const;
    std::vector<std::string> formulas() const;

private:
    Worksheet* worksheet_{};
    std::size_t minRow_{};
    std::size_t minColumn_{};
    std::size_t maxRow_{};
    std::size_t maxColumn_{};
};
}
