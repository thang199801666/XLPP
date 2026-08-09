#pragma once

#include <XLPP/Cell/Cell.h>

#include <cstddef>
#include <vector>

namespace xlpp {

class Worksheet;

class Row {
public:
    Row(Worksheet& sheet, std::size_t rowNumber);
    Cell& cell(std::size_t column);
    const Cell* tryCell(std::size_t column) const noexcept;
    std::size_t number() const noexcept { return rowNumber_; }
    std::vector<Cell*> cells();
    std::vector<CellValue> values() const;

private:
    Worksheet* sheet_;
    std::size_t rowNumber_;
};

} // namespace xlpp
