#include "xlpp_capi.h"
#include <cassert>
#include <cstring>

int main() {
    xlpp_clear_error();
    assert(std::strcmp(xlpp_last_error(), "") == 0);

    assert(xlpp_chart_create(99) == nullptr);
    assert(std::strcmp(xlpp_last_error(), "Chart type is invalid") == 0);

    assert(xlpp_workbook_sheet_count(nullptr) == 0);
    assert(std::strcmp(xlpp_last_error(), "Workbook handle is null") == 0);

    auto workbook = xlpp_workbook_create();
    assert(workbook != nullptr);
    assert(xlpp_workbook_add_sheet(workbook, nullptr) == nullptr);
    assert(std::strcmp(xlpp_last_error(), "Workbook and sheet name are required") == 0);

    xlpp_clear_error();
    auto sheet = xlpp_workbook_add_sheet(workbook, "Smoke");
    assert(sheet != nullptr);
    assert(std::strcmp(xlpp_last_error(), "") == 0);

    auto chart = xlpp_chart_create(0);
    assert(chart != nullptr);
    xlpp_chart_set_grouping(chart, 99);
    assert(std::strcmp(xlpp_last_error(), "Chart grouping is invalid") == 0);
    xlpp_chart_set_size(chart, 0, 100);
    assert(std::strcmp(xlpp_last_error(), "Chart size must be positive") == 0);
    assert(xlpp_sheet_add_chart(sheet, chart) == 1);

    xlpp_workbook_destroy(workbook);
    return 0;
}
