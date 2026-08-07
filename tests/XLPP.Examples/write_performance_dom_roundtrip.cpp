#include "PerformanceCommon.h"
#include <iostream>
void testPerformanceDomRoundTrip() {
    using namespace xlpp_numbered_tests; const std::size_t rows=5000,cols=12; const auto dataPath=outputPath("92_performance_dom_roundtrip_data.xlsx"); xlpp::Workbook wb; auto& ws=wb.addWorksheet("DOM"); for(std::size_t r=1;r<=rows;++r) for(std::size_t c=1;c<=cols;++c) ws.cell(r,c).setValue(c%3?xlpp::CellValue(static_cast<double>(r*c)):xlpp::CellValue(std::string("value-")+std::to_string(r%250)));
    BenchmarkResult result{"dom_save_load",rows,cols,rows*cols}; result.seconds=measure([&]{wb.save(dataPath); xlpp::Workbook loaded; loaded.load(dataPath); if(loaded.sheetCount()!=1) throw std::runtime_error("DOM load mismatch");}); result.bytes=std::filesystem::file_size(dataPath); writeBenchmarkWorkbook("92_performance_dom_roundtrip.xlsx",result,100.0,"cells_per_second"); std::cout<<"Saved: 92_performance_dom_roundtrip.xlsx\n";
}
