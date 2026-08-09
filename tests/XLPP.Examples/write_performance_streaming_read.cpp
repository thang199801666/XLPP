#include "PerformanceCommon.h"
#include <iostream>
void testPerformanceStreamingRead() {
    using namespace xlpp_numbered_tests; const auto dataPath=outputPath("90_performance_streaming_write_data.xlsx"); std::size_t rows=0,cells=0;
    BenchmarkResult r{"streaming_read",0,10,0}; r.seconds=measure([&]{ xlpp::StreamingWorkbookReader reader(dataPath); reader.forEachRow("Data",[&](std::size_t,const xlpp::StreamingRow& row){++rows; cells+=row.size(); return true;}); });
    r.rows=rows; r.cells=cells; r.bytes=std::filesystem::file_size(dataPath); writeBenchmarkWorkbook("91_performance_streaming_read.xlsx",r,250.0,"rows_per_second"); std::cout<<"Saved: 91_performance_streaming_read.xlsx\n";
}
