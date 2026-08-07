#include "PerformanceCommon.h"
#include <iostream>
#include <vector>
void testPerformanceStreamingWrite() {
    using namespace xlpp_numbered_tests; const std::size_t rows=50000, cols=10; const auto dataPath=outputPath("90_performance_streaming_write_data.xlsx");
    BenchmarkResult r{"streaming_write",rows,cols,rows*cols};
    r.seconds=measure([&]{ xlpp::StreamingWorkbookWriter writer(dataPath,xlpp::SharedStringMode::Hash,4096); writer.setCompressionLevel(xlpp::CompressionLevel::Fastest); writer.setParallelWorkers(2); auto& ws=writer.addWorksheet("Data"); std::vector<xlpp::CellValue> row; row.reserve(cols); for(std::size_t i=0;i<rows;++i){ row.clear(); for(std::size_t c=0;c<cols;++c) row.emplace_back(c%2?xlpp::CellValue(std::string("group-")+std::to_string(i%100)):xlpp::CellValue(static_cast<double>(i*cols+c))); ws.append(row);} writer.close(); });
    r.bytes=std::filesystem::file_size(dataPath); writeBenchmarkWorkbook("90_performance_streaming_write.xlsx",r,250.0,"rows_per_second"); std::cout<<"Saved: 90_performance_streaming_write.xlsx\n";
}
