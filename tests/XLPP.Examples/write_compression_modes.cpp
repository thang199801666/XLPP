#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
void testWriteCompressionModes() {
    xlpp::Workbook wb; auto& ws=wb.addWorksheet("Compression"); for(int r=1;r<=1000;++r) ws.append({static_cast<double>(r),std::string("repeated-value"),std::string("group-")+std::to_string(r%10)});
    xlpp::SaveOptions options; options.compressionLevel=xlpp::CompressionLevel::Best; options.parallelWorkers=2; wb.save(xlpp_numbered_tests::outputPath("35_compression_modes.xlsx"), options); std::cout<<"Saved: 35_compression_modes.xlsx\n";
}
