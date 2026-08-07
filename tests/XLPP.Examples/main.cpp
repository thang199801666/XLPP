#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteBasic(); void testWriteRows(); void testWriteFormula(); void testWriteStyle(); void testWriteTypes();
void testWriteSheets(); void testWriteLayout(); void testWriteTable(); void testWriteFilter(); void testWriteLinks();
void testWriteValidation(); void testWriteConditional(); void testWriteProperties(); void testWriteRoundtrip(); void testWriteRange();
void testWriteNamedStyle(); void testWritePageSetup(); void testWriteProtection(); void testWriteStreaming(); void testWriteCustomProperties();
void testWriteDateSystem(); void testWriteArrayFormula(); void testWriteCopySheet(); void testWriteLoadOptions(); void testWriteChart();
void testWriteRichText(); void testWriteComments(); void testWriteSheetView(); void testWriteDefinedNamesFull(); void testWriteCalcProperties();
void testWriteDocumentPropertiesFull(); void testWriteMultiChart(); void testWritePivot(); void testWriteFormulaMetadata(); void testWriteCompressionModes();
void testWriteImage(const std::filesystem::path&);
void testWriteVbaText();
void testPerformanceStreamingWrite(); void testPerformanceStreamingRead(); void testPerformanceDomRoundTrip(); void testPerformanceStyled();

int main(int argc, char** argv) {
    try {
        const std::filesystem::path output = argc > 1 ? std::filesystem::path(argv[1]) : std::filesystem::current_path()/"XLPP.TestOutputs";
        xlpp_numbered_tests::setOutputDirectory(output);
        testWriteBasic(); testWriteRows(); testWriteFormula(); testWriteStyle(); testWriteTypes(); testWriteSheets(); testWriteLayout();
        testWriteTable(); testWriteFilter(); testWriteLinks(); testWriteValidation(); testWriteConditional(); testWriteProperties();
        testWriteRoundtrip(); testWriteRange(); testWriteNamedStyle(); testWritePageSetup(); testWriteProtection(); testWriteStreaming();
        testWriteCustomProperties(); testWriteDateSystem(); testWriteArrayFormula(); testWriteCopySheet(); testWriteLoadOptions(); testWriteChart();
        testWriteRichText(); testWriteComments(); testWriteSheetView(); testWriteDefinedNamesFull(); testWriteCalcProperties();
        testWriteDocumentPropertiesFull(); testWriteMultiChart(); testWritePivot(); testWriteFormulaMetadata(); testWriteCompressionModes();
        const std::filesystem::path defaultImage = R"(D:\icons8-front-view-64.png)";
        const std::filesystem::path imagePath = argc > 2 ? std::filesystem::path(argv[2]) : defaultImage;
        if (std::filesystem::exists(imagePath)) testWriteImage(imagePath);
        else std::cout << "Skipping 36_image.xlsx: image not found at " << imagePath << '\n';
        testWriteVbaText();
        testPerformanceStreamingWrite(); testPerformanceStreamingRead(); testPerformanceDomRoundTrip(); testPerformanceStyled();
        std::cout << "All numbered feature and performance workbooks were generated in: " << output << '\n';
        return 0;
    } catch (const std::exception& e) { std::cerr << "[NUMBERED TEST FAIL] " << e.what() << '\n'; return 1; }
}
