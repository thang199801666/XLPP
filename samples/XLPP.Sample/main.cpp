#include <XLPP/XLPP.h>
#include <XLPP/Streaming/StreamingWorkbookWriter.h>
#include <XLPP/Streaming/StreamingWorkbookReader.h>
#include <iostream>
#include <iomanip>

int main() {
    std::cout << "=== XL++ Milestone 21 Demo ===\n\n";

    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Sales");
    sheet.append({std::string("Product"), std::string("Status"), std::string("Amount")});
    sheet.append({std::string("XL-A"), std::string("Open"), 150.0});
    sheet.append({std::string("XL-B"), std::string("Closed"), 80.0});
    sheet.append({std::string("XL-C"), std::string("Pending"), 200.0});

    sheet.autoFilter().setReference("A1:C4");
    sheet.autoFilter().column(1).addValue("Open");
    sheet.autoFilter().column(2).addCustomFilter(xlpp::FilterOperator::GreaterThanOrEqual, "100");

    auto negativeRule = xlpp::ConditionalRule::cellIs(xlpp::ConditionalOperator::LessThan, "100");
    negativeRule.setPriority(1);
    negativeRule.differentialStyle().font().color().setArgb("FF9C0006");
    negativeRule.differentialStyle().fill().foregroundColor().setArgb("FFFFC7CE");
    sheet.conditionalFormatting().addRule("C2:C4", std::move(negativeRule));

    auto statusValidation = xlpp::DataValidation::list("B2:B4", "\"Open,Closed,Pending\"");
    statusValidation.setShowDropDown(true);
    statusValidation.setShowErrorMessage(true);
    statusValidation.setErrorTitle("Invalid status");
    sheet.dataValidations().add(std::move(statusValidation));

    auto& cell = sheet.cell("A1");
    cell.font().setBold(true);
    cell.font().setSize(14);
    cell.fill().setPatternType("solid");
    cell.fill().foregroundColor().setArgb("FF4472C4");
    cell.font().color().setArgb("FFFFFFFF");

    workbook.save("XLPP_DataValidation_Sample.xlsx");
    std::cout << "Saved XLPP_DataValidation_Sample.xlsx with styles, filters, CF, and DV.\n\n";

    workbook.clear();
    workbook.load("XLPP_DataValidation_Sample.xlsx");
    auto* loadedSheet = workbook.worksheet("Sales");
    std::cout << "Reloaded: " << loadedSheet->dimensions() << '\n';
    std::cout << "  Row count: " << loadedSheet->rowCount() << '\n';
    std::cout << "  A1: " << (loadedSheet->cell("A1").isString() ? std::get<std::string>(loadedSheet->cell("A1").value()) : std::string{"-"}) << '\n';
    std::cout << "  A1 font bold: " << (loadedSheet->cell("A1").font().bold() ? "yes" : "no") << '\n';
    std::cout << "  Data validations: " << loadedSheet->dataValidations().items().size() << '\n';
    std::cout << "  Conditional formatting entries: " << loadedSheet->conditionalFormatting().entries().size() << '\n';
    std::cout << "  AutoFilter enabled: " << (loadedSheet->autoFilter().enabled() ? "yes" : "no") << '\n';

    // Streaming writer + reader
    const std::string streamingPath = "XLPP_Streaming_Sample.xlsx";
    {
        xlpp::StreamingWorkbookWriter writer(streamingPath);
        auto& logSheet = writer.addWorksheet("Log");
        for (int row = 1; row <= 10000; ++row)
            logSheet.append({std::string("event-") + std::to_string(row), static_cast<double>(row), row % 2 == 0});
        writer.close();
    }
    {
        xlpp::StreamingWorkbookReader reader(streamingPath);
        auto logSheet = reader.worksheet("Log");
        std::size_t rows = 0;
        double total = 0.0;
        for (auto it = logSheet.begin(); it != logSheet.end(); ++it) {
            ++rows;
            for (const auto& sc : *it)
                if (std::holds_alternative<double>(sc.value)) total += std::get<double>(sc.value);
        }
        std::cout << "\nStreaming: " << rows << " rows, numeric total = " << std::fixed << std::setprecision(0) << total << ".\n";
    }

    return 0;
}
