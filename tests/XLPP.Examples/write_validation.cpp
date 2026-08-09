#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteValidation() {
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Validation");
    sheet.cell("A1").setValue("Whole number (1-100)");
    sheet.cell("C1").setValue("Status list");
    sheet.columnDimension("A").width = 24.0;
    sheet.columnDimension("C").width = 18.0;

    auto& whole = sheet.dataValidations().add(xlpp::DataValidationType::Whole, "A2:A10");
    whole.setOperator(xlpp::DataValidationOperator::Between);
    whole.setFormula1("1");
    whole.setFormula2("100");
    whole.setAllowBlank(true);
    whole.setShowInputMessage(true);
    whole.setPromptTitle("Allowed values");
    whole.setPrompt("Enter a whole number from 1 through 100.");
    whole.setShowErrorMessage(true);
    whole.setErrorTitle("Invalid number");
    whole.setError("The value must be a whole number between 1 and 100.");

    auto& list = sheet.dataValidations().add(xlpp::DataValidationType::List, "C2:C10");
    list.setFormula1("\"Open,In progress,Closed\"");
    list.setAllowBlank(true);
    list.setShowInputMessage(true);
    list.setPromptTitle("Select status");
    list.setPrompt("Choose a value from the drop-down list.");
    list.setShowErrorMessage(true);
    list.setErrorTitle("Invalid status");
    list.setError("Choose Open, In progress, or Closed.");

    workbook.save(xlpp_numbered_tests::outputPath("11_validation.xlsx"));
    std::cout << "Saved: 11_validation.xlsx\n";
}
