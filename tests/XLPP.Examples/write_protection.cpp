#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>

void testWriteProtection() {
    xlpp::Workbook workbook;
    auto& protectedSheet = workbook.addWorksheet("WithPassword");
    protectedSheet.cell("A1").setValue("Protected with password: secret");
    protectedSheet.protection().setEnabled(true);
    protectedSheet.protection().setSort(false);
    protectedSheet.protection().setPassword("secret");

    auto& removedSheet = workbook.addWorksheet("PasswordRemoved");
    removedSheet.cell("A1").setValue("Password was added, then removed before save");
    removedSheet.protection().setEnabled(true);
    removedSheet.protection().setPassword("temporary");
    removedSheet.protection().clearPassword();

    workbook.protection().setLockStructure(true);
    workbook.protection().setPassword("structure");
    workbook.save(xlpp_numbered_tests::outputPath("18_protection.xlsx"));
    std::cout << "Saved: 18_protection.xlsx (sheet password: secret; workbook password: structure)\n";
}
