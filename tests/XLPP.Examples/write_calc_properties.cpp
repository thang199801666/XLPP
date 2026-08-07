#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
void testWriteCalcProperties() {
    xlpp::Workbook wb; auto& ws=wb.addWorksheet("Calc"); ws.cell("A1").setValue(10.0); ws.cell("A2").setValue(20.0); ws.cell("A3").setFormula("SUM(A1:A2)");
    auto& c=wb.calcProperties(); c.setCalcId(191029); c.setCalcMode("auto"); c.setCalcOnSave(true); c.setFullCalcOnLoad(true); c.setFullPrecision(true); c.setIterate(true); c.setIterateCount(50); c.setIterateDelta(0.0001);
    wb.save(xlpp_numbered_tests::outputPath("30_calc_properties.xlsx")); std::cout<<"Saved: 30_calc_properties.xlsx\n";
}
