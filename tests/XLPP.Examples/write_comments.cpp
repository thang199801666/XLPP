#include <XLPP/XLPP.h>
#include "TestOutput.h"
#include <iostream>
void testWriteComments() {
    xlpp::Workbook wb; auto& ws=wb.addWorksheet("Comments");
    ws.cell("A1").setValue("Review"); ws.cell("A1").setComment(xlpp::Comment("Review this value", "Alice"));
    ws.cell("C4").setValue(42.0); ws.cell("C4").setComment(xlpp::Comment("Second note <XML> & spaces", "Bob"));
    wb.save(xlpp_numbered_tests::outputPath("27_comments.xlsx")); std::cout<<"Saved: 27_comments.xlsx\n";
}
