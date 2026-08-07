#include <XLPP/XLPP.h>
#include <XLPP/Cell/RichText.h>
#include "TestOutput.h"
#include <iostream>
void testWriteRichText() {
    xlpp::RichText text; xlpp::RichTextRun a("Bold red "); a.setBold(true); a.setColor("FFFF0000");
    xlpp::RichTextRun b("italic blue"); b.setItalic(true); b.setColor("FF0000FF"); b.setFontName("Arial"); b.setSize(14.0);
    text.addRun(a); text.addRun(b);
    xlpp::Workbook wb; auto& ws=wb.addWorksheet("RichText API");
    ws.cell("A1").setRichText(text);
    ws.columnDimension(1).width = 32.0;
    ws.append({std::string("Run"),std::string("Text"),std::string("Bold"),std::string("Italic"),std::string("Color"),std::string("Font"),std::string("Size")});
    std::size_t index=1; for(const auto& run:text.runs()) ws.append({static_cast<double>(index++),run.text(),run.bold(),run.italic(),run.color(),run.fontName(),run.size()});
    wb.save(xlpp_numbered_tests::outputPath("26_rich_text.xlsx")); std::cout<<"Saved: 26_rich_text.xlsx\n";
}
