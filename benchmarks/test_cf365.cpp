// Test Excel 365 conditional formatting: data bars, color scales, icon sets
#include <XLPP/XLPP.h>
#include <iostream>
#include <filesystem>

using namespace xlpp;

int main() {
    auto tmp = std::filesystem::temp_directory_path() / "xlpp_cf365.xlsx";

    {
        Workbook wb;
        auto& ws = wb.addWorksheet("CF");

        for (int r = 1; r <= 10; ++r)
            ws.cell(r, 1).setValue(static_cast<double>(r * 10));

        // Data bar
        auto db = ConditionalRule::dataBar("FF00B050");
        db.setPriority(1);
        db.getDataBar().min = Cfvo("min", 0.0);
        db.getDataBar().max = Cfvo("max", 0.0);
        db.getDataBar().min.type = "min";
        db.getDataBar().max.type = "max";
        ws.conditionalFormatting().addRule("A1:A10", std::move(db));

        // Color scale (3 stops)
        auto cs = ConditionalRule::colorScale();
        cs.setPriority(2);
        Cfvo stop1; stop1.type = "min"; stop1.color = "FFF8696B";
        Cfvo stop2; stop2.type = "percent"; stop2.value = 50; stop2.color = "FFFFEB84";
        Cfvo stop3; stop3.type = "max"; stop3.color = "FF63BE7B";
        cs.getColorScale().addStop(stop1);
        cs.getColorScale().addStop(stop2);
        cs.getColorScale().addStop(stop3);
        ws.conditionalFormatting().addRule("B1:B10", std::move(cs));

        // Icon set (3 arrows)
        auto is = ConditionalRule::iconSet("3Arrows");
        is.setPriority(3);
        Cfvo t1; t1.type = "percent"; t1.value = 0;
        Cfvo t2; t2.type = "percent"; t2.value = 33;
        Cfvo t3; t3.type = "percent"; t3.value = 67;
        is.getIconSet().addThreshold(t1);
        is.getIconSet().addThreshold(t2);
        is.getIconSet().addThreshold(t3);
        ws.conditionalFormatting().addRule("C1:C10", std::move(is));

        wb.save(tmp);
        std::cout << "Saved CF365 workbook\n";
    }

    {
        Workbook wb;
        wb.load(tmp);
        auto& ws = wb.worksheet("CF")->conditionalFormatting();
        std::cout << "Rules loaded: " << ws.entries().size() << " entries\n";
        for (const auto& entry : ws.entries()) {
            for (const auto& rule : entry.rules()) {
                switch (rule.type()) {
                    case ConditionalRuleType::DataBar:
                        std::cout << "  DataBar: color=" << rule.getDataBar().color
                                  << " minType=" << rule.getDataBar().min.type
                                  << " maxType=" << rule.getDataBar().max.type << "\n";
                        break;
                    case ConditionalRuleType::ColorScale:
                        std::cout << "  ColorScale: " << rule.getColorScale().stops.size() << " stops, "
                                  << "c1=" << (rule.getColorScale().stops.size() > 0 ? *rule.getColorScale().stops[0].color : "")
                                  << "\n";
                        break;
                    case ConditionalRuleType::IconSet:
                        std::cout << "  IconSet: " << rule.getIconSet().icons
                                  << " thresholds=" << rule.getIconSet().thresholds.size() << "\n";
                        break;
                    default:
                        std::cout << "  Other rule\n";
                }
            }
        }
    }
    std::filesystem::remove(tmp);
    return 0;
}
