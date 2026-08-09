#include <XLPP/XLPP.h>

#include <filesystem>
#include <iostream>
#include <string>

int main(int argc, char** argv) {
    if (argc != 4) {
        std::cerr << "usage: xlpp-corpus-exercise <input.xlsx> <output.xlsx> <noop|edit-a1|add-sheet|copy-sheet>\n";
        return 2;
    }
    try {
        const std::filesystem::path input = argv[1];
        const std::filesystem::path output = argv[2];
        const std::string action = argv[3];
        xlpp::Workbook workbook;
        workbook.load(input);
        if (action == "edit-a1") {
            if (workbook.sheetCount() == 0) workbook.addWorksheet("Sheet1");
            workbook[0].cell("A1").setValue("XLPP corpus probe");
        } else if (action == "add-sheet") {
            std::string name = "XLPP_Probe";
            for (std::size_t suffix = 2; workbook.worksheet(name) != nullptr; ++suffix) name = "XLPP_Probe_" + std::to_string(suffix);
            workbook.addWorksheet(name).cell("A1").setValue("added by enterprise corpus probe");
        } else if (action == "copy-sheet") {
            if (workbook.sheetCount() == 0) workbook.addWorksheet("Sheet1");
            std::string name = "XLPP_Copy";
            for (std::size_t suffix = 2; workbook.worksheet(name) != nullptr; ++suffix) name = "XLPP_Copy_" + std::to_string(suffix);
            workbook.copyWorksheet(workbook[0], name);
        } else if (action != "noop") {
            std::cerr << "unknown action: " << action << '\n';
            return 2;
        }
        const auto validation = workbook.validate();
        if (!validation.ok()) {
            std::cerr << "pre-save validation failed with " << validation.errorCount << " errors\n";
            return 3;
        }
        workbook.save(output);
        xlpp::Workbook reopened;
        reopened.load(output);
        const auto post = reopened.validate();
        if (!post.ok()) {
            std::cerr << "post-save validation failed with " << post.errorCount << " errors\n";
            return 4;
        }
        std::cout << "PASS " << action << " sheets=" << reopened.sheetCount() << '\n';
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
