#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteChart();

int main(int argc, char** argv) {
    try {
        xlpp_numbered_tests::setOutputDirectory(
            argc > 1 ? std::filesystem::path(argv[1]) : std::filesystem::current_path());
        testWriteChart();
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
