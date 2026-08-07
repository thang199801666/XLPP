#include "TestOutput.h"
#include <filesystem>
#include <iostream>

void testWriteImage(const std::filesystem::path& imagePath);

int main(int argc, char** argv) {
    if (argc < 2) {
        std::cerr << "Usage: XLPP_WriteImage <image-path> [output-directory]\n";
        return 2;
    }
    try {
        xlpp_numbered_tests::setOutputDirectory(
            argc > 2 ? std::filesystem::path(argv[2]) : std::filesystem::current_path());
        testWriteImage(std::filesystem::path(argv[1]));
        return 0;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 1;
    }
}
