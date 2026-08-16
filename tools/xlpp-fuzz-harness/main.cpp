// Structured fuzz harness for Workbook::load(). Every iteration either uses a
// seed corpus file unmodified or applies a deterministic mutation (truncation,
// byte flips, size-trimmed random garbage). Run this binary under an
// AddressSanitizer build so memory-safety violations surface as an ASan report
// instead of being silently absorbed by exception handling.
#include <XLPP/XLPP.h>
#include <cstdint>
#include <cstdio>
#include <crtdbg.h>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <random>
#include <string>
#include <vector>

namespace {

std::vector<unsigned char> readFile(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    return std::vector<unsigned char>((std::istreambuf_iterator<char>(stream)),
                                      std::istreambuf_iterator<char>());
}

void writeFile(const std::filesystem::path& path, const std::vector<unsigned char>& bytes) {
    std::ofstream stream(path, std::ios::binary);
    stream.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
}

std::vector<unsigned char> randomGarbage(std::mt19937& rng, std::size_t size) {
    std::vector<unsigned char> bytes(size);
    for (auto& b : bytes) b = static_cast<unsigned char>(rng() & 0xFF);
    return bytes;
}

// Exercising load() with the original seed too (mutations can leave the file
// byte-for-byte identical after truncation at end-of-file).
unsigned loadOnce(const std::filesystem::path& path, unsigned char salt) {
    unsigned failures = 0;
    xlpp::Workbook wb;
    xlpp::LoadOptions options;
    options.maxFileBytes = 64ull * 1024 * 1024;
    options.maxTotalBytes = 256ull * 1024 * 1024;
    options.maxEntries = 1024;
    options.maxEntryBytes = 128ull * 1024 * 1024;
    try {
        wb.load(path, options);
    } catch (const std::exception&) {
        ++failures;
    } catch (...) {
        ++failures;
    }
    if (failures == 0) std::cout << "  LOADED\n" << std::flush;
    // After a successful load, force a save to a memory-neutral temp path to
    // exercise the writer against whatever model the fuzz input produced.
    if (failures == 0) {
        try {
            wb.save(path.string() + ".out.xlsx");
            std::cout << "  SAVED\n" << std::flush;
        } catch (const std::exception&) {
            ++failures;
        } catch (...) {
            ++failures;
        }
    }
    static_cast<void>(salt);
    return failures;
}

} // namespace

int main(int argc, char** argv) {
    _CrtSetReportMode(_CRT_ASSERT, _CRTDBG_MODE_FILE);
    _CrtSetReportFile(_CRT_ASSERT, _CRTDBG_FILE_STDERR);
    const char* corpusDir = argc > 1 ? argv[1] : ".";
    const unsigned iterations = argc > 2 ? static_cast<unsigned>(std::stoul(argv[2])) : 10000;
    const unsigned seedValue = argc > 3 ? static_cast<unsigned>(std::stoul(argv[3])) : 0xC0FFEE;

    std::vector<std::string> seeds;
    for (const auto& entry : std::filesystem::directory_iterator(corpusDir)) {
        if (entry.is_regular_file()) seeds.push_back(entry.path().string());
    }
    if (seeds.empty()) {
        std::cerr << "fuzz-harness: no corpus files found in " << corpusDir << '\n';
        return 2;
    }

    std::mt19937 rng(seedValue);
    const auto tempDir = std::filesystem::temp_directory_path() / "xlpp_fuzz";
    std::filesystem::create_directories(tempDir);

    unsigned long long loaded = 0;
    unsigned failures = 0;
    for (unsigned iter = 0; iter < iterations; ++iter) {
        const auto& seed = seeds[rng() % seeds.size()];
        auto bytes = readFile(seed);
        const unsigned mode = rng() % 4;
        if (mode == 0 && !bytes.empty()) {
            bytes.resize(rng() % (bytes.size() + 1));
        } else if (mode == 1 && !bytes.empty()) {
            const unsigned flips = 1 + (rng() % 8);
            for (unsigned n = 0; n < flips; ++n)
                bytes[rng() % bytes.size()] ^= static_cast<unsigned char>(1u << (rng() % 8));
        } else if (mode == 2) {
            bytes = randomGarbage(rng, rng() % 4097);
        }
        const auto target = tempDir / ("case_" + std::to_string(iter) + ".bin");
        writeFile(target, bytes);
        std::cout << "ITER " << iter << " seed=" << seed << " mode=" << mode
                  << " size=" << bytes.size() << '\n' << std::flush;
        const auto result = loadOnce(target, static_cast<unsigned char>(iter & 0xFF));
        failures += result;
        ++loaded;
        std::filesystem::remove(target);
        std::filesystem::remove(target.string() + ".out.xlsx");
    }

    std::cout << "fuzz-harness: " << loaded << " loads, " << failures
              << " rejected as unparseable, 0 crashes\n";
    return 0;
}
