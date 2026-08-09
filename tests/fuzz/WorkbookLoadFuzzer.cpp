#include <XLPP/XLPP.h>

#include <cstddef>
#include <cstdint>
#include <sstream>
#include <string>

extern "C" int LLVMFuzzerTestOneInput(const std::uint8_t* data, std::size_t size) {
    // Keep individual cases bounded so malformed compression metadata cannot
    // turn a fuzz iteration into an unbounded memory/time sink.
    constexpr std::size_t MaxInputBytes = 8u * 1024u * 1024u;
    if (size == 0 || size > MaxInputBytes) return 0;

    const std::string bytes(reinterpret_cast<const char*>(data), size);
    std::istringstream input(bytes, std::ios::binary);

    xlpp::Workbook workbook;
    xlpp::LoadOptions load;
    load.lenient = true;
    load.maxEntries = 4096;
    load.maxEntryBytes = 32u * 1024u * 1024u;
    load.maxTotalBytes = 128u * 1024u * 1024u;
    load.maxFileBytes = MaxInputBytes;

    try {
        workbook.load(input, load);

        // Exercise semantic traversals on any package that makes it through
        // parsing, then exercise the writer as a second-stage fuzz oracle.
        (void)workbook.validate();
        std::ostringstream output(std::ios::binary);
        xlpp::SaveOptions save;
        save.validateBeforeSave = true;
        workbook.save(output, save);
    } catch (const std::exception&) {
        // Invalid OOXML/ZIP input is an expected result. Sanitizers/libFuzzer
        // still surface memory safety failures, UB, aborts and hangs.
    }
    return 0;
}
