#include "Packaging/RelationshipGraph.h"
#include "Packaging/ZipArchive.h"
#include <algorithm>
#include <filesystem>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>

namespace {

void printItems(const char* heading, const std::vector<std::string>& items) {
    std::cout << heading << " (" << items.size() << ")\n";
    for (const auto& item : items) std::cout << "  - " << item << '\n';
}

void printValidation(const xlpp::internal::RelationshipValidationReport& report) {
    printItems("Relationship syntax errors", report.relationshipSyntaxErrors);
    printItems("Duplicate relationship IDs", report.duplicateRelationshipIds);
    printItems("Dangling relationships", report.danglingRelationships);
    printItems("Orphaned internal parts", report.orphanedParts);
    printItems("Content-type errors", report.contentTypeErrors);
    printItems("Broken owner references", report.ownerReferenceErrors);
    std::cout << "Validation result: " << (report.ok() ? "PASS" : "FAIL") << "\n";
}

void printInventory(const xlpp::internal::PackageObjectInventory& inventory) {
    std::cout << "Reachable object inventory\n"
              << "  Worksheets: " << inventory.worksheets << '\n'
              << "  Drawings: " << inventory.drawings << '\n'
              << "  Images: " << inventory.images << '\n'
              << "  Charts: " << inventory.charts << '\n'
              << "  Shapes: " << inventory.shapes << '\n'
              << "  Text boxes: " << inventory.textBoxes << '\n'
              << "  Connectors: " << inventory.connectors << '\n'
              << "  Groups: " << inventory.groups << '\n'
              << "  Other drawing objects: " << inventory.otherDrawingObjects << '\n'
              << "  Tables: " << inventory.tables << '\n'
              << "  Comments: " << inventory.comments << '\n'
              << "  External links: " << inventory.externalLinks << '\n'
              << "  Pivot tables: " << inventory.pivotTables << '\n'
              << "  Pivot caches: " << inventory.pivotCaches << '\n';
}

std::string jsonEscape(const std::string& value) {
    std::ostringstream out;
    for (const unsigned char c : value) {
        switch (c) {
        case '"': out << "\\\""; break;
        case '\\': out << "\\\\"; break;
        case '\b': out << "\\b"; break;
        case '\f': out << "\\f"; break;
        case '\n': out << "\\n"; break;
        case '\r': out << "\\r"; break;
        case '\t': out << "\\t"; break;
        default:
            if (c < 0x20) {
                static constexpr char hex[] = "0123456789abcdef";
                out << "\\u00" << hex[(c >> 4) & 0xf] << hex[c & 0xf];
            } else {
                out << static_cast<char>(c);
            }
        }
    }
    return out.str();
}

void jsonStringArray(std::ostream& out, const std::vector<std::string>& values) {
    out << '[';
    for (std::size_t i = 0; i < values.size(); ++i) {
        if (i) out << ',';
        out << '"' << jsonEscape(values[i]) << '"';
    }
    out << ']';
}

void jsonInventory(std::ostream& out, const xlpp::internal::PackageObjectInventory& inventory) {
    out << '{'
        << "\"worksheets\":" << inventory.worksheets << ','
        << "\"drawings\":" << inventory.drawings << ','
        << "\"images\":" << inventory.images << ','
        << "\"charts\":" << inventory.charts << ','
        << "\"shapes\":" << inventory.shapes << ','
        << "\"text_boxes\":" << inventory.textBoxes << ','
        << "\"connectors\":" << inventory.connectors << ','
        << "\"groups\":" << inventory.groups << ','
        << "\"other_drawing_objects\":" << inventory.otherDrawingObjects << ','
        << "\"tables\":" << inventory.tables << ','
        << "\"comments\":" << inventory.comments << ','
        << "\"external_links\":" << inventory.externalLinks << ','
        << "\"pivot_tables\":" << inventory.pivotTables << ','
        << "\"pivot_caches\":" << inventory.pivotCaches
        << '}';
}

void jsonValidation(std::ostream& out, const xlpp::internal::RelationshipValidationReport& report) {
    out << "{\"ok\":" << (report.ok() ? "true" : "false") << ",\"relationship_syntax_errors\":";
    jsonStringArray(out, report.relationshipSyntaxErrors);
    out << ",\"duplicate_relationship_ids\":";
    jsonStringArray(out, report.duplicateRelationshipIds);
    out << ",\"dangling_relationships\":";
    jsonStringArray(out, report.danglingRelationships);
    out << ",\"orphaned_parts\":";
    jsonStringArray(out, report.orphanedParts);
    out << ",\"content_type_errors\":";
    jsonStringArray(out, report.contentTypeErrors);
    out << ",\"owner_reference_errors\":";
    jsonStringArray(out, report.ownerReferenceErrors);
    out << '}';
}

void printUsage(const char* program) {
    std::cerr << "Usage:\n"
              << "  " << program << " <workbook.xlsx> [--json]\n"
              << "  " << program << " <before.xlsx> --compare <after.xlsx> [--json]\n";
}

} // namespace

int main(int argc, char** argv) {
    try {
        bool json = false;
        std::vector<std::string> arguments;
        arguments.reserve(argc > 1 ? static_cast<std::size_t>(argc - 1) : 0);
        for (int i = 1; i < argc; ++i) {
            if (std::string(argv[i]) == "--json") json = true;
            else arguments.emplace_back(argv[i]);
        }
        if (arguments.size() != 1 && arguments.size() != 3) {
            printUsage(argv[0]);
            return 2;
        }

        const std::filesystem::path beforePath = arguments[0];
        const auto before = xlpp::internal::ZipArchive::open(beforePath);
        const auto beforeGraph = xlpp::internal::RelationshipGraph::fromArchive(before);

        if (arguments.size() == 1) {
            const auto validation = beforeGraph.validate();
            if (json) {
                std::cout << "{\"package\":\"" << jsonEscape(beforePath.string()) << "\",\"parts\":"
                          << before.entryNames().size() << ",\"relationships\":" << beforeGraph.relationships().size()
                          << ",\"inventory\":";
                jsonInventory(std::cout, beforeGraph.objectInventory());
                std::cout << ",\"validation\":";
                jsonValidation(std::cout, validation);
                std::cout << "}\n";
            } else {
                std::cout << "Package: " << beforePath.string() << '\n';
                std::cout << "Parts: " << before.entryNames().size() << '\n';
                std::cout << "Relationships: " << beforeGraph.relationships().size() << "\n";
                printInventory(beforeGraph.objectInventory());
                printValidation(validation);
            }
            return validation.ok() ? 0 : 1;
        }

        if (arguments[1] != "--compare") {
            printUsage(argv[0]);
            return 2;
        }

        const std::filesystem::path afterPath = arguments[2];
        const auto after = xlpp::internal::ZipArchive::open(afterPath);
        const auto diff = xlpp::internal::comparePackages(before, after);
        if (json) {
            std::cout << "{\"before\":\"" << jsonEscape(beforePath.string()) << "\",\"after\":\""
                      << jsonEscape(afterPath.string()) << "\",\"added_parts\":";
            jsonStringArray(std::cout, diff.addedParts);
            std::cout << ",\"removed_parts\":";
            jsonStringArray(std::cout, diff.removedParts);
            std::cout << ",\"changed_parts\":";
            jsonStringArray(std::cout, diff.changedParts);
            std::cout << ",\"object_count_regressions\":";
            jsonStringArray(std::cout, diff.objectCountRegressions);
            std::cout << ",\"before_inventory\":";
            jsonInventory(std::cout, diff.beforeObjects);
            std::cout << ",\"after_inventory\":";
            jsonInventory(std::cout, diff.afterObjects);
            std::cout << ",\"before_validation\":";
            jsonValidation(std::cout, diff.beforeValidation);
            std::cout << ",\"after_validation\":";
            jsonValidation(std::cout, diff.afterValidation);
            std::cout << "}\n";
        } else {
            std::cout << "Package: " << beforePath.string() << '\n';
            std::cout << "Parts: " << before.entryNames().size() << '\n';
            std::cout << "Relationships: " << beforeGraph.relationships().size() << "\n";
            printInventory(beforeGraph.objectInventory());
            std::cout << "Compared with: " << afterPath.string() << '\n';
            printItems("Added parts", diff.addedParts);
            printItems("Removed parts", diff.removedParts);
            printItems("Changed parts", diff.changedParts);
            printItems("Object-count regressions", diff.objectCountRegressions);
            std::cout << "\nBefore object inventory\n";
            printInventory(diff.beforeObjects);
            std::cout << "\nAfter object inventory\n";
            printInventory(diff.afterObjects);
            std::cout << "\nBefore package validation\n";
            printValidation(diff.beforeValidation);
            std::cout << "\nAfter package validation\n";
            printValidation(diff.afterValidation);
        }
        return diff.afterValidation.ok() && diff.objectCountRegressions.empty() ? 0 : 1;
    } catch (const std::exception& error) {
        if (argc > 1 && std::find_if(argv + 1, argv + argc, [](const char* value) {
                return std::string(value) == "--json";
            }) != argv + argc) {
            std::cout << "{\"error\":\"" << jsonEscape(error.what()) << "\"}\n";
        } else {
            std::cerr << "xlpp-package-validator: " << error.what() << '\n';
        }
        return 2;
    }
}
