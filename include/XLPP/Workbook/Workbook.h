#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Styles/NamedStyle.h>
#include <XLPP/Workbook/DefinedNames.h>
#include <XLPP/Workbook/DocumentProperties.h>
#include <XLPP/Workbook/Protection.h>
#include <XLPP/Workbook/CalcProperties.h>
#include <XLPP/Workbook/CustomProperties.h>
#include <XLPP/Compression.h>
#include <filesystem>
#include <functional>
#include <istream>
#include <ostream>
#include <string>
#include <vector>

namespace xlpp {

// A package part that XLPP does not model (custom XML, charts, VBA, ...).
// `load` captures such parts verbatim so a subsequent `save` round-trips them
// instead of silently dropping them.
struct PreservedPart {
    std::string name;          // package path, e.g. "customXml/item1.xml"
    std::string data;          // raw part bytes
    std::string overrideType;  // content type to emit as <Override> ("" if covered by a <Default>)
    std::string extension;     // file extension ("" if none)
    std::string defaultType;   // content type of the <Default> rule ("" if none)
    bool compress{true};       // whether to compress this part on save
};

// Issues collected while loading. With `LoadOptions::lenient` the load continues
// past recoverable (sheet-level) failures instead of aborting.
struct LoadDiagnostics {
    std::vector<std::string> warnings;
    std::vector<std::string> errors;
    bool hadErrors() const noexcept { return !errors.empty(); }
};

// Guards applied while opening a package for load; 0 means unlimited.
struct LoadOptions {
    bool lenient{false};        // continue past malformed sheets when possible
    std::size_t maxEntries{0};
    std::size_t maxEntryBytes{0};
    std::size_t maxTotalBytes{0};
    std::size_t maxFileBytes{0};
    std::function<bool()> cancel{};
    std::function<void(std::size_t done, std::size_t total)> progress{};
};

class Workbook {
public:
    Worksheet& addWorksheet(std::string name);
    Worksheet& copyWorksheet(const Worksheet& source, std::string newName);
    bool removeWorksheet(const std::string& name);
    Worksheet* worksheet(const std::string& name) noexcept;
    const Worksheet* worksheet(const std::string& name) const noexcept;
    Worksheet& operator[](std::size_t index);
    const Worksheet& operator[](std::size_t index) const;
    std::size_t index(const Worksheet& sheet) const;
    std::vector<std::string> sheetNames() const;
    std::size_t sheetCount() const noexcept { return sheets_.size(); }
    std::vector<Worksheet>& worksheets() noexcept { return sheets_; }
    const std::vector<Worksheet>& worksheets() const noexcept { return sheets_; }
    NamedStyle& addNamedStyle(NamedStyle style);
    NamedStyle* namedStyle(const std::string& name) noexcept;
    const NamedStyle* namedStyle(const std::string& name) const noexcept;
    const std::vector<NamedStyle>& namedStyles() const noexcept { return namedStyles_; }
    void applyNamedStyle(Cell& cell, const std::string& name) const;
    DefinedName& addDefinedName(DefinedName name);
    DefinedName* definedName(const std::string& name) noexcept;
    const DefinedName* definedName(const std::string& name) const noexcept;
    std::vector<DefinedName>& definedNames() noexcept { return definedNames_; }
    const std::vector<DefinedName>& definedNames() const noexcept { return definedNames_; }
    DocumentProperties& properties() noexcept { return properties_; }
    const DocumentProperties& properties() const noexcept { return properties_; }
    WorkbookProtection& protection() noexcept { return protection_; }
    const WorkbookProtection& protection() const noexcept { return protection_; }
    CalcProperties& calcProperties() noexcept { return calcProps_; }
    const CalcProperties& calcProperties() const noexcept { return calcProps_; }
    CustomProperties& customProperties() noexcept { return customProps_; }
    const CustomProperties& customProperties() const noexcept { return customProps_; }

    // Date system: false uses the 1900 epoch (Excel's default, including the
    // phantom 1900-02-29), true uses the 1904 epoch.
    void setDate1904(bool enabled) noexcept { date1904_ = enabled; }
    bool date1904() const noexcept { return date1904_; }

    void clear() { sheets_.clear(); namedStyles_.clear(); definedNames_.clear(); properties_ = {}; protection_ = {}; date1904_ = false; preservedParts_.clear(); strictNamespaces_ = false; diagnostics_ = {}; calcProps_ = {}; customProps_ = {}; }
    void load(const std::filesystem::path& path);
    void load(const std::filesystem::path& path, const LoadOptions& options);
    void load(std::istream& stream);
    void load(std::istream& stream, const LoadOptions& options);
    void save(const std::filesystem::path& path) const;
    void save(const std::filesystem::path& path, const SaveOptions& options) const;
    void save(std::ostream& stream) const;
    void save(std::ostream& stream, const SaveOptions& options) const;
    const std::vector<PreservedPart>& preservedParts() const noexcept { return preservedParts_; }
    std::vector<PreservedPart>& preservedParts() noexcept { return preservedParts_; }
    // True when the last load read a package using strict OOXML namespaces.
    bool strictNamespaces() const noexcept { return strictNamespaces_; }
    // Diagnostics from the most recent load.
    const LoadDiagnostics& diagnostics() const noexcept { return diagnostics_; }
private:
    std::vector<Worksheet> sheets_;
    std::vector<NamedStyle> namedStyles_;
    std::vector<DefinedName> definedNames_;
    DocumentProperties properties_;
    WorkbookProtection protection_;
    CalcProperties calcProps_;
    CustomProperties customProps_;
    std::vector<PreservedPart> preservedParts_;
    LoadDiagnostics diagnostics_;
    bool date1904_{false};
    bool strictNamespaces_{false};
};
}
