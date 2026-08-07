#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Styles/NamedStyle.h>
#include <XLPP/Workbook/DefinedNames.h>
#include <XLPP/Workbook/DocumentProperties.h>
#include <XLPP/Workbook/Protection.h>
#include <XLPP/Workbook/CalcProperties.h>
#include <XLPP/Workbook/CustomProperties.h>
#include <XLPP/Vba/VbaModule.h>
#include <XLPP/Compression.h>
#include <deque>
#include <filesystem>
#include <functional>
#include <istream>
#include <ostream>
#include <optional>
#include <string>
#include <vector>

namespace xlpp {

// A package part that XLPP does not model (custom XML, charts, VBA, ...).
// `load` captures such parts verbatim so a subsequent `save` round-trips them
// instead of silently dropping them.

// A relationship captured from an existing OPC package. XLPP keeps these
// separately from raw parts so regenerated workbook/worksheet .rels files can
// merge unsupported relationships instead of silently disconnecting objects.
struct PreservedRelationship {
    std::string sourcePart;   // empty means package root (_rels/.rels)
    std::string id;
    std::string type;
    std::string target;
    std::string targetMode;   // empty/Internal or External
};

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

struct ChartCacheSyncOptions {
    bool synchronizeTitles{true};
    bool synchronizeCategories{true};
    bool synchronizeValues{true};
    // Unsupported/external/2-D references are preserved by default. If true,
    // their existing cache is cleared so the host is forced to recalculate it.
    bool clearUnsupportedReferences{false};
};

struct ChartCacheSyncReport {
    std::size_t chartsVisited{0};
    std::size_t seriesVisited{0};
    std::size_t cachesUpdated{0};
    std::size_t cachesCleared{0};
    std::size_t referencesSkipped{0};
    std::vector<std::string> warnings;
    bool success() const noexcept { return warnings.empty(); }
};

class Workbook {
public:
    // Stable worksheet storage: std::deque never invalidates references on
    // insertion at either end, so `Worksheet&` / `Worksheet*` obtained from
    // addWorksheet/copyWorksheet/operator[] remain valid across further
    // worksheet insertion. Only removal of a sheet invalidates references to
    // that sheet.
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
    std::deque<Worksheet>& worksheets() noexcept { return sheets_; }
    const std::deque<Worksheet>& worksheets() const noexcept { return sheets_; }
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

    void clear() { sheets_.clear(); namedStyles_.clear(); definedNames_.clear(); properties_ = {}; protection_ = {}; date1904_ = false; preservedParts_.clear(); preservedRelationships_.clear(); sourceWorkbookXml_.clear(); sourceSheetParts_.clear(); sourceSheetXml_.clear(); sourceSheetNames_.clear(); cachedSheetXml_.clear(); cachedSheetXmlStrict_ = false; cachedSheetXmlDate1904_ = false; generatedVbaProject_ = false; strictNamespaces_ = false; diagnostics_ = {}; calcProps_ = {}; customProps_ = {}; }
    void load(const std::filesystem::path& path);
    void load(const std::filesystem::path& path, const LoadOptions& options);
    void load(std::istream& stream);
    void load(std::istream& stream, const LoadOptions& options);
    void save(const std::filesystem::path& path) const;
    void save(const std::filesystem::path& path, const SaveOptions& options) const;
    void save(std::ostream& stream) const;
    void save(std::ostream& stream, const SaveOptions& options) const;

    // Rebuild chart title/category/value caches from their worksheet A1
    // references. Supported references are single-cell or one-dimensional
    // ranges, optionally qualified with a local worksheet name. External
    // workbooks, structured references and unions are intentionally skipped.
    ChartCacheSyncReport synchronizeChartCaches(const ChartCacheSyncOptions& options = {});
    // Attach an existing vbaProject.bin and save the workbook as .xlsm.
    // XL++ preserves externally supplied project bytes verbatim.
    void addVbaProject(const std::filesystem::path& path);
    void setVbaProject(std::vector<unsigned char> bytes);
    bool hasVbaProject() const noexcept;
    bool removeVbaProject() noexcept;

    // Create or replace a standard VBA module directly from source text. XL++
    // normalizes line endings and builds the CFB/OVBA project streams required
    // by vbaProject.bin. It packages source; it does not validate VBA syntax or
    // compile forms, references, signatures, or designer state.
    void setVbaModuleText(std::string moduleName, std::string source);
    std::optional<std::string> vbaModuleText(const std::string& moduleName) const;
    std::vector<VbaModule> vbaModules() const;
    bool removeVbaModule(const std::string& moduleName);

    const std::vector<PreservedPart>& preservedParts() const noexcept { return preservedParts_; }
    std::vector<PreservedPart>& preservedParts() noexcept { return preservedParts_; }
    const std::vector<PreservedRelationship>& preservedRelationships() const noexcept { return preservedRelationships_; }
    // True when the last load read a package using strict OOXML namespaces.
    bool strictNamespaces() const noexcept { return strictNamespaces_; }
    // Diagnostics from the most recent load.
    const LoadDiagnostics& diagnostics() const noexcept { return diagnostics_; }
private:
    std::deque<Worksheet> sheets_;
    std::vector<NamedStyle> namedStyles_;
    std::vector<DefinedName> definedNames_;
    DocumentProperties properties_;
    WorkbookProtection protection_;
    CalcProperties calcProps_;
    CustomProperties customProps_;
    std::vector<PreservedPart> preservedParts_;
    std::vector<PreservedRelationship> preservedRelationships_;
    std::string sourceWorkbookXml_;
    std::vector<std::string> sourceSheetParts_;
    std::vector<std::string> sourceSheetXml_;
    std::vector<std::string> sourceSheetNames_;
    LoadDiagnostics diagnostics_;
    bool date1904_{false};
    bool strictNamespaces_{false};
    // True when vbaProject.bin was generated from source text by XL++. Such
    // projects are rebuilt at save time so document modules stay synchronized
    // with worksheets added or removed after setVbaModuleText().
    bool generatedVbaProject_{false};
    // Differential-save cache: serialized sheet XML reused when a sheet is clean.
    // Keyed by the serialization-affecting options so a later save with different
    // strict/date1904 settings re-serializes instead of reusing stale XML.
    mutable std::vector<std::string> cachedSheetXml_;
    mutable bool cachedSheetXmlStrict_{false};
    mutable bool cachedSheetXmlDate1904_{false};
};
}
