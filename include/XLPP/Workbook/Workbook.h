#pragma once
#include <XLPP/Core/StableVector.h>
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Styles/NamedStyle.h>
#include <XLPP/Workbook/DefinedNames.h>
#include <XLPP/Workbook/DocumentProperties.h>
#include <XLPP/Workbook/Protection.h>
#include <XLPP/Workbook/CalcProperties.h>
#include <XLPP/Workbook/CustomProperties.h>
#include <XLPP/Workbook/Validation.h>
#include <XLPP/Workbook/LoadOptions.h>
#include <XLPP/Workbook/ChartCacheSynchronization.h>
#include <XLPP/Vba/VbaModule.h>
#include <XLPP/Compression.h>
#include <XLPP/Package/Preservation.h>
#include <XLPP/Formula/Calculation.h>
#include <XLPP/Formula/DependencyGraph.h>
#include <XLPP/Workbook/StructuralEditing.h>
#include <XLPP/ExternalData/ExternalData.h>
#include <XLPP/DataModel/DataModel.h>
#include <XLPP/Enterprise/EnterpriseFeatures.h>
#include <deque>
#include <filesystem>
#include <istream>
#include <ostream>
#include <optional>
#include <string>
#include <vector>
#include <unordered_map>

namespace xlpp {

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
    const StableVector<NamedStyle>& namedStyles() const noexcept { return namedStyles_; }
    void applyNamedStyle(Cell& cell, const std::string& name) const;
    DefinedName& addDefinedName(DefinedName name);
    // Unscoped lookup follows Excel resolution order: workbook-scoped name
    // first, then the first matching local name. Name comparison is ASCII
    // case-insensitive, matching Excel's behavior for normal identifiers.
    DefinedName* definedName(const std::string& name) noexcept;
    const DefinedName* definedName(const std::string& name) const noexcept;
    DefinedName* definedName(const std::string& name, std::optional<std::size_t> localSheetId) noexcept;
    const DefinedName* definedName(const std::string& name, std::optional<std::size_t> localSheetId) const noexcept;
    StableVector<DefinedName>& definedNames() noexcept { return definedNames_; }
    const StableVector<DefinedName>& definedNames() const noexcept { return definedNames_; }
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

    void clear();
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
    // Evaluate supported formulas in-process and refresh their cached values.
    // Unsupported formulas are preserved verbatim and reported rather than
    // rewritten, so Excel/LibreOffice can still calculate them on open.
    CalculationReport calculateFormulas(const CalculationOptions& options = {});

    // Build an immutable workbook-wide dependency graph without calculating or mutating cells.
    FormulaDependencyGraph dependencyGraph() const;

    // Excel-grade structural transactions. Unlike Worksheet-only edits, these
    // rewrite cross-sheet formulas, defined names, charts, pivots and other
    // modeled references throughout the workbook.
    StructuralEditReport applyStructuralEdit(const StructuralEdit& edit, const StructuralEditOptions& options = {});
    StructuralEditReport insertRows(const std::string& sheetName, std::size_t index, std::size_t amount = 1, const StructuralEditOptions& options = {});
    StructuralEditReport deleteRows(const std::string& sheetName, std::size_t index, std::size_t amount = 1, const StructuralEditOptions& options = {});
    StructuralEditReport insertColumns(const std::string& sheetName, std::size_t index, std::size_t amount = 1, const StructuralEditOptions& options = {});
    StructuralEditReport deleteColumns(const std::string& sheetName, std::size_t index, std::size_t amount = 1, const StructuralEditOptions& options = {});

    // Dependency-aware worksheet rename. Prefer this over Worksheet::rename()
    // when the sheet belongs to a workbook: explicit qualifiers in formulas,
    // names, charts, pivots and internal hyperlinks are rewritten in-place.
    WorksheetRenameReport renameWorksheet(const std::string& oldName, const std::string& newName,
                                           const WorksheetRenameOptions& options = {});

    ChartCacheSyncReport synchronizeChartCaches(const ChartCacheSyncOptions& options = {});
    // Clears dependency fingerprints used by changedReferencesOnly mode. The
    // next synchronization treats every supported chart reference as new.
    void resetChartCacheDependencyTracking() noexcept;
    std::size_t trackedChartCacheDependencyCount() const noexcept;
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
    void setVbaClassModuleText(std::string moduleName, std::string source,
                               bool readOnly = false, bool privateModule = false);
    void setVbaDocumentModuleText(std::string moduleName, std::string source);
    void setVbaModule(VbaModule module);
    std::optional<std::string> vbaModuleText(const std::string& moduleName) const;
    std::vector<VbaModule> vbaModules() const;
    bool removeVbaModule(const std::string& moduleName);
    std::vector<unsigned char> vbaProjectBytes() const;
    void saveVbaProject(const std::filesystem::path& path) const;
    bool hasVbaSignature() const noexcept;
    bool vbaSourceEditable() const noexcept { return generatedVbaProject_ && !hasVbaSignature(); }
    VbaProjectProperties vbaProjectProperties() const;
    void setVbaProjectProperties(VbaProjectProperties properties);

    // Preservation-first inspection of external-data and Data Model package
    // topology. These APIs never regenerate proprietary query/model payloads.
    ExternalDataInspection inspectExternalData() const;
    DataModelInspection inspectDataModel() const;
    EnterpriseFeatureInspection inspectEnterpriseFeatures() const;

    // Controlled semantic edits for preserved enterprise XML. These methods
    // patch only the selected attribute and leave opaque sibling payloads and
    // relationships byte-for-byte unchanged.
    EnterpriseEditReport setConnectionRefreshOnLoad(const std::string& connectionId, bool enabled);
    EnterpriseEditReport setQueryTableRefreshOnLoad(const std::string& queryName, bool enabled);
    EnterpriseEditReport setOlapPivotCacheRefreshOnLoad(const std::string& partName, bool enabled);
    EnterpriseEditReport setPivotChartSourceName(const std::string& partName, const std::string& sourceName);

    const std::vector<PreservedPart>& preservedParts() const noexcept { return preservedParts_; }
    std::vector<PreservedPart>& preservedParts() noexcept { return preservedParts_; }
    // Validate modeled workbook invariants before serialization. save() runs
    // this by default; callers may also use it as an explicit preflight step.
    WorkbookValidationReport validate(const WorkbookValidationOptions& options = {}) const;

    const std::vector<PreservedRelationship>& preservedRelationships() const noexcept { return preservedRelationships_; }
    // True when the last load read a package using strict OOXML namespaces.
    bool strictNamespaces() const noexcept { return strictNamespaces_; }
    // Diagnostics from the most recent load.
    const LoadDiagnostics& diagnostics() const noexcept { return diagnostics_; }
private:
    // Internal path I/O implementations. Public load() provides a strong
    // exception guarantee by loading into a temporary Workbook first; public
    // save() optionally stages to a same-directory file for atomic replace.
    void loadInPlace(const std::filesystem::path& path, const LoadOptions& options);
    void saveInPlace(const std::filesystem::path& path, const SaveOptions& options) const;

    std::deque<Worksheet> sheets_;
    StableVector<NamedStyle> namedStyles_;
    StableVector<DefinedName> definedNames_;
    DocumentProperties properties_;
    WorkbookProtection protection_;
    CalcProperties calcProps_;
    CustomProperties customProps_;
    bool customPropertiesPartPresent_{false};
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
    // Exact source snapshots keyed by chart/series/cache reference. Keeping the
    // snapshot at Workbook level avoids intrusive Cell ownership callbacks while
    // still making cache synchronization dependency-aware and deterministic.
    std::unordered_map<std::string, std::string> chartCacheDependencySnapshots_;
};
}
