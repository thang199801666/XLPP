#pragma once
#include <XLPP/Worksheet/Worksheet.h>
#include <XLPP/Worksheet/Slicer.h>
#include <XLPP/Chart/Chartsheet.h>
#include <XLPP/Styles/NamedStyle.h>
#include <XLPP/Workbook/DefinedNames.h>
#include <XLPP/Workbook/StructuralEdit.h>
#include <XLPP/Workbook/ModelValidation.h>
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
    // P1R model-materialization guards. ZIP byte limits alone do not bound the
    // number of C++ objects an adversarial XML document can create. 0 means
    // unlimited for source compatibility.
    std::size_t maxWorksheets{0};
    std::size_t maxChartsheets{0};
    std::size_t maxCells{0};             // total materialized cells across workbook
    std::size_t maxSharedStrings{0};
    std::size_t maxDefinedNames{0};
    std::function<bool()> cancel{};
    std::function<void(std::size_t done, std::size_t total)> progress{};
    // Password used when the input is an ECMA-376 encrypted CFB package.
    // Plain ZIP workbooks ignore this field. Empty is a valid password.
    std::string passwordToOpen{};
    // Resource guard for password KDF work on untrusted encrypted documents.
    // 0 accepts the format maximum (10,000,000 Agile spins).
    std::uint32_t maxEncryptionSpinCount{1000000};
    // Optional guard for the decrypted inner OOXML package size before it is
    // materialized. 0 means use only maxFileBytes / ciphertext geometry.
    std::size_t maxDecryptedPackageBytes{0};
    // P1I encryption policy controls for untrusted inputs. Standard Encryption
    // has no authenticated data-integrity field, so applications can disable
    // it when only Agile packages are acceptable.
    bool allowStandardEncryption{true};
    // Agile packages produced by Microsoft Office normally carry DataIntegrity.
    // Keep false for compatibility; security-sensitive callers can require it.
    bool requireAgileDataIntegrity{false};
    // Bound the outer EncryptionInfo stream before XML/header parsing.
    // 0 means unlimited.
    std::size_t maxEncryptionInfoBytes{1024u * 1024u};
};

enum class ChartCacheDependencyKind { Title, Category, Value };

struct ChartCacheDependency {
    std::string ownerSheet;
    std::string sourceSheet;
    std::string chartStableId;
    std::size_t chartIndex{0};
    std::size_t seriesIndex{0};
    ChartCacheDependencyKind kind{ChartCacheDependencyKind::Value};
    std::string reference;
    CellReference first{};
    CellReference last{};
    bool supported{false};
    std::string issue;
};

struct ChartCacheSyncOptions {
    bool synchronizeTitles{true};
    bool synchronizeCategories{true};
    bool synchronizeValues{true};
    // When enabled, only references intersecting cells accessed through a
    // non-const Worksheet::cell()/range() since load/save are rebuilt.
    bool onlyChangedCells{false};
    // Formula cells with no cached value retain the matching existing chart
    // cache point rather than silently deleting it.
    bool preserveFormulaCachedValues{true};
    // Optionally clear the per-sheet touched-cell tracker after synchronization.
    bool clearTrackedChangesAfterSync{false};
    // Propagate tracked-cell changes through A1/range formula precedents,
    // including 2-D regions and workbook/local defined names that resolve to
    // A1 references. If a chart source formula depends on a changed precedent,
    // the existing chart
    // cache is preserved (XL++ is not a calculation engine) and host
    // recalculation can be requested below.
    bool propagateFormulaDependencies{true};
    bool requestHostRecalculationForFormulaDependencies{true};
    std::size_t maxFormulaDependencyDepth{8};
    // Unsupported/external/2-D references are preserved by default. If true,
    // their existing cache is cleared so the host is forced to recalculate it.
    bool clearUnsupportedReferences{false};
};

struct ChartCacheSyncReport {
    std::size_t chartsVisited{0};
    std::size_t seriesVisited{0};
    std::size_t dependenciesVisited{0};
    std::size_t dependenciesMatched{0};
    std::size_t dependenciesSkippedUnchanged{0};
    std::size_t cachesUpdated{0};
    std::size_t cachesCleared{0};
    std::size_t referencesSkipped{0};
    std::size_t formulaCachePointsReused{0};
    std::size_t formulaDependenciesVisited{0};
    std::size_t formulaDependenciesMatched{0};
    // P0X dependency-grammar diagnostics. Formula A1 references may be 2-D
    // even though materialized chart caches remain intentionally 1-D.
    std::size_t formulaReferencesVisited{0};
    std::size_t formulaReferencesResolved{0};
    std::size_t formulaReferencesSkipped{0};
    std::size_t definedNameDependenciesVisited{0};
    std::size_t definedNameDependenciesResolved{0};
    std::size_t definedNameDependenciesSkipped{0};
    // P0Y structured-table and bounded dynamic-name diagnostics.
    std::size_t structuredReferencesVisited{0};
    std::size_t structuredReferencesResolved{0};
    std::size_t structuredReferencesSkipped{0};
    std::size_t dynamicDefinedNamesVisited{0};
    std::size_t dynamicDefinedNamesResolved{0};
    std::size_t dynamicDefinedNamesSkipped{0};
    std::size_t staleFormulaCachesPreserved{0};
    bool hostRecalculationRequested{false};
    // Informational dependency diagnostics do not make success() false.
    std::vector<std::string> formulaDependencyDiagnostics;
    std::vector<std::string> warnings;
    bool success() const noexcept { return warnings.empty(); }
};


struct ChartStyleApplyReport {
    std::size_t seriesVisited{0};
    std::size_t seriesStyled{0};
    std::size_t colorsAvailable{0};
    std::size_t fillStylesAvailable{0};
    std::size_t lineStylesAvailable{0};
    std::size_t effectStylesAvailable{0};
    // P0Z rule-model application counters.
    std::size_t rulesAvailable{0};
    std::size_t rulesVisited{0};
    std::size_t rulesApplied{0};
    std::size_t targetsStyled{0};
    std::size_t effectReferencesResolved{0};
    std::vector<std::string> warnings;
    bool success() const noexcept { return warnings.empty(); }
};

enum class WorkbookSheetKind { Worksheet, Chartsheet };
enum class WorkbookSheetVisibility { Visible, Hidden, VeryHidden };

struct WorkbookSheetDescriptor {
    WorkbookSheetKind kind{WorkbookSheetKind::Worksheet};
    std::size_t kindIndex{0};
    std::string name;
    WorkbookSheetVisibility visibility{WorkbookSheetVisibility::Visible};
    bool active{false};
};

class Workbook {
    friend struct internal::WorkbookSheetOperationsAccess;
public:
    // Stable worksheet storage: std::deque never invalidates references on
    // insertion at either end, so `Worksheet&` / `Worksheet*` obtained from
    // addWorksheet/copyWorksheet/operator[] remain valid across further
    // worksheet insertion. Only removal of a sheet invalidates references to
    // that sheet.
    Worksheet& addWorksheet(std::string name);
    Worksheet& copyWorksheet(const Worksheet& source, std::string newName);
    // Adds a slicer control bound to a pivot-table field on the given
    // worksheet. Creates the slicer cache and slicer parts on save.
    Slicer& addSlicer(const std::string& worksheetName, Slicer slicer);
    const std::vector<Slicer>& slicers() const noexcept { return slicers_; }
    // Workbook-aware worksheet rename/removal. These operations repair
    // explicit worksheet-qualified references across formulas, defined names,
    // charts, Pivot sources and internal hyperlinks. Worksheet::rename() remains
    // a local object rename for source compatibility; prefer this API whenever
    // the worksheet belongs to a Workbook.
    bool renameWorksheet(const std::string& oldName, std::string newName);
    bool removeWorksheet(const std::string& name);
    Worksheet* worksheet(const std::string& name) noexcept;
    const Worksheet* worksheet(const std::string& name) const noexcept;
    Worksheet& operator[](std::size_t index);
    const Worksheet& operator[](std::size_t index) const;
    std::size_t index(const Worksheet& sheet) const;
    std::vector<std::string> sheetNames() const;
    std::size_t sheetCount() const noexcept { return sheets_.size(); }

    // Chart-only sheets participate in the workbook sheet order without
    // changing the legacy worksheet-only sheetCount()/sheetNames() semantics.
    Chartsheet& addChartsheet(std::string name, Chart chart);
    Chartsheet* chartsheet(const std::string& name) noexcept;
    const Chartsheet* chartsheet(const std::string& name) const noexcept;
    bool renameChartsheet(const std::string& oldName, std::string newName);
    bool removeChartsheet(const std::string& name);
    std::size_t chartsheetCount() const noexcept { return chartsheets_.size(); }
    std::deque<Chartsheet>& chartsheets() noexcept { return chartsheets_; }
    const std::deque<Chartsheet>& chartsheets() const noexcept { return chartsheets_; }
    std::size_t workbookSheetCount() const noexcept { return sheetOrder_.size(); }
    std::vector<std::string> workbookSheetNames() const;
    std::vector<WorkbookSheetDescriptor> workbookSheets() const;
    WorkbookSheetVisibility workbookSheetVisibility(std::size_t index) const;
    void setWorkbookSheetVisibility(std::size_t index, WorkbookSheetVisibility visibility);
    std::size_t activeWorkbookSheetIndex() const noexcept { return activeWorkbookSheetIndex_; }
    void setActiveWorkbookSheetIndex(std::size_t index);
    bool setActiveWorkbookSheet(const std::string& name);
    std::size_t firstVisibleWorkbookSheetIndex() const noexcept;
    // Reorders the mixed worksheet/chartsheet workbook tab order.
    void moveWorkbookSheet(std::size_t fromIndex, std::size_t toIndex);
    // Spreadsheet template package identity. When true, save() emits the
    // template workbook content type (.xltx, or .xltm when VBA is present).
    // The flag is also recovered automatically when loading template packages.
    bool isTemplate() const noexcept { return template_; }
    void setTemplate(bool value) noexcept { template_ = value; }
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

    // Workbook-aware structural edits. Unlike Worksheet::insertRows()/etc.,
    // these operations rewrite references across every worksheet plus workbook
    // defined names, charts and Pivot cache sources that point at the edited
    // sheet. Use these APIs for reference-safe structural mutation.
    StructuralEditReport insertRows(const std::string& worksheetName, std::size_t index, std::size_t amount = 1,
                                    const StructuralEditOptions& options = {});
    StructuralEditReport deleteRows(const std::string& worksheetName, std::size_t index, std::size_t amount = 1,
                                    const StructuralEditOptions& options = {});
    StructuralEditReport insertColumns(const std::string& worksheetName, std::size_t index, std::size_t amount = 1,
                                       const StructuralEditOptions& options = {});
    StructuralEditReport deleteColumns(const std::string& worksheetName, std::size_t index, std::size_t amount = 1,
                                       const StructuralEditOptions& options = {});

    // Validates semantic invariants of the in-memory workbook model. This is
    // complementary to package/relationship validation and is safe to call
    // before save or after complex structural edits.
    WorkbookModelValidationReport validateModelIntegrity() const;
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

    void clear() { sheets_.clear(); chartsheets_.clear(); sheetOrder_.clear(); namedStyles_.clear(); definedNames_.clear(); slicers_.clear(); properties_ = {}; protection_ = {}; date1904_ = false; template_ = false; activeWorkbookSheetIndex_ = 0; firstVisibleWorkbookSheetIndex_ = 0; preservedParts_.clear(); preservedRelationships_.clear(); sourceWorkbookXml_.clear(); sourceSheetParts_.clear(); sourceSheetXml_.clear(); sourceSheetNames_.clear(); cachedSheetXml_.clear(); cachedSheetXmlStrict_ = false; cachedSheetXmlDate1904_ = false; generatedVbaProject_ = false; strictNamespaces_ = false; diagnostics_ = {}; calcProps_ = {}; customProps_ = {}; }
    void load(const std::filesystem::path& path);
    void load(const std::filesystem::path& path, const LoadOptions& options);
    void load(std::istream& stream);
    void load(std::istream& stream, const LoadOptions& options);
    void save(const std::filesystem::path& path) const;
    void save(const std::filesystem::path& path, const SaveOptions& options) const;
    void save(std::ostream& stream) const;
    void save(std::ostream& stream, const SaveOptions& options) const;
    static bool isPasswordEncryptedFile(const std::filesystem::path& path) noexcept;
    static PackageEncryptionInfo inspectPasswordEncryptionFile(const std::filesystem::path& path);

    // Rebuild chart title/category/value caches from worksheet references.
    // Supported direct cache sources include single-cell/one-dimensional A1
    // ranges, resolvable defined names, and one-dimensional structured table
    // references. External workbooks, unions, and 2-D cache materialization are
    // intentionally skipped; 2-D/table ranges can still participate in formula
    // dependency tracking.
    ChartCacheSyncReport synchronizeChartCaches(const ChartCacheSyncOptions& options = {});
    ChartCacheSyncReport synchronizeChangedChartCaches(ChartCacheSyncOptions options = {});
    std::vector<ChartCacheDependency> chartCacheDependencies() const;
    void clearChartCacheChangeTracking() noexcept;
    // Apply the imported chart color-style palette to series fill/line/marker
    // colors using the workbook theme. This is intentionally a conservative
    // foundation: it does not attempt to reproduce Excel's full chart-style
    // matrix/effect engine.
    ChartStyleApplyReport applyChartColorStyle(const std::string& worksheetName, const std::string& chartStableId,
                                                bool applyFill = true, bool applyLine = true, bool applyMarker = true);
    // Apply one fill/line entry from the workbook theme fmtScheme to all series
    // of an imported chart. Indices are zero-based. Theme effect styles are
    // materialized for inspection but not blindly stamped onto a target; the
    // Office chart-style rules decide which chart object receives an effect.
    ChartStyleApplyReport applyChartThemeStyleMatrix(const std::string& worksheetName, const std::string& chartStableId,
                                                      std::size_t fillStyleIndex, std::size_t lineStyleIndex,
                                                      bool applyMarker = true);
    // Resolve the imported Office chart-style rule model against the workbook
    // theme/color-style resources and apply supported targets conservatively.
    // P0Z materializes chartArea, plotArea, dataPoint/dataPointLine,
    // dataPointMarker, legend and axis line/fill rules. Text/effect references
    // are inspected/resolved but are not serialized into unsupported targets.
    ChartStyleApplyReport applyChartStyleRules(const std::string& worksheetName,
                                                const std::string& chartStableId,
                                                bool applyMarkers = true);
    // Patch refresh/cache-retention options directly on an imported physical
    // PivotCache without regenerating sibling PivotTables that share it. The
    // cache identity is discovered from the loaded PivotTable model.
    bool updateImportedPivotCacheOptions(const std::string& worksheetName,
                                         const std::string& pivotTableName,
                                         const PivotCacheOptionsPatch& patch);
    // Patch one imported cacheField start tag in-place, preserving sharedItems,
    // fieldGroup, cache records and every sibling PivotTable part byte-for-byte.
    bool updateImportedPivotCacheField(const std::string& worksheetName,
                                       const std::string& pivotTableName,
                                       std::size_t fieldIndex,
                                       const PivotCacheFieldPatch& patch);
    // Patch one <item> under an imported pivotField without regenerating the
    // PivotTable or shared PivotCache. fieldIndex is the cache/pivotField index;
    // itemIndex is the ordinal inside that field's <items> collection.
    bool updateImportedPivotFieldItem(const std::string& worksheetName,
                                      const std::string& pivotTableName,
                                      std::size_t fieldIndex,
                                      std::size_t itemIndex,
                                      const PivotFieldItemPatch& patch);
    // Patch one imported PivotFilter in-place. Existing advanced nested
    // autoFilter criteria remain byte-preserved unless autoFilterXml is supplied.
    bool updateImportedPivotFilter(const std::string& worksheetName,
                                   const std::string& pivotTableName,
                                   std::size_t filterIndex,
                                   const PivotFilterPatch& patch);
    // Patch one imported <dataField> in-place without regenerating the
    // PivotTable or shared PivotCache.
    bool updateImportedPivotDataField(const std::string& worksheetName,
                                      const std::string& pivotTableName,
                                      std::size_t dataFieldIndex,
                                      const PivotDataFieldPatch& patch);
    // Patch one imported <pageField> in-place while preserving the rest of the
    // PivotTable definition byte-for-byte.
    bool updateImportedPivotPageField(const std::string& worksheetName,
                                      const std::string& pivotTableName,
                                      std::size_t pageFieldIndex,
                                      const PivotPageFieldPatch& patch);
    // Patch one value in an imported physical pivotCacheRecords part without
    // regenerating the owning PivotTables or cache definition.
    bool updateImportedPivotCacheRecordValue(const std::string& worksheetName,
                                             const std::string& pivotTableName,
                                             std::size_t recordIndex,
                                             std::size_t fieldIndex,
                                             const PivotCacheRecordValuePatch& patch);
    // OLAP pivot metadata. OLAP caches bind to cube/data-connection sources
    // rather than a worksheet range; XL++ inspects the olapInfo/cacheSource
    // identity and patches selected attributes in-place without regenerating
    // the physical cache or truncating unmodeled olapInfo/calculatedMember XML.
    bool updateImportedPivotOlapSource(const std::string& worksheetName,
                                       const std::string& pivotTableName,
                                       const PivotOlapSourcePatch& patch);
    // Patch one <calculatedMember> in an imported OLAP pivotCacheDefinition.
    bool updateImportedPivotCalculatedMember(const std::string& worksheetName,
                                             const std::string& pivotTableName,
                                             std::size_t memberIndex,
                                             const PivotCalculatedMemberPatch& patch);
    // Validate DrawingML PivotChart sources against loaded PivotTable ownership
    // and PivotTable chartFormat/chartFormats metadata. This is read-only.
    PivotChartLinkValidationReport validatePivotChartLinks() const;
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
    void setVbaClassModuleText(std::string moduleName, std::string source);
    // Add or replace a designer module (for example a UserForm) together with
    // its raw recursive designer storage. The source is editable; unsupported
    // MS-OFORMS streams remain byte-preserved.
    void setVbaDesignerModule(std::string moduleName, std::string source, VbaDesignerStorage storage,
                              std::string designerClassId = "{AC9F2F90-E877-11CE-9F68-00AA00574A4F}");
    std::vector<VbaDesignerStorage> vbaDesignerStorages() const;
    void setVbaDesignerStorage(VbaDesignerStorage storage);
    bool removeVbaDesignerStorage(const std::string& storageName);
    // Decode the MS-OFORMS Form stream ("f") for an existing Designer/UserForm
    // without interpreting embedded child controls. Malformed streams return a
    // diagnostic rather than being silently accepted.
    VbaUserFormInspection inspectVbaUserForm(const std::string& storageName) const;
    // Semantically patch properties already materialized in the Form PropMask.
    // Unknown StreamData/SiteData and all other designer streams remain exact.
    bool updateVbaUserFormProperties(const std::string& storageName,
                                     const VbaUserFormPropertiesPatch& patch);
    // Inspect the FormSiteData/OleSiteConcreteControl records for embedded
    // controls. This is site-level semantics; control-specific stream "o" bytes
    // are left untouched.
    VbaUserFormControlsInspection inspectVbaUserFormControls(const std::string& storageName) const;
    // Decode the object-stream slice belonging to one embedded control. P1F
    // provides semantic CommandButton/Label parsing and validated headers for
    // the other built-in MSForms control families.
    VbaUserFormControlObjectInspection inspectVbaUserFormControlObject(const std::string& storageName,
                                                                       std::size_t controlIndex) const;
    // Patch semantic properties already present in a CommandButton/Label object
    // PropMask. If object size changes (for example Unicode caption growth),
    // XL++ rewrites both stream o and OleSite.ObjectStreamSize safely.
    bool updateVbaUserFormControlObject(const std::string& storageName,
                                        std::size_t controlIndex,
                                        const VbaUserFormControlObjectPatch& patch);
    // Patch properties already present in a control site's SitePropMask. String
    // growth/shrink is supported and CountOfBytes/cbSite are rewritten safely.
    bool updateVbaUserFormControlSite(const std::string& storageName,
                                      std::size_t controlIndex,
                                      const VbaUserFormControlSitePatch& patch);
    VbaDesignerValidationReport validateVbaDesignerProject() const;
    // Edit host document modules by code name (Sheet1, Sheet2, ... or
    // ThisWorkbook). The module identity remains synchronized with the workbook.
    void setVbaDocumentModuleText(std::string moduleName, std::string source);
    void setVbaModule(VbaModule module);
    std::optional<std::string> vbaModuleText(const std::string& moduleName) const;
    std::vector<VbaModule> vbaModules() const;
    bool removeVbaModule(const std::string& moduleName);
    VbaProjectInfo vbaProjectInfo() const;
    void setVbaProjectInfo(VbaProjectInfo info);

    const std::vector<PreservedPart>& preservedParts() const noexcept { return preservedParts_; }
    std::vector<PreservedPart>& preservedParts() noexcept { return preservedParts_; }
    const std::vector<PreservedRelationship>& preservedRelationships() const noexcept { return preservedRelationships_; }
    // True when the last load read a package using strict OOXML namespaces.
    bool strictNamespaces() const noexcept { return strictNamespaces_; }
    // Diagnostics from the most recent load.
    const LoadDiagnostics& diagnostics() const noexcept { return diagnostics_; }
private:
    // P1R: transactional load implementation. Public load() parses into a
    // temporary Workbook and commits only after the package is fully accepted.
    void loadInPlace(const std::filesystem::path& path, const LoadOptions& options);

    StructuralEditReport applyStructuralEdit(const std::string& worksheetName, std::size_t index, std::size_t amount,
                                             StructuralEditKind kind, const StructuralEditOptions& options);
    StructuralEditReport applyStructuralEditImpl(const std::string& worksheetName, std::size_t index, std::size_t amount,
                                                 StructuralEditKind kind, const StructuralEditOptions& options);

    struct SheetOrderEntry {
        WorkbookSheetKind kind{WorkbookSheetKind::Worksheet};
        std::size_t kindIndex{0};
        WorkbookSheetVisibility visibility{WorkbookSheetVisibility::Visible};
    };
    std::deque<Worksheet> sheets_;
    std::deque<Chartsheet> chartsheets_;
    std::vector<SheetOrderEntry> sheetOrder_;
    std::vector<NamedStyle> namedStyles_;
    std::vector<DefinedName> definedNames_;
    std::vector<Slicer> slicers_;
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
    bool template_{false};
    std::size_t activeWorkbookSheetIndex_{0};
    std::size_t firstVisibleWorkbookSheetIndex_{0};
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
