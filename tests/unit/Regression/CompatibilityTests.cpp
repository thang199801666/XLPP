#include <XLPP/XLPP.h>
#include "Package/Zip/ZipArchive.h"
#include "Package/Opc/RelationshipGraph.h"
#include "VBA/VbaProjectBinary.h"
#include <filesystem>
#include <fstream>
#include <string>
#include <vector>

#include "../TestFramework.h"
#include "RegressionTests.h"

void testVbaProjectPackageLifecycle(TestContext& test) {
    const auto macroPath = std::filesystem::temp_directory_path() / "xlpp_vba_package.xlsm";
    const auto noMacroPath = std::filesystem::temp_directory_path() / "xlpp_vba_removed.xlsx";
    const std::vector<unsigned char> bytes{'V','B','A',0,'X','L','P','P'};
    xlpp::Workbook workbook;
    workbook.addWorksheet("MacroHost").cell("A1").setValue("VBA host");
    workbook.setVbaProject(bytes);
    test.checkTrue(workbook.hasVbaProject(), "Workbook reports attached VBA project");
    workbook.save(macroPath);

    auto zip = xlpp::internal::ZipArchive::open(macroPath);
    test.checkTrue(zip.contains("xl/vbaProject.bin"), "VBA project binary is packaged");
    const auto packaged = zip.get("xl/vbaProject.bin");
    test.checkEqual(packaged.size(), bytes.size(), "VBA binary size is preserved");
    test.checkEqual(packaged[3], '\0', "VBA binary NUL byte is preserved");
    const auto types = zip.get("[Content_Types].xml");
    const auto rels = zip.get("xl/_rels/workbook.xml.rels");
    test.checkTrue(types.find("application/vnd.ms-excel.sheet.macroEnabled.main+xml") != std::string::npos,
                   "Macro-enabled workbook main content type is emitted");
    test.checkTrue(types.find("application/vnd.ms-office.vbaProject") != std::string::npos,
                   "VBA project content type is emitted");
    test.checkTrue(rels.find("/vbaProject\" Target=\"vbaProject.bin\"") != std::string::npos,
                   "Workbook relationship points to VBA project");
    test.checkTrue(xlpp::internal::RelationshipGraph::fromArchive(zip).validate().ok(),
                   "Macro package passes OPC relationship and content-type validation");

    xlpp::Workbook loaded;
    loaded.load(macroPath);
    test.checkTrue(loaded.hasVbaProject(), "Loaded XLSM reports VBA project");
    test.checkTrue(loaded.removeVbaProject(), "removeVbaProject removes attached project");
    test.checkTrue(!loaded.hasVbaProject(), "VBA project is absent after removal");
    loaded.save(noMacroPath);
    zip = xlpp::internal::ZipArchive::open(noMacroPath);
    test.checkTrue(!zip.contains("xl/vbaProject.bin"), "Removed VBA binary is not packaged");
    test.checkTrue(zip.get("[Content_Types].xml").find("macroEnabled") == std::string::npos,
                   "Workbook returns to normal XLSX content type after VBA removal");
    test.checkTrue(zip.get("xl/_rels/workbook.xml.rels").find("/vbaProject") == std::string::npos,
                   "VBA relationship is removed");
    std::filesystem::remove(macroPath);
    std::filesystem::remove(noMacroPath);
}

void testWorkbookClearConstAndVbaFileApi(TestContext& test) {
    const auto binPath = std::filesystem::temp_directory_path() / "xlpp_vba_file_api.bin";
    { std::ofstream stream(binPath, std::ios::binary); stream << "VBAPROJECT"; }
    xlpp::Workbook workbook;
    auto& sheet = workbook.addWorksheet("Data");
    sheet.cell("A1").setValue("value");
    workbook.addNamedStyle(xlpp::NamedStyle("Accent", xlpp::Style{}));
    workbook.addDefinedName(xlpp::DefinedName("Name", "Data!$A$1"));
    workbook.properties().setTitle("title");
    workbook.protection().setLockStructure(true);
    workbook.calcProperties().setCalcMode("manual");
    workbook.customProperties().add(xlpp::CustomProperty("Custom", "Text"));
    workbook.setDate1904(true);
    workbook.preservedParts().push_back({"custom/item.bin", "x", "application/octet-stream", "bin", "application/octet-stream", false});
    workbook.addVbaProject(binPath);
    test.checkTrue(workbook.hasVbaProject(), "VBA file API attaches project");

    const xlpp::Workbook& constant = workbook;
    test.checkEqual(constant.worksheets().size(), std::size_t{1}, "Const worksheets accessor");
    test.checkTrue(constant.worksheet("Data") != nullptr, "Const worksheet lookup");
    test.checkTrue(constant.namedStyle("Accent") != nullptr, "Const named-style lookup");
    test.checkTrue(constant.definedName("Name") != nullptr, "Const defined-name lookup");
    test.checkEqual(constant.properties().title(), std::string("title"), "Const properties accessor");
    test.checkTrue(constant.protection().lockStructure(), "Const protection accessor");
    test.checkEqual(constant.calcProperties().calcMode(), std::string("manual"), "Const calc-properties accessor");
    test.checkEqual(constant.customProperties().items().size(), std::size_t{1}, "Const custom-properties accessor");
    test.checkEqual(constant.preservedParts().size(), std::size_t{2}, "Const preserved-parts accessor includes VBA");

    workbook.clear();
    test.checkEqual(workbook.sheetCount(), std::size_t{0}, "Workbook clear removes sheets");
    test.checkTrue(workbook.namedStyles().empty(), "Workbook clear removes named styles");
    test.checkTrue(workbook.definedNames().empty(), "Workbook clear removes defined names");
    test.checkTrue(workbook.properties().title().empty(), "Workbook clear resets properties");
    test.checkTrue(!workbook.protection().lockStructure(), "Workbook clear resets protection");
    test.checkTrue(!workbook.date1904(), "Workbook clear resets date system");
    test.checkTrue(workbook.customProperties().items().empty(), "Workbook clear removes custom properties");
    test.checkTrue(workbook.preservedParts().empty(), "Workbook clear removes preserved parts");
    test.checkTrue(!workbook.hasVbaProject(), "Workbook clear removes VBA project");
    test.checkTrue(!workbook.removeVbaProject(), "Removing absent VBA project reports false");
    std::filesystem::remove(binPath);
}

void writeExternalReaderFixture(const std::filesystem::path& path,
                                const std::string& workbookXml,
                                const std::string& sheetXml,
                                const std::string& stylesXml,
                                const std::string& sharedStringsXml = {},
                                const std::string& sheetRelationships = {},
                                const std::string& commentsXml = {},
                                const std::string& corePropertiesXml = {},
                                const std::string& customPropertiesXml = {}) {
    constexpr auto packageNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    constexpr auto documentNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    xlpp::internal::ZipArchive zip;
    std::string contentTypes =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
    if (!sharedStringsXml.empty())
        contentTypes += "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>";
    if (!commentsXml.empty())
        contentTypes += "<Override PartName=\"/xl/comments1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>";
    if (!corePropertiesXml.empty())
        contentTypes += "<Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>";
    if (!customPropertiesXml.empty())
        contentTypes += "<Override PartName=\"/docProps/custom.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.custom-properties+xml\"/>";
    contentTypes += "</Types>";
    zip.add("[Content_Types].xml", contentTypes);

    std::string rootRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"";
    rootRels += packageNs;
    rootRels += "\"><Relationship Id=\"rIdWorkbook\" Type=\"";
    rootRels += documentNs;
    rootRels += "/officeDocument\" Target=\"xl/workbook.xml\"/>";
    if (!corePropertiesXml.empty())
        rootRels += "<Relationship Id=\"rIdCore\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"docProps/core.xml\"/>";
    if (!customPropertiesXml.empty()) {
        rootRels += "<Relationship Id=\"rIdCustom\" Type=\"";
        rootRels += documentNs;
        rootRels += "/custom-properties\" Target=\"docProps/custom.xml\"/>";
    }
    rootRels += "</Relationships>";
    zip.add("_rels/.rels", rootRels);
    zip.add("xl/workbook.xml", workbookXml);

    std::string workbookRels = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"";
    workbookRels += packageNs;
    workbookRels += "\"><Relationship Id=\"rIdSheet1\" Type=\"";
    workbookRels += documentNs;
    workbookRels += "/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rIdStyles\" Type=\"";
    workbookRels += documentNs;
    workbookRels += "/styles\" Target=\"styles.xml\"/>";
    if (!sharedStringsXml.empty()) {
        workbookRels += "<Relationship Id=\"rIdShared\" Type=\"";
        workbookRels += documentNs;
        workbookRels += "/sharedStrings\" Target=\"sharedStrings.xml\"/>";
    }
    workbookRels += "</Relationships>";
    zip.add("xl/_rels/workbook.xml.rels", workbookRels);
    zip.add("xl/worksheets/sheet1.xml", sheetXml);
    zip.add("xl/styles.xml", stylesXml);
    if (!sharedStringsXml.empty()) zip.add("xl/sharedStrings.xml", sharedStringsXml);
    if (!sheetRelationships.empty()) zip.add("xl/worksheets/_rels/sheet1.xml.rels", sheetRelationships);
    if (!commentsXml.empty()) zip.add("xl/comments1.xml", commentsXml);
    if (!corePropertiesXml.empty()) zip.add("docProps/core.xml", corePropertiesXml);
    if (!customPropertiesXml.empty()) zip.add("docProps/custom.xml", customPropertiesXml);
    zip.save(path);
}

void testVbaSourceTextBuildAndRead(TestContext& test) {
    const auto firstPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source.xlsm";
    const auto secondPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source_updated.xlsm";
    const auto sheetMutationPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_sheet_mutation.xlsm";
    const auto removedPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_source_removed.xlsx";
    const auto advancedPath = std::filesystem::temp_directory_path() / "xlpp_vba_advanced_modules.xlsm";
    const auto topologyPath = std::filesystem::temp_directory_path() / "xlpp_vba_document_topology.xlsm";
    const auto exportedProjectPath = std::filesystem::temp_directory_path() / "xlpp_vba_exported_project.bin";
    const std::string source =
        "Option Explicit\n"
        "Public Sub XLPP_Hello()\n"
        "    Range(\"A2\").Value = 42\n"
        "End Sub\n"
        "Public Function AddTwo(ByVal value As Double) As Double\n"
        "    AddTwo = value + 2\n"
        "End Function\n";

    xlpp::Workbook workbook;
    workbook.addWorksheet("MacroHost").cell("A1").setValue("=AddTwo(40)");
    workbook.setVbaModuleText("MathModule", source);
    // Add a worksheet after creating the VBA project. save() must regenerate
    // document modules so Sheet2 and sheetPr codeName stay synchronized.
    workbook.addWorksheet("AddedAfterVba").cell("A1").setValue("document-module sync");
    test.checkTrue(workbook.hasVbaProject(), "VBA text API creates a binary project");
    const auto inMemory = workbook.vbaModuleText("mathmodule");
    test.checkTrue(inMemory.has_value(), "VBA module lookup is case-insensitive");
    test.checkEqual(*inMemory,
                    std::string("Option Explicit\r\nPublic Sub XLPP_Hello()\r\n    Range(\"A2\").Value = 42\r\nEnd Sub\r\nPublic Function AddTwo(ByVal value As Double) As Double\r\n    AddTwo = value + 2\r\nEnd Function\r\n"),
                    "VBA source is normalized to CRLF");
    test.checkTrue(inMemory->find("Public Sub XLPP_Hello()") != std::string::npos,
                   "Generated standard module contains a public parameterless macro");
    test.checkTrue(inMemory->find("Private Sub XLPP_Hello") == std::string::npos,
                   "Generated macro is not private");
    workbook.save(firstPath);

    auto zip = xlpp::internal::ZipArchive::open(firstPath);
    test.checkTrue(zip.contains("xl/vbaProject.bin"), "Text-generated vbaProject.bin is packaged");
    test.checkTrue(zip.get("[Content_Types].xml").find("application/vnd.ms-excel.sheet.macroEnabled.main+xml") != std::string::npos,
                   "Workbook content type is macro-enabled");
    test.checkTrue(zip.get("xl/_rels/workbook.xml.rels").find("/relationships/vbaProject") != std::string::npos,
                   "Workbook relationship targets vbaProject.bin");
    const auto& binary = zip.get("xl/vbaProject.bin");
    test.checkTrue(binary.size() >= 512, "Generated VBA project is a compound file");
    test.checkEqual(static_cast<unsigned char>(binary[0]), static_cast<unsigned char>(0xD0), "CFB signature byte 0");
    test.checkEqual(static_cast<unsigned char>(binary[1]), static_cast<unsigned char>(0xCF), "CFB signature byte 1");
    test.checkTrue(zip.get("xl/workbook.xml").find("codeName=\"ThisWorkbook\"") != std::string::npos,
                   "Macro workbook emits ThisWorkbook code name");
    test.checkTrue(zip.get("xl/worksheets/sheet1.xml").find("<sheetPr codeName=\"Sheet1\">") != std::string::npos,
                   "Macro worksheet emits the Sheet1 document-module code name");
    test.checkTrue(zip.get("xl/worksheets/sheet2.xml").find("<sheetPr codeName=\"Sheet2\">") != std::string::npos,
                   "Worksheet added after VBA creation emits a synchronized Sheet2 code name");

    xlpp::Workbook loaded;
    loaded.load(firstPath);
    const auto modules = loaded.vbaModules();
    test.checkEqual(modules.size(), std::size_t{4}, "Reader exposes both sheets, workbook, and standard VBA modules");
    test.checkEqual(modules[0].name, std::string("Sheet1"), "First worksheet document module is present");
    test.checkEqual(static_cast<int>(modules[0].type), static_cast<int>(xlpp::VbaModuleType::Document), "First worksheet module type is document");
    test.checkEqual(modules[1].name, std::string("Sheet2"), "Worksheet added after VBA creation has a document module");
    test.checkEqual(static_cast<int>(modules[1].type), static_cast<int>(xlpp::VbaModuleType::Document), "Second worksheet module type is document");
    test.checkEqual(modules[2].name, std::string("ThisWorkbook"), "Workbook document module is present");
    test.checkEqual(static_cast<int>(modules[2].type), static_cast<int>(xlpp::VbaModuleType::Document), "Workbook module type is document");
    test.checkEqual(modules[3].name, std::string("MathModule"), "Standard module name is read from dir stream");
    test.checkEqual(static_cast<int>(modules[3].type), static_cast<int>(xlpp::VbaModuleType::Standard), "Visible macro resides in a standard module");
    test.checkEqual(loaded.vbaModuleText("MathModule").value(), *inMemory,
                    "VBA source survives XLSM save/load");

    // A generated project loaded from disk remains identifiable as XL++-owned.
    // Adding a sheet without editing VBA must still rebuild document modules.
    loaded.addWorksheet("AddedAfterReload").cell("A1").setValue("reload sync");
    loaded.save(sheetMutationPath);
    xlpp::Workbook sheetMutated;
    sheetMutated.load(sheetMutationPath);
    const auto mutatedModules = sheetMutated.vbaModules();
    test.checkEqual(mutatedModules.size(), std::size_t{5},
                    "Loaded XL++ VBA project rebuilds document modules after worksheet insertion");
    test.checkEqual(mutatedModules[2].name, std::string("Sheet3"),
                    "Sheet inserted after reload receives a Sheet3 VBA document module");
    test.checkTrue(sheetMutated.vbaModuleText("MathModule").has_value(),
                   "Standard macro module survives sheet-driven VBA project regeneration");

    loaded.setVbaModuleText("MathModule", "Public Sub Updated()\n    Range(\"A2\").Value = 99\nEnd Sub");
    loaded.setVbaModuleText("SecondModule", "Public Const Answer As Long = 42\n");
    loaded.save(secondPath);
    xlpp::Workbook updated;
    updated.load(secondPath);
    test.checkTrue(updated.vbaModuleText("MathModule")->find("Updated") != std::string::npos,
                   "Existing VBA module source can be replaced after load");
    test.checkTrue(updated.vbaModuleText("SecondModule")->find("Answer") != std::string::npos,
                   "Second VBA source module can be added after load");
    test.checkTrue(updated.removeVbaModule("MathModule"), "Individual VBA module can be removed");
    test.checkTrue(!updated.vbaModuleText("MathModule").has_value(), "Removed VBA module is absent from reader");
    test.checkTrue(updated.removeVbaModule("SecondModule"), "Last standard VBA module can be removed");
    test.checkTrue(!updated.hasVbaProject(), "Removing the last standard module removes the VBA project");
    updated.save(removedPath);
    zip = xlpp::internal::ZipArchive::open(removedPath);
    test.checkTrue(!zip.contains("xl/vbaProject.bin"), "Workbook returns to XLSX after all text modules are removed");

    // Class modules, document-event modules, module flags, stable code names,
    // raw project export, and topology mutations are all source-authorable.
    xlpp::Workbook advanced;
    auto& eventSheet = advanced.addWorksheet("Event Host");
    eventSheet.setVbaCodeName("EventSheet");
    auto& keepSheet = advanced.addWorksheet("Keep Host");
    keepSheet.setVbaCodeName("KeepSheet");
    xlpp::VbaProjectProperties projectProperties;
    projectProperties.name = "XLPPMacros";
    projectProperties.description = "XL++ generated VBA project";
    projectProperties.helpFile = "xlpp-vba-help.chm";
    projectProperties.helpContextId = 42;
    projectProperties.constants = "FeatureX = 1:DebugMode = 0";
    advanced.setVbaProjectProperties(projectProperties);
    advanced.setVbaDocumentModuleText("ThisWorkbook",
        "Option Explicit\nPrivate Sub Workbook_Open()\n    Worksheets(1).Range(\"A1\").Value = \"opened\"\nEnd Sub\n");
    advanced.setVbaDocumentModuleText("KeepSheet",
        "Option Explicit\nPrivate Sub Worksheet_Activate()\n    Range(\"A2\").Value = \"active\"\nEnd Sub\n");
    advanced.setVbaClassModuleText("Greeter",
        "Option Explicit\nPublic Function Message() As String\n    Message = \"hello\"\nEnd Function\n",
        true, true);
    advanced.save(advancedPath);

    xlpp::Workbook advancedLoaded;
    advancedLoaded.load(advancedPath);
    test.checkEqual(advancedLoaded[0].vbaCodeName(), std::string("EventSheet"),
                    "Worksheet VBA code name survives XLSM load");
    test.checkEqual(advancedLoaded[1].vbaCodeName(), std::string("KeepSheet"),
                    "Second worksheet VBA code name survives XLSM load");
    const auto loadedProjectProperties = advancedLoaded.vbaProjectProperties();
    test.checkEqual(loadedProjectProperties.name, std::string("XLPPMacros"), "VBA project name round-trips");
    test.checkEqual(loadedProjectProperties.description, std::string("XL++ generated VBA project"), "VBA project description round-trips");
    test.checkEqual(loadedProjectProperties.helpFile, std::string("xlpp-vba-help.chm"), "VBA project help file round-trips");
    test.checkEqual(loadedProjectProperties.helpContextId, std::uint32_t{42}, "VBA project help context round-trips");
    test.checkEqual(loadedProjectProperties.constants, std::string("FeatureX = 1:DebugMode = 0"), "VBA conditional compilation constants round-trip");
    const auto advancedModules = advancedLoaded.vbaModules();
    const auto greeter = std::find_if(advancedModules.begin(), advancedModules.end(), [](const auto& module) {
        return module.name == "Greeter";
    });
    test.checkTrue(greeter != advancedModules.end(), "Class VBA module is readable");
    if (greeter != advancedModules.end()) {
        test.checkEqual(static_cast<int>(greeter->type), static_cast<int>(xlpp::VbaModuleType::Class),
                        "PROJECT stream distinguishes class modules from document modules");
        test.checkTrue(greeter->readOnly, "VBA MODULEREADONLY flag round-trips");
        test.checkTrue(greeter->privateModule, "VBA MODULEPRIVATE flag round-trips");
    }
    test.checkTrue(advancedLoaded.vbaModuleText("ThisWorkbook")->find("Workbook_Open") != std::string::npos,
                   "ThisWorkbook event source round-trips");
    test.checkTrue(advancedLoaded.vbaModuleText("KeepSheet")->find("Worksheet_Activate") != std::string::npos,
                   "Worksheet event source round-trips");
    test.checkTrue(!advancedLoaded.vbaProjectBytes().empty(), "Raw vbaProject.bin bytes are exposed");
    advancedLoaded.saveVbaProject(exportedProjectPath);
    test.checkEqual(static_cast<std::uintmax_t>(advancedLoaded.vbaProjectBytes().size()),
                    std::filesystem::file_size(exportedProjectPath),
                    "Raw VBA project export writes the exact binary payload");
    test.checkTrue(!advancedLoaded.hasVbaSignature(), "Unsigned generated project reports no VBA signature");

    // Copying a worksheet copies its document-event code under a fresh code
    // name. Deleting a preceding worksheet must not retarget the event module.
    auto* keepBeforeCopy = advancedLoaded.worksheet("Keep Host");
    auto& cloned = advancedLoaded.copyWorksheet(*keepBeforeCopy, "Keep Copy");
    const auto clonedCodeName = cloned.vbaCodeName();
    test.checkTrue(clonedCodeName != "KeepSheet", "Copied worksheet receives a unique VBA code name");
    test.checkTrue(advancedLoaded.vbaModuleText(clonedCodeName)->find("Worksheet_Activate") != std::string::npos,
                   "Copied worksheet receives copied document-event source");
    test.checkTrue(advancedLoaded.removeWorksheet("Event Host"), "Worksheet preceding an event host can be removed");
    test.checkTrue(advancedLoaded.vbaModuleText("KeepSheet")->find("Worksheet_Activate") != std::string::npos,
                   "Worksheet deletion does not retarget a stable VBA document module");
    advancedLoaded.save(topologyPath);
    xlpp::Workbook topologyLoaded;
    topologyLoaded.load(topologyPath);
    test.checkEqual(topologyLoaded[0].vbaCodeName(), std::string("KeepSheet"),
                    "Stable worksheet code name survives deletion and reload");
    test.checkTrue(topologyLoaded.vbaModuleText("KeepSheet")->find("Worksheet_Activate") != std::string::npos,
                   "Worksheet event source survives deletion and reload");
    test.checkTrue(topologyLoaded.vbaModuleText(clonedCodeName)->find("Worksheet_Activate") != std::string::npos,
                   "Copied worksheet event source survives reload");
    test.checkEqual(topologyLoaded.vbaProjectProperties().name, std::string("XLPPMacros"),
                    "Project metadata survives module/topology regeneration");
    test.checkEqual(topologyLoaded.vbaProjectProperties().constants, std::string("FeatureX = 1:DebugMode = 0"),
                    "Conditional compilation constants survive module/topology regeneration");

    // Binary-owned projects are preserved verbatim but source mutation is
    // intentionally blocked because rebuilding would discard unknown
    // references, designer streams, password/signature state, or metadata.
    xlpp::Workbook binaryOwned;
    binaryOwned.addWorksheet("Binary Host");
    binaryOwned.setVbaProject(topologyLoaded.vbaProjectBytes());
    bool binaryEditRejected = false;
    try { binaryOwned.setVbaModuleText("UnsafeEdit", "Sub X(): End Sub"); }
    catch (const std::runtime_error&) { binaryEditRejected = true; }
    test.checkTrue(binaryEditRejected, "Externally supplied VBA project rejects destructive source rewrite");
    test.checkEqual(binaryOwned.vbaProjectProperties().description, std::string("XL++ generated VBA project"),
                    "Project metadata can be inspected from externally owned vbaProject.bin");
    bool binaryPropertiesRejected = false;
    try { binaryOwned.setVbaProjectProperties(projectProperties); }
    catch (const std::runtime_error&) { binaryPropertiesRejected = true; }
    test.checkTrue(binaryPropertiesRejected, "Externally supplied VBA project rejects destructive metadata rewrite");

    // Digital signatures are intentionally preserve-only. Any source/metadata
    // rewrite would invalidate the signature, so XL++ refuses the mutation.
    xlpp::Workbook signedGenerated;
    signedGenerated.addWorksheet("Signed Host");
    signedGenerated.setVbaModuleText("SignedModule", "Sub SignedCode(): End Sub");
    xlpp::PreservedPart signaturePart;
    signaturePart.name = "xl/vbaProjectSignature.bin";
    signaturePart.data = "synthetic-signature-fixture";
    signaturePart.overrideType = "application/vnd.ms-office.vbaProjectSignature";
    signaturePart.extension = "bin";
    signaturePart.defaultType = "application/vnd.ms-office.vbaProjectSignature";
    signaturePart.compress = false;
    signedGenerated.preservedParts().push_back(std::move(signaturePart));
    test.checkTrue(signedGenerated.hasVbaSignature(), "VBA signature part is detected");
    test.checkTrue(!signedGenerated.vbaSourceEditable(), "Signed VBA project is not advertised as source-editable");
    bool signedEditRejected = false;
    try { signedGenerated.setVbaModuleText("SignedModule", "Sub Changed(): End Sub"); }
    catch (const std::runtime_error&) { signedEditRejected = true; }
    test.checkTrue(signedEditRejected, "Source rewrite is rejected while a VBA signature is present");

    bool invalidRejected = false;
    try { workbook.setVbaModuleText("1 Bad Name", "Sub X(): End Sub"); }
    catch (const std::invalid_argument&) { invalidRejected = true; }
    test.checkTrue(invalidRejected, "Invalid VBA identifier is rejected");

    // Exercise regular CFB sectors, multiple OVBA chunks, an empty source stream,
    // and a larger directory red-black tree rather than only mini-stream modules.
    const auto stressPath = std::filesystem::temp_directory_path() / "xlpp_vba_text_stress.xlsm";
    std::string longSource = "Option Explicit\nPublic Sub LongGeneratedMacro()\n";
    for (int line = 0; line < 700; ++line)
        longSource += "    ' generated source line " + std::to_string(line) + " for CFB and OVBA coverage\n";
    longSource += "End Sub\n";
    xlpp::Workbook stress;
    stress.addWorksheet("MacroHost");
    stress.setVbaModuleText("EmptyModule", "");
    stress.setVbaModuleText("ModuleA", "Public Const A As Long = 1\n");
    stress.setVbaModuleText("ModuleB", "Public Const B As Long = 2\n");
    stress.setVbaModuleText("ModuleC", "Public Const C As Long = 3\n");
    stress.setVbaModuleText("ModuleD", "Public Const D As Long = 4\n");
    stress.setVbaModuleText("LongModule", longSource);
    stress.save(stressPath);
    auto stressZip = xlpp::internal::ZipArchive::open(stressPath);
    test.checkTrue(stressZip.get("xl/vbaProject.bin").size() > 8192,
                   "Large text module uses a multi-sector CFB project");
    xlpp::Workbook stressLoaded;
    stressLoaded.load(stressPath);
    test.checkEqual(stressLoaded.vbaModules().size(), std::size_t{8},
                    "Multiple standard modules plus Sheet1 and ThisWorkbook are read");
    test.checkEqual(stressLoaded.vbaModuleText("EmptyModule").value(), std::string{},
                    "Empty VBA source stream round-trips");
    test.checkTrue(stressLoaded.vbaModuleText("LongModule").value() ==
                       xlpp::internal::normalizeVbaSource(longSource),
                   "Long multi-chunk VBA source round-trips exactly");

    std::filesystem::remove(firstPath);
    std::filesystem::remove(secondPath);
    std::filesystem::remove(sheetMutationPath);
    std::filesystem::remove(removedPath);
    std::filesystem::remove(stressPath);
    std::filesystem::remove(advancedPath);
    std::filesystem::remove(topologyPath);
    std::filesystem::remove(exportedProjectPath);
}

void testExternalCellAndStyleReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_cells.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Imported\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<numFmts count=\"0\"/>"
        "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
        "<font><b/><i/><sz val=\"14\"/><color rgb=\"FF112233\"/><name val=\"Arial\"/></font></fonts>"
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFC000\"/><bgColor indexed=\"64\"/></patternFill></fill></fills>"
        "<borders count=\"2\"><border><left/><right/><top/><bottom/><diagonal/></border>"
        "<border><left style=\"thin\"><color rgb=\"FF00B050\"/></left><right style=\"double\"><color rgb=\"FF0070C0\"/></right><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"1\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" wrapText=\"1\" textRotation=\"30\"/></xf>"
        "<xf numFmtId=\"14\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\" applyNumberFormat=\"1\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "<dxfs count=\"0\"/><tableStyles count=\"0\" defaultTableStyle=\"TableStyleMedium2\" defaultPivotStyle=\"PivotStyleLight16\"/>"
        "</styleSheet>";
    const std::string sharedStrings =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"2\" uniqueCount=\"2\">"
        "<si><t>Shared value</t></si>"
        "<si><r><rPr><b/><color rgb=\"FFFF0000\"/><sz val=\"12\"/><rFont val=\"Calibri\"/></rPr><t>Rich</t></r><r><rPr><i/></rPr><t xml:space=\"preserve\"> text</t></r></si>"
        "</sst>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<dimension ref=\"A1:I1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/>"
        "<sheetData><row r=\"1\">"
        "<c r=\"A1\" t=\"s\"><v>0</v></c>"
        "<c r=\"B1\" t=\"inlineStr\"><is><t>Inline value</t></is></c>"
        "<c r=\"C1\"><v>42.5</v></c>"
        "<c r=\"D1\" t=\"b\"><v>1</v></c>"
        "<c r=\"E1\" t=\"e\"><v>#DIV/0!</v></c>"
        "<c r=\"F1\"><f>SUM(C1,7.5)</f><v>50</v></c>"
        "<c r=\"G1\" t=\"s\"><v>1</v></c>"
        "<c r=\"H1\" s=\"2\"><v>45292</v></c>"
        "<c r=\"I1\" s=\"1\" t=\"inlineStr\"><is><t>Styled</t></is></c>"
        "</row></sheetData></worksheet>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, sharedStrings);

    xlpp::Workbook workbook;
    workbook.load(path);
    auto* sheet = workbook.worksheet("Imported");
    test.checkTrue(sheet != nullptr, "Handcrafted worksheet is discovered by relationship reader");
    test.checkEqual(std::get<std::string>(sheet->cell("A1").value()), std::string("Shared value"), "Shared string is read from external fixture");
    test.checkEqual(std::get<std::string>(sheet->cell("B1").value()), std::string("Inline value"), "Inline string is read");
    test.checkNear(std::get<double>(sheet->cell("C1").value()), 42.5, 1e-12, "Numeric cell is read");
    test.checkTrue(std::get<bool>(sheet->cell("D1").value()), "Boolean cell is read");
    test.checkEqual(static_cast<unsigned>(*sheet->cell("E1").error()), static_cast<unsigned>(xlpp::CellError::DivisionByZero), "Error cell is read");
    test.checkEqual(sheet->cell("F1").formula(), std::string("SUM(C1,7.5)"), "Formula text is read");
    test.checkNear(std::get<double>(sheet->cell("F1").value()), 50.0, 1e-12, "Formula cached value is read");
    test.checkTrue(sheet->cell("G1").hasRichText(), "Shared rich-text cell is preserved as runs");
    test.checkEqual(sheet->cell("G1").richTextValue()->runs().size(), std::size_t{2}, "Two rich-text runs are read");
    test.checkTrue(sheet->cell("G1").richTextValue()->runs()[0].bold(), "Rich-text bold property is read");
    test.checkEqual(sheet->cell("G1").richTextValue()->runs()[0].color(), std::string("FFFF0000"), "Rich-text color is read");
    test.checkTrue(sheet->cell("H1").isDate(), "Built-in date style converts numeric serial to DateTime");
    test.checkEqual(*sheet->cell("H1").date(), xlpp::DateTime{2024, 1, 1}, "Date serial is decoded");
    const auto& styled = sheet->cell("I1");
    test.checkTrue(styled.font().bold(), "External style font bold is read");
    test.checkTrue(styled.font().italic(), "External style font italic is read");
    test.checkEqual(styled.font().name(), std::string("Arial"), "External font name is read");
    test.checkNear(styled.font().size(), 14.0, 1e-12, "External font size is read");
    test.checkEqual(styled.font().color().argb(), std::string("FF112233"), "External font color is read");
    test.checkEqual(styled.fill().foregroundColor().argb(), std::string("FFFFC000"), "External fill color is read");
    test.checkEqual(styled.border().left().style(), std::string("thin"), "External left border is read");
    test.checkEqual(styled.border().right().style(), std::string("double"), "External right border is read");
    test.checkEqual(styled.alignment().horizontal(), std::string("center"), "External horizontal alignment is read");
    test.checkTrue(styled.alignment().wrapText(), "External wrap-text alignment is read");
    test.checkEqual(styled.alignment().textRotation(), 30, "External text rotation is read");
    std::filesystem::remove(path);
}

void testExternalWorksheetFeatureReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_features.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheets><sheet name=\"Features\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>"
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs><cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>"
        "<dxfs count=\"1\"><dxf><font><color rgb=\"FFFF0000\"/><b/></font></dxf></dxfs></styleSheet>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<sheetPr><tabColor rgb=\"FF00B0F0\"/><pageSetUpPr fitToPage=\"1\"/></sheetPr><dimension ref=\"A1:D9\"/>"
        "<sheetViews><sheetView workbookViewId=\"2\" zoomScale=\"135\" zoomScaleNormal=\"110\" showGridLines=\"0\" rightToLeft=\"1\" showOutlineSymbols=\"0\"><pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"split\"/></sheetView></sheetViews>"
        "<sheetFormatPr defaultRowHeight=\"15\"/><cols><col min=\"2\" max=\"2\" width=\"22.5\" customWidth=\"1\" hidden=\"1\" bestFit=\"1\" outlineLevel=\"2\" collapsed=\"1\"/></cols>"
        "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Category</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Value</t></is></c></row>"
        "<row r=\"3\" ht=\"27\" customHeight=\"1\" hidden=\"1\" outlineLevel=\"2\" collapsed=\"1\"><c r=\"A3\" t=\"inlineStr\"><is><t>A</t></is></c><c r=\"B3\"><v>12</v></c></row></sheetData>"
        "<sheetProtection sheet=\"1\" password=\"DAA7\" formatCells=\"1\" autoFilter=\"1\"/>"
        "<mergeCells count=\"1\"><mergeCell ref=\"C2:D2\"/></mergeCells>"
        "<autoFilter ref=\"A1:B9\"><filterColumn colId=\"0\"><filters blank=\"1\"><filter val=\"A\"/><filter val=\"B\"/></filters></filterColumn><sortState ref=\"A1:B9\" caseSensitive=\"1\"><sortCondition ref=\"B2:B9\" descending=\"1\"/></sortState></autoFilter>"
        "<conditionalFormatting sqref=\"B2:B9\"><cfRule type=\"cellIs\" dxfId=\"0\" priority=\"1\" operator=\"greaterThan\" stopIfTrue=\"1\"><formula>10</formula></cfRule></conditionalFormatting>"
        "<dataValidations count=\"1\"><dataValidation type=\"list\" errorStyle=\"warning\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" promptTitle=\"Pick\" prompt=\"Choose A or B\" errorTitle=\"Invalid\" error=\"Not allowed\" sqref=\"A2:A9\"><formula1>\"A,B\"</formula1></dataValidation></dataValidations>"
        "<hyperlinks><hyperlink ref=\"A3\" r:id=\"rIdExternal\" display=\"External\" tooltip=\"Open site\"/><hyperlink ref=\"B3\" location=\"Features!A1\" display=\"Internal\"/></hyperlinks>"
        "<printOptions horizontalCentered=\"1\" verticalCentered=\"1\" headings=\"1\" gridLines=\"1\"/><pageMargins left=\"0.4\" right=\"0.5\" top=\"0.6\" bottom=\"0.7\" header=\"0.2\" footer=\"0.3\"/>"
        "<pageSetup orientation=\"landscape\" paperSize=\"9\" fitToWidth=\"1\" fitToHeight=\"2\" blackAndWhite=\"1\" draft=\"1\" firstPageNumber=\"3\" useFirstPageNumber=\"1\"/>"
        "<headerFooter differentOddEven=\"1\" differentFirst=\"1\"><oddHeader>&amp;CExternal Header</oddHeader><oddFooter>&amp;RPage &amp;P</oddFooter></headerFooter><legacyDrawing r:id=\"rIdVml\"/></worksheet>";
    const std::string rels =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rIdExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://example.com/read-fixture\" TargetMode=\"External\"/>"
        "<Relationship Id=\"rIdComments\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments1.xml\"/>"
        "<Relationship Id=\"rIdVml\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/commentsDrawing1.vml\"/>"
        "</Relationships>";
    const std::string comments =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><authors><author>External Author</author></authors><commentList><comment ref=\"D4\" authorId=\"0\"><text><r><t>Read </t></r><r><t>comment</t></r></text></comment></commentList></comments>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, {}, rels, comments);

    xlpp::Workbook workbook;
    workbook.load(path);
    auto* sheet = workbook.worksheet("Features");
    test.checkTrue(sheet != nullptr, "External feature fixture sheet loads");
    test.checkTrue(sheet->tryColumnDimension(2) != nullptr, "External column dimension is read");
    test.checkNear(*sheet->tryColumnDimension(2)->width, 22.5, 1e-12, "External column width is read");
    test.checkTrue(sheet->tryColumnDimension(2)->hidden, "External hidden column flag is read");
    test.checkTrue(sheet->tryRowDimension(3) != nullptr, "External row dimension is read");
    test.checkNear(*sheet->tryRowDimension(3)->height, 27.0, 1e-12, "External row height is read");
    test.checkTrue(sheet->tryRowDimension(3)->hidden, "External hidden row flag is read");
    test.checkTrue(sheet->isMerged("C2"), "External merged range is read");
    test.checkEqual(sheet->sheetView().zoomScale(), 135, "External zoom scale is read");
    test.checkEqual(sheet->sheetView().zoomScaleNormal(), 110, "External normal zoom is read");
    test.checkTrue(!sheet->sheetView().showGridLines(), "External gridline visibility is read");
    test.checkTrue(sheet->sheetView().rightToLeft(), "External right-to-left view is read");
    test.checkEqual(sheet->sheetView().topLeftCell(), std::string("B3"), "External split-pane top-left cell is read");
    test.checkEqual(sheet->sheetView().pane(), std::string("bottomRight"), "External active pane is read");
    test.checkTrue(sheet->protection().enabled(), "External worksheet protection is read");
    test.checkEqual(sheet->protection().passwordHash(), std::string("DAA7"), "External worksheet password hash is read");
    test.checkTrue(sheet->autoFilter().enabled(), "External AutoFilter is read");
    test.checkEqual(sheet->autoFilter().reference(), std::string("A1:B9"), "External AutoFilter reference is read");
    test.checkEqual(sheet->autoFilter().columns().at(0).values().size(), std::size_t{2}, "External filter values are read");
    test.checkTrue(sheet->autoFilter().columns().at(0).includeBlank(), "External filter blank option is read");
    test.checkTrue(sheet->autoFilter().sortStateValue().has_value(), "External sort state is read");
    test.checkTrue(sheet->autoFilter().sortStateValue()->conditions().front().descending, "External descending sort is read");
    test.checkEqual(sheet->conditionalFormatting().entries().size(), std::size_t{1}, "External conditional formatting is read");
    const auto& rule = sheet->conditionalFormatting().entries().front().rules().front();
    test.checkEqual(static_cast<unsigned>(rule.type()), static_cast<unsigned>(xlpp::ConditionalRuleType::CellIs), "External cell-is rule type is read");
    test.checkEqual(rule.formulas().front(), std::string("10"), "External conditional formula is read");
    test.checkTrue(rule.hasDifferentialStyle(), "External differential style is linked");
    test.checkEqual(rule.differentialStyle().font().color().argb(), std::string("FFFF0000"), "External differential font color is read");
    test.checkEqual(sheet->dataValidations().items().size(), std::size_t{1}, "External data validation is read");
    const auto& validation = sheet->dataValidations().items().front();
    test.checkEqual(static_cast<unsigned>(validation.type()), static_cast<unsigned>(xlpp::DataValidationType::List), "External validation type is read");
    test.checkEqual(validation.formula1(), std::string("\"A,B\""), "External validation formula is read");
    test.checkEqual(validation.promptTitle(), std::string("Pick"), "External validation prompt title is read");
    test.checkTrue(sheet->cell("A3").hasHyperlink(), "External hyperlink relationship is read");
    test.checkEqual(sheet->cell("A3").hyperlinkValue()->target(), std::string("https://example.com/read-fixture"), "External hyperlink target is resolved");
    test.checkTrue(sheet->cell("B3").hasHyperlink(), "Internal hyperlink is read without relationship");
    test.checkTrue(sheet->cell("D4").hasComment(), "External legacy comment is read");
    test.checkEqual(sheet->cell("D4").commentValue()->text(), std::string("Read comment"), "External rich comment runs are concatenated");
    test.checkEqual(sheet->cell("D4").commentValue()->author(), std::string("External Author"), "External comment author is read");
    test.checkEqual(static_cast<unsigned>(sheet->pageSetup().orientation()), static_cast<unsigned>(xlpp::PageOrientation::Landscape), "External page orientation is read");
    test.checkEqual(static_cast<unsigned>(sheet->pageSetup().paperSize()), static_cast<unsigned>(xlpp::PaperSize::A4), "External paper size is read");
    test.checkEqual(sheet->pageSetup().fitToWidth(), 1u, "External fit-to-width is read");
    test.checkEqual(sheet->pageSetup().fitToHeight(), 2u, "External fit-to-height is read");
    test.checkNear(sheet->pageMargins().left(), 0.4, 1e-12, "External left page margin is read");
    test.checkTrue(sheet->printOptions().horizontalCentered(), "External horizontal print centering is read");
    test.checkEqual(sheet->headerFooter().oddHeader(), std::string("&CExternal Header"), "External header text is read");
    std::filesystem::remove(path);
}

void testExternalWorkbookMetadataReaderFixture(TestContext& test) {
    const auto path = std::filesystem::temp_directory_path() / "xlpp_external_reader_metadata.xlsx";
    const std::string workbookXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<workbookPr date1904=\"1\"/><workbookProtection lockStructure=\"1\" lockWindows=\"1\" workbookPassword=\"83AF\"/>"
        "<sheets><sheet name=\"Metadata\" sheetId=\"1\" r:id=\"rIdSheet1\"/></sheets>"
        "<definedNames><definedName name=\"InputValue\" comment=\"external name\">Metadata!$A$1</definedName><definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">Metadata!$A$1:$D$20</definedName><definedName name=\"_xlnm.Print_Titles\" localSheetId=\"0\">Metadata!$A:$B,Metadata!$1:$2</definedName></definedNames>"
        "<calcPr calcId=\"777\" calcMode=\"manual\" fullPrecision=\"0\" iterate=\"1\" iterateCount=\"55\" iterateDelta=\"0.0005\" fullCalcOnLoad=\"1\" calcOnSave=\"1\"/></workbook>";
    const std::string stylesXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><fonts count=\"1\"><font/></fonts><fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills><borders count=\"1\"><border/></borders><cellStyleXfs count=\"1\"><xf/></cellStyleXfs><cellXfs count=\"1\"><xf/></cellXfs></styleSheet>";
    const std::string sheetXml =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
    const std::string core =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\"><dc:title>External title</dc:title><dc:subject>Read fixture</dc:subject><dc:creator>Fixture Author</dc:creator><dc:description>Loaded without XLPP writer</dc:description><cp:keywords>read,test</cp:keywords><cp:category>QA</cp:category><cp:lastModifiedBy>External Tool</cp:lastModifiedBy></cp:coreProperties>";
    const std::string custom =
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/custom-properties\" xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\"><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"2\" name=\"StringProp\"><vt:lpwstr>hello</vt:lpwstr></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"3\" name=\"IntProp\"><vt:i4>42</vt:i4></property><property fmtid=\"{D5CDD505-2E9C-101B-9397-08002B2CF9AE}\" pid=\"4\" name=\"BoolProp\"><vt:bool>true</vt:bool></property></Properties>";
    writeExternalReaderFixture(path, workbookXml, sheetXml, stylesXml, {}, {}, {}, core, custom);

    xlpp::Workbook workbook;
    workbook.load(path);
    test.checkTrue(workbook.date1904(), "External 1904 date system is read");
    test.checkEqual(workbook.properties().title(), std::string("External title"), "External document title is read");
    test.checkEqual(workbook.properties().creator(), std::string("Fixture Author"), "External document creator is read");
    test.checkEqual(workbook.properties().lastModifiedBy(), std::string("External Tool"), "External lastModifiedBy is read");
    test.checkEqual(workbook.customProperties().items().size(), std::size_t{3}, "External custom properties are read");
    test.checkEqual(workbook.customProperties().items()[0].value(), std::string("hello"), "External custom string property is read");
    test.checkEqual(workbook.customProperties().items()[1].type(), std::string("i4"), "External custom integer type is retained");
    test.checkTrue(workbook.protection().lockStructure(), "External workbook structure protection is read");
    test.checkTrue(workbook.protection().lockWindows(), "External workbook window protection is read");
    test.checkEqual(workbook.protection().workbookPasswordHash(), std::string("83AF"), "External workbook password hash is read");
    test.checkEqual(workbook.calcProperties().calcId(), 777, "External calcId is read");
    test.checkEqual(workbook.calcProperties().calcMode(), std::string("manual"), "External calculation mode is read");
    test.checkTrue(workbook.calcProperties().iterate(), "External iterative calculation flag is read");
    test.checkEqual(workbook.calcProperties().iterateCount(), 55, "External iteration count is read");
    test.checkNear(workbook.calcProperties().iterateDelta(), 0.0005, 1e-12, "External iteration delta is read");
    test.checkTrue(workbook.calcProperties().fullCalcOnLoad(), "External full-calc-on-load flag is read");
    test.checkTrue(workbook.calcProperties().calcOnSave(), "External calc-on-save flag is read");
    test.checkTrue(!workbook.calcProperties().fullPrecision(), "External full-precision false is read");
    test.checkTrue(workbook.definedName("InputValue") != nullptr, "External user-defined name is read");
    test.checkEqual(workbook.definedName("InputValue")->comment(), std::string("external name"), "External defined-name comment is read");
    const auto* sheet = workbook.worksheet("Metadata");
    test.checkEqual(sheet->printArea(), std::string("A1:D20"), "External built-in print area is applied to worksheet");
    test.checkEqual(sheet->printTitlesCols(), std::string("A:B"), "External print-title columns are applied");
    test.checkEqual(sheet->printTitlesRows(), std::string("1:2"), "External print-title rows are applied");
    std::filesystem::remove(path);
}

