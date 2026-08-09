#include <XLPP/Workbook/Workbook.h>
#include "Package/Xml/XmlUtilities.h"
#include "Package/Zip/ZipArchive.h"
#include "Package/Opc/RelationshipGraph.h"
#include "Encryption/OfficeEncryption.h"
#include "IO/FileTransaction.h"
#include "IO/BinaryFile.h"
#include "Preservation/PartPolicy.h"
#include "OOXML/Common/Namespaces.h"
#include "OOXML/Common/PackageRelationships.h"
#include "OOXML/Styles/StyleCodec.h"
#include "OOXML/Common/RichTextCodec.h"
#include "OOXML/Worksheet/WorksheetReader.h"
#include "OOXML/Pivot/PivotCodec.h"
#include "OOXML/Workbook/WorkbookReferenceSupport.h"
#include "VBA/VbaProjectBinary.h"
#include <algorithm>
#include <array>
#include <filesystem>
#include <fstream>
#include <string>
#include <unordered_map>
#include <vector>

using xlpp::internal::ooxml::StyleCatalog;
using xlpp::internal::ooxml::parseStyleCatalog;
using xlpp::internal::ooxml::parseDifferentialStyles;
using xlpp::internal::ooxml::LoadedSharedString;
using xlpp::internal::ooxml::parseRichTextRuns;
using xlpp::internal::ooxml::localPrintReference;
using xlpp::internal::preservation::isRegeneratedPart;
using xlpp::internal::preservation::extensionOf;

namespace xlpp {
void Workbook::loadInPlace(const std::filesystem::path& p, const LoadOptions& options) {
    // Encrypted OOXML uses a CFB outer container rather than a ZIP signature.
    // Decrypt to a temporary package and then reuse the normal hardened ZIP
    // loading pipeline, including its entry/size/cancellation limits.
    {
        std::ifstream input(p, std::ios::binary);
        if (!input) throw std::runtime_error("Cannot open workbook: " + p.string());
        std::array<unsigned char, 8> signature{};
        input.read(reinterpret_cast<char*>(signature.data()), static_cast<std::streamsize>(signature.size()));
        if (internal::hasCompoundFileSignature(signature.data(), static_cast<std::size_t>(input.gcount()))) {
            if (options.maxFileBytes != 0) {
                std::error_code error;
                const auto size = std::filesystem::file_size(p, error);
                if (!error && size > options.maxFileBytes)
                    throw std::runtime_error("Encrypted workbook exceeds maxFileBytes");
            }
            if (options.password.empty())
                throw std::runtime_error("Workbook is encrypted; LoadOptions::password is required");
            const auto compound = xlpp::internal::readBinaryFile(p);
            const auto encryption = internal::inspectOfficeEncryptionBytes(compound);
            if (!encryption.encrypted || !encryption.supported)
                throw std::runtime_error("Unsupported OLE/Office encryption container");
            const auto package = internal::decryptOfficePackage(
                compound, options.password, options.verifyEncryptionIntegrity);
            if (options.maxFileBytes != 0 && package.size() > options.maxFileBytes)
                throw std::runtime_error("Decrypted workbook exceeds maxFileBytes");
            const auto temporary = internal::xlppTemporaryPath("decrypt");
            try {
                xlpp::internal::writeBinaryFile(temporary, package);
                auto plainOptions = options;
                plainOptions.password.clear();
                loadInPlace(temporary, plainOptions);
                std::filesystem::remove(temporary);
            } catch (...) {
                std::error_code ignored;
                std::filesystem::remove(temporary, ignored);
                throw;
            }
            return;
        }
    }
    clear(); diagnostics_ = LoadDiagnostics{}; internal::ZipOpenLimits limits; limits.maxEntries = options.maxEntries; limits.maxEntryBytes = options.maxEntryBytes; limits.maxTotalBytes = options.maxTotalBytes; limits.maxFileBytes = options.maxFileBytes; limits.cancel = options.cancel; limits.progress = options.progress; auto z = internal::ZipArchive::open(p, limits);
    const auto relationshipGraph = internal::RelationshipGraph::fromArchive(z);
    preservedRelationships_ = relationshipGraph.relationships();
    const auto relationshipValidation = relationshipGraph.validate();
    for (const auto& issue : relationshipValidation.relationshipSyntaxErrors)
        diagnostics_.warnings.push_back("Relationship syntax issue: " + issue);
    for (const auto& issue : relationshipValidation.duplicateRelationshipIds)
        diagnostics_.warnings.push_back("Duplicate relationship: " + issue);
    for (const auto& issue : relationshipValidation.danglingRelationships)
        diagnostics_.warnings.push_back("Dangling relationship: " + issue);
    for (const auto& issue : relationshipValidation.orphanedParts)
        diagnostics_.warnings.push_back("Orphaned package part: " + issue);
    for (const auto& issue : relationshipValidation.contentTypeErrors)
        diagnostics_.warnings.push_back("Content-type issue: " + issue);
    for (const auto& issue : relationshipValidation.ownerReferenceErrors)
        diagnostics_.warnings.push_back("Broken owner reference: " + issue);
    customPropertiesPartPresent_ = z.contains("docProps/custom.xml");
    if(z.contains("docProps/core.xml")){auto cp=z.get("docProps/core.xml");properties_.setTitle(internal::tagText(cp,"dc:title"));properties_.setSubject(internal::tagText(cp,"dc:subject"));properties_.setCreator(internal::tagText(cp,"dc:creator"));properties_.setDescription(internal::tagText(cp,"dc:description"));properties_.setKeywords(internal::tagText(cp,"cp:keywords"));properties_.setCategory(internal::tagText(cp,"cp:category"));properties_.setLastModifiedBy(internal::tagText(cp,"cp:lastModifiedBy"));}if(z.contains("docProps/custom.xml")){auto cust=z.get("docProps/custom.xml");for(const auto& propertyNode:internal::tags(cust,"property")){const auto n=internal::attribute(propertyNode,"name");if(n.empty())continue;const auto vtText=internal::tagText(propertyNode,"vt:lpwstr");if(!vtText.empty()){customProps_.add(CustomProperty(std::string(n),vtText));continue;}if(const auto i4Text=internal::tagText(propertyNode,"vt:i4");!i4Text.empty()){customProps_.add(CustomProperty(std::string(n),std::stoi(i4Text)));continue;}if(const auto r8Text=internal::tagText(propertyNode,"vt:r8");!r8Text.empty()){customProps_.add(CustomProperty(std::string(n),std::stod(r8Text)));continue;}if(const auto boolText=internal::tagText(propertyNode,"vt:bool");!boolText.empty()){customProps_.add(CustomProperty(std::string(n),boolText=="true"));continue;}customProps_.add(CustomProperty(std::string(n),vtText));}}StyleCatalog styleCatalog;std::vector<Style> dxfStyles;if(z.contains("xl/styles.xml")){const auto stylesText=z.get("xl/styles.xml");styleCatalog=parseStyleCatalog(stylesText);dxfStyles=parseDifferentialStyles(stylesText);for(const auto& node:internal::tags(stylesText,"cellStyle")){const auto name=internal::attribute(node,"name");if(name.empty()||name=="Normal")continue;const auto xf=internal::attribute(node,"xfId");if(xf.empty())continue;const auto id=static_cast<std::size_t>(std::stoul(xf));if(id<styleCatalog.items.size())namedStyles_.emplace_back(name,styleCatalog.items[id]);}}std::vector<LoadedSharedString> shared;if(z.contains("xl/sharedStrings.xml")){const auto sstXml = z.get("xl/sharedStrings.xml"); for(auto&si:internal::tags(sstXml,"si")){LoadedSharedString item; item.richText=parseRichTextRuns(si); item.plainText=item.richText?item.richText->plainText():internal::tagText(si,"t"); shared.push_back(std::move(item));}}auto wb=z.get("xl/workbook.xml");
sourceWorkbookXml_ = wb;
strictNamespaces_ = wb.find("http://purl.oclc.org/ooxml/spreadsheetml/main") != std::string::npos;
for(const auto& protectionNode:internal::tags(wb,"workbookProtection")){
    protection_.setLockStructure(internal::attribute(protectionNode,"lockStructure")=="1");
    protection_.setLockWindows(internal::attribute(protectionNode,"lockWindows")=="1");
    protection_.setLockRevision(internal::attribute(protectionNode,"lockRevision")=="1");
    protection_.setWorkbookPasswordHash(internal::attribute(protectionNode,"workbookPassword"));
}
for(const auto& prNode:internal::tags(wb,"workbookPr"))
    date1904_ = internal::attribute(prNode,"date1904")=="1";
for(const auto& calcNode:internal::tags(wb,"calcPr")){
    const auto cm = internal::attribute(calcNode,"calcMode"); if(!cm.empty()) calcProps_.setCalcMode(std::string(cm));
    const auto ci = internal::attribute(calcNode,"calcId"); if(!ci.empty()) calcProps_.setCalcId(static_cast<int>(std::stoul(ci)));
    calcProps_.setCalcOnSave(internal::attribute(calcNode,"calcOnSave") == "1");
    calcProps_.setFullCalcOnLoad(internal::attribute(calcNode,"fullCalcOnLoad") == "1");
    calcProps_.setFullPrecision(internal::attribute(calcNode,"fullPrecision") != "0");
    calcProps_.setIterate(internal::attribute(calcNode,"iterate") == "1");
    const auto ic = internal::attribute(calcNode,"iterateCount"); if(!ic.empty()) calcProps_.setIterateCount(static_cast<int>(std::stoul(ic)));
    const auto id = internal::attribute(calcNode,"iterateDelta"); if(!id.empty()) calcProps_.setIterateDelta(std::stod(id));
}
struct PendingPrintName { std::string name; std::size_t sheetIndex; std::string value; };
std::vector<PendingPrintName> pendingPrintNames;
for(const auto& node:internal::tags(wb,"definedName")){
    const auto name = internal::attribute(node,"name");
    const auto value = internal::tagText(node,"definedName");
    const auto local=internal::attribute(node,"localSheetId");
    if ((name == "_xlnm.Print_Area" || name == "_xlnm.Print_Titles") && !local.empty()) {
        pendingPrintNames.push_back({name, static_cast<std::size_t>(std::stoul(local)), value});
        continue;
    }
    DefinedName item(name,value);
    if(!local.empty()) item.setLocalSheetId(static_cast<std::size_t>(std::stoul(local)));
    item.setHidden(internal::attribute(node,"hidden")=="1");
    item.setComment(internal::attribute(node,"comment"));
    definedNames_.push_back(std::move(item));
}
auto relxml=z.get("xl/_rels/workbook.xml.rels");
std::unordered_map<std::string,std::string> targets;
for(auto&r:internal::tags(relxml,"Relationship"))
    targets[internal::attribute(r,"Id")]=internal::attribute(r,"Target");
for(auto&s:internal::tags(wb,"sheet")) {
    const auto name = internal::attribute(s,"name");
    try {
        const auto rid = internal::attribute(s,"r:id");
        auto targetIt = targets.find(rid);
        if (targetIt == targets.end())
            throw std::runtime_error("worksheet relationship '" + rid + "' was not found");
        const auto target = internal::RelationshipGraph::resolveTarget("xl/workbook.xml", targetIt->second);
        if (target.empty() || !z.contains(target))
            throw std::runtime_error("worksheet part was not found: " + target);
        auto& ws=addWorksheet(name);
        sourceSheetNames_.push_back(name);
        sourceSheetParts_.push_back(target);
        sourceSheetXml_.push_back(z.get(target));
        xlpp::internal::ooxml::loadWorksheetModel(ws, sourceSheetXml_.back(), z, target, styleCatalog, dxfStyles, shared, date1904_);
    } catch (const std::exception& e) {
        diagnostics_.errors.push_back("Sheet '" + name + "': " + e.what());
        if (!options.lenient) throw;
        continue;
    }
}
    // Build a semantic read model for imported pivots while preserving their
    // original package parts byte-for-byte until the caller explicitly edits them.
    for (std::size_t i = 0; i < sheets_.size() && i < sourceSheetParts_.size(); ++i)
        xlpp::internal::ooxml::loadImportedPivotModels(sheets_[i], z, sourceSheetParts_[i], preservedRelationships_);

    // Apply built-in print names after all worksheets have been created.
    for (const auto& pending : pendingPrintNames) {
        if (pending.sheetIndex >= sheets_.size()) continue;
        auto& sheet = sheets_[pending.sheetIndex];
        if (pending.name == "_xlnm.Print_Area") {
            sheet.setPrintArea(localPrintReference(pending.value));
            continue;
        }
        std::string rows;
        std::string cols;
        const auto localValue = localPrintReference(pending.value);
        std::size_t areaStart = 0;
        while (areaStart <= localValue.size()) {
            const auto comma = localValue.find(',', areaStart);
            const auto area = localValue.substr(areaStart, comma == std::string::npos ? std::string::npos : comma - areaStart);
            const auto first = area.find_first_not_of(' ');
            if (first != std::string::npos && std::isdigit(static_cast<unsigned char>(area[first]))) {
                if (!rows.empty()) rows += ',';
                rows += area;
            } else if (!area.empty()) {
                if (!cols.empty()) cols += ',';
                cols += area;
            }
            if (comma == std::string::npos) break;
            areaStart = comma + 1;
        }
        sheet.setPrintTitlesRows(rows);
        sheet.setPrintTitlesCols(cols);
    }
    preservedParts_.clear();
    std::unordered_map<std::string, std::string> overrides, defaults;
    if (z.contains("[Content_Types].xml")) {
        const auto& ctText = z.get("[Content_Types].xml");
        for (const auto& node : internal::tags(ctText, "Override")) {
            const auto part = internal::attribute(node, "PartName");
            const auto ct = internal::attribute(node, "ContentType");
            if (!part.empty() && !ct.empty()) overrides[part] = ct;
        }
        for (const auto& node : internal::tags(ctText, "Default")) {
            const auto ext = internal::attribute(node, "Extension");
            const auto ct = internal::attribute(node, "ContentType");
            if (!ext.empty() && !ct.empty()) defaults[ext] = ct;
        }
    }
    for (const auto& partName : z.entryNames()) {
        if (isRegeneratedPart(partName)) continue;
        PreservedPart part;
        part.name = partName;
        part.data = z.get(partName);
        const auto overrideIt = overrides.find("/" + partName);
        if (overrideIt != overrides.end()) {
            part.overrideType = overrideIt->second;
        } else {
            const auto ext = extensionOf(partName);
            const auto defaultIt = defaults.find(ext);
            if (defaultIt != defaults.end()) { part.extension = ext; part.defaultType = defaultIt->second; }
        }
        part.compress = !(part.extension == "png" || part.extension == "jpg" || part.extension == "jpeg"
            || part.extension == "gif" || part.extension == "bmp" || part.extension == "bin");
        preservedParts_.push_back(std::move(part));
    }
    // Source-generated XL++ projects carry a stable project ID/profile. Mark
    // them so a later sheet insertion/removal followed by save() rebuilds the
    // document-module list. Arbitrary externally authored projects remain
    // preserved verbatim.
    if (const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
            return part.name == "xl/vbaProject.bin";
        }); it != preservedParts_.end()) {
        const std::vector<unsigned char> bytes(it->data.begin(), it->data.end());
        generatedVbaProject_ = internal::isXlppGeneratedVbaProjectBinary(bytes);
    }
    // After a successful load, mark all sheets clean so the next save can
    // benefit from differential caching.
    for (auto& s : sheets_) s.clearDirty();
}
}

