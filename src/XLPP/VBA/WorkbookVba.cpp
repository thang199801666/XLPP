#include <XLPP/Workbook/Workbook.h>
#include "VBA/VbaProjectBinary.h"
#include <algorithm>
#include <cctype>
#include <fstream>
#include <iterator>
#include <stdexcept>
#include <utility>

namespace xlpp {
namespace {
std::string folded(std::string value) {
    std::transform(value.begin(), value.end(), value.begin(), [](unsigned char ch) {
        return static_cast<char>(std::tolower(ch));
    });
    return value;
}

bool sameVbaName(const std::string& lhs, const std::string& rhs) {
    return folded(lhs) == folded(rhs);
}

std::vector<std::string> worksheetCodeNames(const std::deque<Worksheet>& sheets) {
    std::vector<std::string> result;
    result.reserve(sheets.size());
    for (std::size_t index = 0; index < sheets.size(); ++index) {
        if (!sheets[index].vbaCodeName().empty()) result.push_back(sheets[index].vbaCodeName());
        else result.push_back("Sheet" + std::to_string(index + 1));
    }
    return result;
}

bool hasAuthoredVbaContent(const std::vector<VbaModule>& modules) {
    return std::any_of(modules.begin(), modules.end(), [](const VbaModule& module) {
        return module.type != VbaModuleType::Document || !module.source.empty();
    });
}

bool hasCustomProjectProperties(const VbaProjectProperties& properties) {
    return properties.name != "VBAProject" || !properties.description.empty() || !properties.helpFile.empty()
        || properties.helpContextId != 0 || !properties.constants.empty();
}
} // namespace

void Workbook::addVbaProject(const std::filesystem::path& path) {
    std::ifstream stream(path, std::ios::binary);
    if (!stream) throw std::runtime_error("Cannot open VBA project file: " + path.string());
    std::vector<unsigned char> bytes((std::istreambuf_iterator<char>(stream)), std::istreambuf_iterator<char>());
    if (bytes.empty()) throw std::invalid_argument("VBA project file is empty");
    setVbaProject(std::move(bytes));
}

void Workbook::setVbaProject(std::vector<unsigned char> bytes) {
    if (bytes.empty()) throw std::invalid_argument("VBA project bytes cannot be empty");
    removeVbaProject();
    generatedVbaProject_ = false;
    PreservedPart part;
    part.name = "xl/vbaProject.bin";
    part.data.assign(reinterpret_cast<const char*>(bytes.data()), bytes.size());
    part.overrideType = "application/vnd.ms-office.vbaProject";
    part.extension = "bin";
    part.defaultType = "application/vnd.ms-office.vbaProject";
    part.compress = false;
    preservedParts_.push_back(std::move(part));
}

bool Workbook::hasVbaProject() const noexcept {
    return std::any_of(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
}

bool Workbook::hasVbaSignature() const noexcept {
    return std::any_of(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProjectSignature.bin"
            || part.name == "xl/vbaProjectSignatureAgile.bin"
            || part.name == "xl/vbaProjectSignatureV3.bin";
    });
}

bool Workbook::removeVbaProject() noexcept {
    const auto oldSize = preservedParts_.size();
    preservedParts_.erase(std::remove_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin"
            || part.name == "xl/vbaProjectSignature.bin"
            || part.name == "xl/vbaProjectSignatureAgile.bin"
            || part.name == "xl/vbaProjectSignatureV3.bin"
            || part.name == "xl/_rels/vbaProject.bin.rels";
    }), preservedParts_.end());
    const bool removed = preservedParts_.size() != oldSize;
    if (removed) generatedVbaProject_ = false;
    return removed;
}

std::vector<unsigned char> Workbook::vbaProjectBytes() const {
    const auto it = std::find_if(preservedParts_.begin(), preservedParts_.end(), [](const PreservedPart& part) {
        return part.name == "xl/vbaProject.bin";
    });
    if (it == preservedParts_.end()) return {};
    return std::vector<unsigned char>(it->data.begin(), it->data.end());
}

void Workbook::saveVbaProject(const std::filesystem::path& path) const {
    const auto bytes = vbaProjectBytes();
    if (bytes.empty()) throw std::runtime_error("Workbook has no VBA project");
    std::ofstream stream(path, std::ios::binary | std::ios::trunc);
    if (!stream) throw std::runtime_error("Cannot create VBA project file: " + path.string());
    stream.write(reinterpret_cast<const char*>(bytes.data()), static_cast<std::streamsize>(bytes.size()));
    if (!stream) throw std::runtime_error("Cannot write VBA project file: " + path.string());
}

VbaProjectProperties Workbook::vbaProjectProperties() const {
    const auto bytes = vbaProjectBytes();
    return bytes.empty() ? VbaProjectProperties{} : internal::readVbaProjectProperties(bytes);
}

void Workbook::setVbaProjectProperties(VbaProjectProperties properties) {
    if (hasVbaSignature())
        throw std::runtime_error("Cannot rewrite a digitally signed VBA project; remove/re-sign the project explicitly before source or metadata changes");
    if (hasVbaProject() && !generatedVbaProject_)
        throw std::runtime_error("Cannot rewrite properties inside an externally supplied vbaProject.bin without discarding unsupported project metadata, references, forms, or signatures");
    std::vector<VbaModule> modules;
    if (hasVbaProject()) modules = vbaModules();
    const auto bytes = internal::buildVbaProjectBinary(modules, worksheetCodeNames(sheets_), properties);
    setVbaProject(bytes);
    generatedVbaProject_ = true;
}

std::vector<VbaModule> Workbook::vbaModules() const {
    const auto bytes = vbaProjectBytes();
    if (bytes.empty()) return {};
    auto parsed = internal::readVbaProjectBinary(bytes);
    if (!generatedVbaProject_) return parsed;

    // Present the effective host topology immediately, even before the next
    // save rebuilds vbaProject.bin. Exact code-name matching preserves event
    // code across worksheet deletion/reordering. Positional fallback handles a
    // deliberate code-name rename when no exact document module remains.
    std::vector<VbaModule> worksheetDocuments;
    std::optional<VbaModule> workbookDocument;
    std::vector<VbaModule> userModules;
    for (auto& module : parsed) {
        if (module.type != VbaModuleType::Document) {
            userModules.push_back(std::move(module));
        } else if (sameVbaName(module.name, "ThisWorkbook")) {
            workbookDocument = std::move(module);
        } else {
            worksheetDocuments.push_back(std::move(module));
        }
    }

    std::vector<VbaModule> effective;
    effective.reserve(sheets_.size() + userModules.size() + 1);
    std::vector<bool> consumed(worksheetDocuments.size(), false);
    for (std::size_t sheetIndex = 0; sheetIndex < sheets_.size(); ++sheetIndex) {
        const auto codeName = sheets_[sheetIndex].vbaCodeName().empty()
            ? "Sheet" + std::to_string(sheetIndex + 1)
            : sheets_[sheetIndex].vbaCodeName();
        auto match = worksheetDocuments.size();
        for (std::size_t i = 0; i < worksheetDocuments.size(); ++i) {
            if (!consumed[i] && sameVbaName(worksheetDocuments[i].name, codeName)) { match = i; break; }
        }
        if (match == worksheetDocuments.size() && sheetIndex < worksheetDocuments.size() && !consumed[sheetIndex])
            match = sheetIndex;
        if (match < worksheetDocuments.size()) {
            auto module = std::move(worksheetDocuments[match]);
            consumed[match] = true;
            module.name = codeName;
            effective.push_back(std::move(module));
        } else {
            effective.push_back({codeName, "", VbaModuleType::Document});
        }
    }
    if (workbookDocument) {
        workbookDocument->name = "ThisWorkbook";
        effective.push_back(std::move(*workbookDocument));
    } else {
        effective.push_back({"ThisWorkbook", "", VbaModuleType::Document});
    }
    for (auto& module : userModules) effective.push_back(std::move(module));
    return effective;
}

std::optional<std::string> Workbook::vbaModuleText(const std::string& moduleName) const {
    for (const auto& module : vbaModules())
        if (sameVbaName(module.name, moduleName)) return module.source;
    return std::nullopt;
}

void Workbook::setVbaModule(VbaModule module) {
    internal::validateVbaModuleName(module.name);
    if (hasVbaSignature())
        throw std::runtime_error("Cannot rewrite a digitally signed VBA project; source changes would invalidate its signature");
    if (module.type != VbaModuleType::Document && sameVbaName(module.name, "ThisWorkbook"))
        throw std::invalid_argument("ThisWorkbook is reserved for the workbook document module");
    if (hasVbaProject() && !generatedVbaProject_)
        throw std::runtime_error("Cannot rewrite source inside an externally supplied vbaProject.bin without discarding unsupported project metadata, references, forms, or signatures");

    const auto properties = hasVbaProject() ? vbaProjectProperties() : VbaProjectProperties{};
    std::vector<VbaModule> modules;
    if (hasVbaProject()) modules = vbaModules();
    module.source = internal::normalizeVbaSource(std::move(module.source));

    if (module.type == VbaModuleType::Document) {
        bool hostExists = sameVbaName(module.name, "ThisWorkbook");
        for (const auto& sheet : sheets_)
            hostExists = hostExists || sameVbaName(module.name, sheet.vbaCodeName());
        if (!hostExists)
            throw std::invalid_argument("VBA document module must be ThisWorkbook or an existing worksheet code name: " + module.name);
    }

    const auto it = std::find_if(modules.begin(), modules.end(), [&](const VbaModule& current) {
        return sameVbaName(current.name, module.name);
    });
    if (it == modules.end()) modules.push_back(std::move(module));
    else *it = std::move(module);

    const auto bytes = internal::buildVbaProjectBinary(modules, worksheetCodeNames(sheets_), properties);
    setVbaProject(bytes);
    generatedVbaProject_ = true;
}

void Workbook::setVbaModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Standard});
}

void Workbook::setVbaClassModuleText(std::string moduleName, std::string source,
                                     bool readOnly, bool privateModule) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Class, readOnly, privateModule});
}

void Workbook::setVbaDocumentModuleText(std::string moduleName, std::string source) {
    setVbaModule({std::move(moduleName), std::move(source), VbaModuleType::Document});
}

bool Workbook::removeVbaModule(const std::string& moduleName) {
    if (!hasVbaProject()) return false;
    if (hasVbaSignature())
        throw std::runtime_error("Cannot rewrite a digitally signed VBA project; removing a module would invalidate its signature");
    if (!generatedVbaProject_)
        throw std::runtime_error("Cannot remove a module from an externally supplied vbaProject.bin without discarding unsupported project metadata, references, forms, or signatures");

    const auto properties = vbaProjectProperties();
    auto modules = vbaModules();
    const auto it = std::find_if(modules.begin(), modules.end(), [&](const VbaModule& module) {
        return sameVbaName(module.name, moduleName);
    });
    if (it == modules.end()) return false;

    if (it->type == VbaModuleType::Document) {
        it->source.clear();
        it->readOnly = false;
        it->privateModule = false;
    } else {
        modules.erase(it);
    }

    if (!hasAuthoredVbaContent(modules) && !hasCustomProjectProperties(properties)) {
        removeVbaProject();
        return true;
    }

    const auto bytes = internal::buildVbaProjectBinary(modules, worksheetCodeNames(sheets_), properties);
    setVbaProject(bytes);
    generatedVbaProject_ = true;
    return true;
}

} // namespace xlpp
