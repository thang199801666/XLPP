#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

enum class ModelValidationSeverity { Warning, Error };

struct ModelValidationIssue {
    ModelValidationSeverity severity{ModelValidationSeverity::Error};
    std::string code;
    std::string worksheetName;
    std::string objectId;
    std::string message;
};

struct WorkbookModelValidationReport {
    std::vector<ModelValidationIssue> issues;

    std::size_t errorCount() const noexcept;
    std::size_t warningCount() const noexcept;
    bool ok() const noexcept { return errorCount() == 0; }
};

} // namespace xlpp
