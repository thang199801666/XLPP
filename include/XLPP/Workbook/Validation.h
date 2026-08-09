#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

enum class WorkbookValidationSeverity { Warning, Error };

struct WorkbookValidationIssue {
    WorkbookValidationSeverity severity{WorkbookValidationSeverity::Error};
    std::string code;
    std::string message;
    std::string worksheet;
};

struct WorkbookValidationOptions {
    bool validateWorksheetNames{true};
    bool validateDefinedNames{true};
    bool validateTables{true};
    bool validatePivots{true};
    bool validateCharts{true};
    bool validateVba{true};
};

struct WorkbookValidationReport {
    std::vector<WorkbookValidationIssue> issues;
    std::size_t errorCount{0};
    std::size_t warningCount{0};

    bool ok() const noexcept { return errorCount == 0; }
    explicit operator bool() const noexcept { return ok(); }
};

} // namespace xlpp
