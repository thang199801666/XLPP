#pragma once
#include <cstddef>
#include <string>
#include <vector>

namespace xlpp {

struct ExternalWorkbookLinkInfo {
    std::string partName;
    std::vector<std::string> sheetNames;
    std::vector<std::string> definedNames;
};

struct WorkbookConnectionInfo {
    std::string partName;
    std::string id;
    std::string name;
    std::string description;
    std::string type;
    bool refreshOnLoad{false};
    bool background{false};
    bool deleted{false};
};

struct QueryTableInfo {
    std::string partName;
    std::string name;
    std::string connectionId;
    bool refreshOnLoad{false};
};

struct ExternalDataInspection {
    std::vector<ExternalWorkbookLinkInfo> externalWorkbooks;
    std::vector<WorkbookConnectionInfo> connections;
    std::vector<QueryTableInfo> queryTables;
    std::vector<std::string> powerQueryParts;
    std::vector<std::string> webQueryParts;
    std::vector<std::string> unknownConnectionParts;

    bool hasExternalWorkbooks() const noexcept { return !externalWorkbooks.empty(); }
    bool hasConnections() const noexcept { return !connections.empty(); }
    bool hasQueryTables() const noexcept { return !queryTables.empty(); }
    bool hasPowerQuery() const noexcept { return !powerQueryParts.empty(); }
};

} // namespace xlpp
