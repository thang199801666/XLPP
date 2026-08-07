#pragma once
#include <XLPP/Protection/LegacyPassword.h>
#include <string>
#include <string_view>
namespace xlpp {
class WorkbookProtection {
public:
    bool lockStructure() const noexcept { return lockStructure_; }
    void setLockStructure(bool v) noexcept { lockStructure_ = v; }
    bool lockWindows() const noexcept { return lockWindows_; }
    void setLockWindows(bool v) noexcept { lockWindows_ = v; }
    bool lockRevision() const noexcept { return lockRevision_; }
    void setLockRevision(bool v) noexcept { lockRevision_ = v; }
    const std::string& workbookPasswordHash() const noexcept { return workbookPasswordHash_; }
    void setWorkbookPasswordHash(std::string v) { workbookPasswordHash_ = std::move(v); }
    bool hasPassword() const noexcept { return !workbookPasswordHash_.empty(); }
    void setPassword(std::string_view password) { workbookPasswordHash_ = legacyProtectionPasswordHash(password); }
    void clearPassword() noexcept { workbookPasswordHash_.clear(); }
private:
    bool lockStructure_{false}, lockWindows_{false}, lockRevision_{false};
    std::string workbookPasswordHash_;
};
}
