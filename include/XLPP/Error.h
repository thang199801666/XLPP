#pragma once
#include <stdexcept>
#include <string>
#include <utility>

namespace xlpp {

enum class ErrorCode {
    Unknown,
    InvalidArgument,
    InvalidState,
    Io,
    Package,
    Xml,
    Validation,
    Encryption,
    Formula,
    Unsupported,
    Cancelled,
    ResourceLimit
};

class Exception : public std::runtime_error {
public:
    Exception(ErrorCode code, std::string message)
        : std::runtime_error(std::move(message)), code_(code) {}
    ErrorCode code() const noexcept { return code_; }
private:
    ErrorCode code_{ErrorCode::Unknown};
};

} // namespace xlpp
