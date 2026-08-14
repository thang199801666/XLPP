#pragma once

#include <memory>
#include <string>
#include <utility>

namespace xlpp::internal {

// A compact, copyable string holder intended for model fields whose default
// value is overwhelmingly common.  The holder occupies one pointer when the
// field is at its default and allocates a std::string only after the value is
// changed.  Public model APIs continue to expose const std::string&.
class CompactString {
public:
    CompactString() noexcept = default;

    CompactString(const CompactString& other)
        : value_(other.value_ ? std::make_unique<std::string>(*other.value_) : nullptr) {}

    CompactString& operator=(const CompactString& other) {
        if (this == &other) return *this;
        if (!other.value_) {
            value_.reset();
        } else if (value_) {
            *value_ = *other.value_;
        } else {
            value_ = std::make_unique<std::string>(*other.value_);
        }
        return *this;
    }

    CompactString(CompactString&&) noexcept = default;
    CompactString& operator=(CompactString&&) noexcept = default;

    const std::string& get(const std::string& fallback) const noexcept {
        return value_ ? *value_ : fallback;
    }

    void set(std::string value, const std::string& fallback) {
        if (value == fallback) {
            value_.reset();
            return;
        }
        if (value_) *value_ = std::move(value);
        else value_ = std::make_unique<std::string>(std::move(value));
    }

    bool isDefault() const noexcept { return !value_; }

    friend bool operator==(const CompactString& lhs, const CompactString& rhs) noexcept {
        if (!lhs.value_ || !rhs.value_) return !lhs.value_ && !rhs.value_;
        return *lhs.value_ == *rhs.value_;
    }

private:
    std::unique_ptr<std::string> value_;
};

} // namespace xlpp::internal
