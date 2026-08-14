#pragma once

#include <memory>
#include <utility>

namespace xlpp::internal {

// Copyable lazy storage for value objects whose default state dominates model
// density.  The empty representation is one pointer.  A value is allocated
// only on first mutation, while const callers can read a process-wide default
// object without allocating per owner.
template <class T>
class CompactValue {
public:
    CompactValue() noexcept = default;

    CompactValue(const CompactValue& other)
        : value_(other.value_ ? std::make_unique<T>(*other.value_) : nullptr) {}

    CompactValue& operator=(const CompactValue& other) {
        if (this == &other) return *this;
        if (!other.value_) value_.reset();
        else if (value_) *value_ = *other.value_;
        else value_ = std::make_unique<T>(*other.value_);
        return *this;
    }

    CompactValue(CompactValue&&) noexcept = default;
    CompactValue& operator=(CompactValue&&) noexcept = default;

    bool has_value() const noexcept { return static_cast<bool>(value_); }
    explicit operator bool() const noexcept { return has_value(); }

    T& value() { return *value_; }
    const T& value() const { return *value_; }
    T* operator->() noexcept { return value_.get(); }
    const T* operator->() const noexcept { return value_.get(); }
    T& operator*() noexcept { return *value_; }
    const T& operator*() const noexcept { return *value_; }

    template <class... Args>
    T& getOrCreate(Args&&... args) {
        if (!value_) value_ = std::make_unique<T>(std::forward<Args>(args)...);
        return *value_;
    }

    void set(T value) {
        if (value_) *value_ = std::move(value);
        else value_ = std::make_unique<T>(std::move(value));
    }

    void reset() noexcept { value_.reset(); }

    const T& getOr(const T& fallback) const noexcept {
        return value_ ? *value_ : fallback;
    }

private:
    std::unique_ptr<T> value_;
};

} // namespace xlpp::internal
