#pragma once

#include <memory>
#include <optional>
#include <type_traits>
#include <utility>

namespace xlpp::internal {

// Copyable lazy optional storage for model payloads that are absent in the
// overwhelming majority of objects. The empty state occupies one pointer;
// engaged values allocate an std::optional<T> so public APIs that expose
// `const std::optional<T>&` can retain their exact signatures and semantics.
template <class T>
class CompactOptional {
public:
    CompactOptional() noexcept = default;

    CompactOptional(const CompactOptional& other)
        : value_(other.value_ ? std::make_unique<std::optional<T>>(*other.value_) : nullptr) {}

    CompactOptional& operator=(const CompactOptional& other) {
        if (this == &other) return *this;
        if (!other.value_) {
            value_.reset();
        } else if (value_) {
            *value_ = *other.value_;
        } else {
            value_ = std::make_unique<std::optional<T>>(*other.value_);
        }
        return *this;
    }

    CompactOptional(CompactOptional&&) noexcept = default;
    CompactOptional& operator=(CompactOptional&&) noexcept = default;

    bool has_value() const noexcept { return value_ && value_->has_value(); }
    explicit operator bool() const noexcept { return has_value(); }

    T& value() {
        if (!value_) throw std::bad_optional_access();
        return value_->value();
    }
    const T& value() const {
        if (!value_) throw std::bad_optional_access();
        return value_->value();
    }

    T& operator*() { return **value_; }
    const T& operator*() const { return **value_; }
    T* operator->() { return std::addressof(**value_); }
    const T* operator->() const { return std::addressof(**value_); }

    template <class... Args>
    T& emplace(Args&&... args) {
        if (!value_) value_ = std::make_unique<std::optional<T>>();
        return value_->emplace(std::forward<Args>(args)...);
    }

    void reset() noexcept { value_.reset(); }

    void set(std::optional<T> value) {
        if (!value) {
            value_.reset();
            return;
        }
        if (!value_) value_ = std::make_unique<std::optional<T>>();
        *value_ = std::move(value);
    }

    void set(T value) {
        if (!value_) value_ = std::make_unique<std::optional<T>>();
        *value_ = std::move(value);
    }

    const std::optional<T>& optionalRef() const noexcept {
        static const std::optional<T> empty;
        return value_ ? *value_ : empty;
    }

private:
    std::unique_ptr<std::optional<T>> value_;
};

} // namespace xlpp::internal
