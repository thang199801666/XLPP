#pragma once
#include <memory>
#include <new>
#include <exception>
#include <string>
#include <utility>

namespace xlpp {
class Hyperlink {
public:
    Hyperlink() = default;
    explicit Hyperlink(std::string target) { ensureData().target = std::move(target); }

    Hyperlink(const Hyperlink& other)
        : data_(other.data_ ? std::make_unique<Data>(*other.data_) : nullptr) {}
    Hyperlink& operator=(const Hyperlink& other) {
        if (this == &other) return *this;
        if (!other.data_) data_.reset();
        else if (data_) *data_ = *other.data_;
        else data_ = std::make_unique<Data>(*other.data_);
        return *this;
    }
    Hyperlink(Hyperlink&&) noexcept = default;
    Hyperlink& operator=(Hyperlink&&) noexcept = default;

    const std::string& target() const noexcept { return data_ ? data_->target : emptyText(); }
    void setTarget(std::string v) { ensureData().target = std::move(v); }
    const std::string& display() const noexcept { return data_ ? data_->display : emptyText(); }
    void setDisplay(std::string v) { ensureData().display = std::move(v); }
    const std::string& tooltip() const noexcept { return data_ ? data_->tooltip : emptyText(); }
    void setTooltip(std::string v) { ensureData().tooltip = std::move(v); }
    bool external() const noexcept { return data_ ? data_->external : true; }
    void setExternal(bool v) noexcept {
        if (!data_) {
            if (v) return; // default state needs no allocation
            data_.reset(new (std::nothrow) Data());
            if (!data_) std::terminate();
        }
        data_->external = v;
    }

private:
    struct Data {
        std::string target;
        std::string display;
        std::string tooltip;
        bool external{true};
    };

    static const std::string& emptyText() noexcept { static const std::string value; return value; }
    Data& ensureData() {
        if (!data_) data_ = std::make_unique<Data>();
        return *data_;
    }

    std::unique_ptr<Data> data_;
};
}
