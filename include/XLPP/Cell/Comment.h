#pragma once
#include <memory>
#include <string>
#include <utility>

namespace xlpp {
class Comment {
public:
    Comment() = default;
    Comment(std::string text, std::string author) {
        auto& data = ensureData();
        data.text = std::move(text);
        data.author = std::move(author);
    }

    Comment(const Comment& other)
        : data_(other.data_ ? std::make_unique<Data>(*other.data_) : nullptr) {}
    Comment& operator=(const Comment& other) {
        if (this == &other) return *this;
        if (!other.data_) data_.reset();
        else if (data_) *data_ = *other.data_;
        else data_ = std::make_unique<Data>(*other.data_);
        return *this;
    }
    Comment(Comment&&) noexcept = default;
    Comment& operator=(Comment&&) noexcept = default;

    const std::string& text() const noexcept { return data_ ? data_->text : emptyText(); }
    void setText(std::string v) { ensureData().text = std::move(v); }
    const std::string& author() const noexcept { return data_ ? data_->author : emptyText(); }
    void setAuthor(std::string v) { ensureData().author = std::move(v); }

private:
    struct Data { std::string text, author; };
    static const std::string& emptyText() noexcept { static const std::string value; return value; }
    Data& ensureData() {
        if (!data_) data_ = std::make_unique<Data>();
        return *data_;
    }
    std::unique_ptr<Data> data_;
};
}
