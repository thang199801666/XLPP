#pragma once

#include <cstddef>
#include <initializer_list>
#include <iterator>
#include <memory>
#include <stdexcept>
#include <type_traits>
#include <utility>
#include <vector>

namespace xlpp {

template <class T>
class StableVector {
    using Storage = std::vector<std::unique_ptr<T>>;

    template <bool IsConst>
    class IteratorImpl {
        using BaseIterator = std::conditional_t<IsConst, typename Storage::const_iterator, typename Storage::iterator>;
        BaseIterator it_{};

        explicit IteratorImpl(BaseIterator it) : it_(it) {}
        friend class StableVector;
        template <bool> friend class IteratorImpl;

    public:
        using iterator_category = std::random_access_iterator_tag;
        using iterator_concept = std::random_access_iterator_tag;
        using value_type = T;
        using difference_type = std::ptrdiff_t;
        using reference = std::conditional_t<IsConst, const T&, T&>;
        using pointer = std::conditional_t<IsConst, const T*, T*>;

        IteratorImpl() = default;

        template <bool B = IsConst, class = std::enable_if_t<B>>
        IteratorImpl(const IteratorImpl<false>& other) : it_(other.it_) {}

        reference operator*() const noexcept { return **it_; }
        pointer operator->() const noexcept { return it_->get(); }
        reference operator[](difference_type n) const noexcept { return **(it_ + n); }

        IteratorImpl& operator++() noexcept { ++it_; return *this; }
        IteratorImpl operator++(int) noexcept { auto copy = *this; ++*this; return copy; }
        IteratorImpl& operator--() noexcept { --it_; return *this; }
        IteratorImpl operator--(int) noexcept { auto copy = *this; --*this; return copy; }
        IteratorImpl& operator+=(difference_type n) noexcept { it_ += n; return *this; }
        IteratorImpl& operator-=(difference_type n) noexcept { it_ -= n; return *this; }

        friend IteratorImpl operator+(IteratorImpl it, difference_type n) noexcept { it += n; return it; }
        friend IteratorImpl operator+(difference_type n, IteratorImpl it) noexcept { it += n; return it; }
        friend IteratorImpl operator-(IteratorImpl it, difference_type n) noexcept { it -= n; return it; }
        friend difference_type operator-(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return lhs.it_ - rhs.it_; }

        friend bool operator==(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return lhs.it_ == rhs.it_; }
        friend bool operator!=(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return !(lhs == rhs); }
        friend bool operator<(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return lhs.it_ < rhs.it_; }
        friend bool operator>(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return rhs < lhs; }
        friend bool operator<=(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return !(rhs < lhs); }
        friend bool operator>=(const IteratorImpl& lhs, const IteratorImpl& rhs) noexcept { return !(lhs < rhs); }
    };

public:
    using value_type = T;
    using size_type = std::size_t;
    using difference_type = std::ptrdiff_t;
    using reference = T&;
    using const_reference = const T&;
    using pointer = T*;
    using const_pointer = const T*;
    using iterator = IteratorImpl<false>;
    using const_iterator = IteratorImpl<true>;

    StableVector() = default;

    StableVector(std::initializer_list<T> values) {
        storage_.reserve(values.size());
        for (const auto& value : values) push_back(value);
    }

    StableVector(const StableVector& other) {
        storage_.reserve(other.size());
        for (const auto& value : other) push_back(value);
    }

    StableVector(StableVector&&) noexcept = default;

    StableVector& operator=(const StableVector& other) {
        if (this == &other) return *this;
        StableVector copy(other);
        swap(copy);
        return *this;
    }

    StableVector& operator=(StableVector&&) noexcept = default;

    void swap(StableVector& other) noexcept { storage_.swap(other.storage_); }

    [[nodiscard]] bool empty() const noexcept { return storage_.empty(); }
    [[nodiscard]] size_type size() const noexcept { return storage_.size(); }
    [[nodiscard]] size_type capacity() const noexcept { return storage_.capacity(); }
    void reserve(size_type count) { storage_.reserve(count); }

    reference operator[](size_type index) noexcept { return *storage_[index]; }
    const_reference operator[](size_type index) const noexcept { return *storage_[index]; }

    reference at(size_type index) {
        if (index >= size()) throw std::out_of_range("StableVector index out of range");
        return *storage_[index];
    }
    const_reference at(size_type index) const {
        if (index >= size()) throw std::out_of_range("StableVector index out of range");
        return *storage_[index];
    }

    reference front() noexcept { return *storage_.front(); }
    const_reference front() const noexcept { return *storage_.front(); }
    reference back() noexcept { return *storage_.back(); }
    const_reference back() const noexcept { return *storage_.back(); }

    iterator begin() noexcept { return iterator(storage_.begin()); }
    const_iterator begin() const noexcept { return const_iterator(storage_.begin()); }
    const_iterator cbegin() const noexcept { return const_iterator(storage_.cbegin()); }
    iterator end() noexcept { return iterator(storage_.end()); }
    const_iterator end() const noexcept { return const_iterator(storage_.end()); }
    const_iterator cend() const noexcept { return const_iterator(storage_.cend()); }

    void clear() noexcept { storage_.clear(); }

    void push_back(const T& value) { storage_.push_back(std::make_unique<T>(value)); }
    void push_back(T&& value) { storage_.push_back(std::make_unique<T>(std::move(value))); }

    template <class... Args>
    reference emplace_back(Args&&... args) {
        storage_.push_back(std::make_unique<T>(std::forward<Args>(args)...));
        return *storage_.back();
    }

    iterator erase(iterator position) { return iterator(storage_.erase(position.it_)); }
    iterator erase(const_iterator position) { return iterator(storage_.erase(position.it_)); }
    iterator erase(iterator first, iterator last) { return iterator(storage_.erase(first.it_, last.it_)); }
    iterator erase(const_iterator first, const_iterator last) { return iterator(storage_.erase(first.it_, last.it_)); }

private:
    Storage storage_;
};

template <class T>
void swap(StableVector<T>& lhs, StableVector<T>& rhs) noexcept { lhs.swap(rhs); }

} // namespace xlpp
