#pragma once
#include <cstddef>
#include <cstdint>
#include <memory>
#include <new>
#include <vector>

namespace xlpp::internal {

// Bump-pointer arena allocator. Pre-allocates fixed-size blocks and hands out
// monotonically increasing pointers. Individual allocations are never freed;
// the whole arena is released on destruction (or reset). Ideal for the save
// path where many short-lived temporaries are created.
class Arena {
public:
    explicit Arena(std::size_t blockSize = 1 << 20) : blockSize_(blockSize) {}

    ~Arena() = default;

    Arena(const Arena&) = delete;
    Arena& operator=(const Arena&) = delete;

    void* allocate(std::size_t bytes, std::size_t align = alignof(std::max_align_t)) {
        if (bytes == 0) bytes = 1;
        const auto aligned = [](std::size_t v, std::size_t a) {
            return (v + a - 1) & ~(a - 1);
        };
        const std::size_t off = aligned(offset_, align);
        if (off + bytes > currentSize_) {
            // Allocate a new block
            const std::size_t allocSize = (std::max)(blockSize_, bytes);
            blocks_.emplace_back(allocSize);
            current_ = blocks_.back().get();
            currentSize_ = allocSize;
            offset_ = 0;
            return allocate(bytes, align);
        }
        void* result = current_ + off;
        offset_ = off + bytes;
        return result;
    }

    void reset() {
        blocks_.clear();
        current_ = nullptr;
        currentSize_ = 0;
        offset_ = 0;
    }

private:
    struct Block {
        explicit Block(std::size_t size) : ptr(::operator new(size, std::align_val_t(64))), size(size) {}
        ~Block() { ::operator delete(ptr, std::align_val_t(64)); }
        void* ptr;
        std::size_t size;
    };

    std::size_t blockSize_;
    std::vector<Block> blocks_;
    char* current_{nullptr};
    std::size_t currentSize_{0};
    std::size_t offset_{0};
};

} // namespace xlpp::internal
