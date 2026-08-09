#pragma once
#include <atomic>
#include <condition_variable>
#include <cstddef>
#include <functional>
#include <future>
#include <memory>
#include <mutex>
#include <queue>
#include <thread>
#include <type_traits>
#include <vector>

namespace xlpp::internal {

class ThreadPool {
public:
    explicit ThreadPool(std::size_t numThreads)
        : stop_(false)
    {
        for (std::size_t i = 0; i < numThreads; ++i) {
            workers_.emplace_back([this] {
                for (;;) {
                    std::function<void()> task;
                    {
                        std::unique_lock<std::mutex> lock(queueMutex_);
                        condition_.wait(lock, [this] { return stop_ || !tasks_.empty(); });
                        if (stop_ && tasks_.empty()) return;
                        task = std::move(tasks_.front());
                        tasks_.pop();
                    }
                    task();
                }
            });
        }
    }

    ~ThreadPool() {
        {
            std::lock_guard<std::mutex> lock(queueMutex_);
            stop_ = true;
        }
        condition_.notify_all();
        for (auto& worker : workers_)
            if (worker.joinable()) worker.join();
    }

    ThreadPool(const ThreadPool&) = delete;
    ThreadPool& operator=(const ThreadPool&) = delete;
    ThreadPool(ThreadPool&&) = delete;
    ThreadPool& operator=(ThreadPool&&) = delete;

    std::size_t size() const noexcept { return workers_.size(); }

    template<typename F, typename... Args>
    auto enqueue(F&& f, Args&&... args) -> std::future<std::invoke_result_t<F, Args...>> {
        using returnType = std::invoke_result_t<F, Args...>;
        auto task = std::make_shared<std::packaged_task<returnType()>>(
            std::bind(std::forward<F>(f), std::forward<Args>(args)...));
        std::future<returnType> result = task->get_future();
        {
            std::lock_guard<std::mutex> lock(queueMutex_);
            if (stop_) throw std::runtime_error("enqueue on stopped ThreadPool");
            tasks_.emplace([task] { (*task)(); });
        }
        condition_.notify_one();
        return result;
    }

    // Distribute range [begin, end) across workers, calling fn(i) for each i.
    // Deterministic: workers process contiguous chunks like the old static-chunking.
    template<typename F>
    void parallelFor(std::size_t begin, std::size_t end, F&& fn) {
        if (begin >= end) return;
        const std::size_t count = end - begin;
        const std::size_t numWorkers = workers_.size();
        const std::size_t chunk = (count + numWorkers - 1) / numWorkers;
        std::vector<std::future<void>> futures;
        futures.reserve(numWorkers);
        for (std::size_t w = 0; w < numWorkers; ++w) {
            const std::size_t b = begin + w * chunk;
            if (b >= end) break;
            const std::size_t e = std::min(b + chunk, end);
            futures.push_back(enqueue([b, e, &fn] {
                for (std::size_t i = b; i < e; ++i) fn(i);
            }));
        }
        for (auto& f : futures) f.get();
    }

private:
    std::vector<std::thread> workers_;
    std::queue<std::function<void()>> tasks_;
    std::mutex queueMutex_;
    std::condition_variable condition_;
    bool stop_;
};

} // namespace xlpp::internal
