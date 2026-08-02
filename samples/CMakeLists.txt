# Sample Programs

add_executable(XLPP_Sample
    XLPP.Sample/main.cpp
)

target_include_directories(XLPP_Sample PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(XLPP_Sample PRIVATE xlpp_static)

# Native benchmark
add_executable(XLPP_Benchmark
    ${CMAKE_CURRENT_SOURCE_DIR}/../benchmarks/bench_native.cpp
)

target_include_directories(XLPP_Benchmark PRIVATE
    ${CMAKE_CURRENT_SOURCE_DIR}/../src/XLPP
)

target_link_libraries(XLPP_Benchmark PRIVATE xlpp_static)
