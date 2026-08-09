if(NOT CMAKE_CXX_COMPILER_ID MATCHES "Clang")
    message(WARNING "XLPP_BUILD_FUZZERS requires Clang/libFuzzer; no fuzz target was created")
    return()
endif()

add_executable(XLPP_WorkbookLoadFuzzer WorkbookLoadFuzzer.cpp)
target_link_libraries(XLPP_WorkbookLoadFuzzer PRIVATE xlpp_static)
target_compile_options(XLPP_WorkbookLoadFuzzer PRIVATE
    -fsanitize=fuzzer,address,undefined
    -fno-omit-frame-pointer)
target_link_options(XLPP_WorkbookLoadFuzzer PRIVATE
    -fsanitize=fuzzer,address,undefined
    -fno-omit-frame-pointer)

# Keep the harness usable from CI without checking generated corpora into the
# source tree. Example:
#   XLPP_WorkbookLoadFuzzer corpus/ -max_total_time=60 -artifact_prefix=artifacts/
