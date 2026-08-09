cmake_minimum_required(VERSION 3.18)
project(xlpp_consumer_subdir CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# This smoke validates source-tree integration; unity compilation reduces repeated
# header parse cost without changing production/release build settings.
set(CMAKE_UNITY_BUILD ON)
set(CMAKE_UNITY_BUILD_BATCH_SIZE 8)

# This target verifies source-tree integration, not optimized/debug-symbol output.
# Avoid paying full Debug-info cost while still compiling XLPP from scratch.
if(MSVC)
    add_compile_options(/Od)
elseif(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    add_compile_options(-O0 -g0)
endif()

# Path to the XL++ source tree (passed by the test driver).
set(XLPP_SOURCE_DIR "" CACHE PATH "Path to the XL++ source directory")

set(XLPP_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(XLPP_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(XLPP_BUILD_CAPI OFF CACHE BOOL "" FORCE)

add_subdirectory(${XLPP_SOURCE_DIR} xlpp)
# Bundled zlib C sources are not unity-safe (private enums/macros collide).
if(TARGET zlibstatic)
    set_target_properties(zlibstatic PROPERTIES UNITY_BUILD OFF)
endif()
if(TARGET zlib)
    set_target_properties(zlib PROPERTIES UNITY_BUILD OFF)
endif()
# ZIP implementation TUs intentionally use repeated anonymous-namespace
# signature names, so keep that one module out of unity mode.
if(TARGET xlpp_package_zip_obj)
    set_target_properties(xlpp_package_zip_obj PROPERTIES UNITY_BUILD OFF)
endif()

add_executable(subdir_app main.cpp)
target_link_libraries(subdir_app PRIVATE XLPP::xlpp)

add_custom_target(run_subdir_app COMMAND subdir_app DEPENDS subdir_app)
