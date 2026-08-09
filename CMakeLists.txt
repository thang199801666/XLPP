cmake_minimum_required(VERSION 3.18)
project(XLPP
    VERSION 1.12.0
    DESCRIPTION "High-performance C++20 Excel xlsx read/write library"
    LANGUAGES CXX C
)

# Must be included before any target uses $<INSTALL_INTERFACE:...>.
include(GNUInstallDirs)

# === Options ===
option(XLPP_BUILD_TESTS "Build unit tests" ON)
option(XLPP_BUILD_SAMPLES "Build sample programs" ON)
option(XLPP_BUILD_CAPI "Build C API DLL" ON)
option(XLPP_BUILD_BINDINGS "Build language bindings" OFF)
option(XLPP_BUILD_BENCHMARKS "Build external library benchmarks" OFF)
option(XLPP_BUILD_TOOLS "Build XLPP package inspection tools" ON)
option(XLPP_USE_BUNDLED_ZLIB "Use bundled zlib instead of system" ON)
option(XLPP_ENABLE_SANITIZERS "Enable Address/UndefinedBehavior sanitizers for development builds" OFF)
option(XLPP_ENABLE_STRICT_WARNINGS "Enable strong compiler warnings on XLPP core" OFF)
option(XLPP_BUILD_FUZZERS "Build libFuzzer hardening targets (Clang only)" OFF)

# === C++ Standard ===
set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# === Position independent code ===
# Required so the static libraries (XLPP + bundled zlib) can be linked into
# shared objects such as the Python extension module.
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# === Development hardening ===
if(XLPP_ENABLE_SANITIZERS)
    if(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
        add_compile_options(-fsanitize=address,undefined -fno-omit-frame-pointer)
        add_link_options(-fsanitize=address,undefined -fno-omit-frame-pointer)
    elseif(MSVC)
        add_compile_options(/fsanitize=address)
        add_link_options(/INFERASANLIBS)
    else()
        message(WARNING "XLPP_ENABLE_SANITIZERS is not configured for this compiler")
    endif()
endif()

# === Platform detection ===
if(WIN32)
    add_definitions(-D_UNICODE -DUNICODE)
    if(MSVC)
        add_compile_options(/EHsc /bigobj)
    endif()
endif()

# === zlib ===
if(XLPP_USE_BUNDLED_ZLIB)
    set(ZLIB_BUILD_TESTING OFF CACHE BOOL "" FORCE)
    set(ZLIB_BUILD_EXAMPLES OFF CACHE BOOL "" FORCE)
    set(ZLIB_INSTALL OFF CACHE BOOL "" FORCE)
    set(SKIP_INSTALL_ALL ON CACHE BOOL "" FORCE)
    add_subdirectory(third_party/zlib)
    set(ZLIB_TARGET zlibstatic)
else()
    find_package(ZLIB REQUIRED)
    if(NOT TARGET ZLIB::ZLIB)
        add_library(ZLIB::ZLIB ALIAS ${ZLIB_LIBRARIES})
    endif()
    set(ZLIB_TARGET ZLIB::ZLIB)
endif()

# === XLPP Library ===
add_subdirectory(src)

# === Tests ===
if(XLPP_BUILD_TESTS)
    enable_testing()
    add_subdirectory(tests)
endif()

# === Fuzzing ===
if(XLPP_BUILD_FUZZERS)
    add_subdirectory(tests/fuzz)
endif()

# === Samples ===
if(XLPP_BUILD_SAMPLES)
    if(EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/samples/CMakeLists.txt")
        add_subdirectory(samples)
    else()
        message(WARNING "XLPP_BUILD_SAMPLES is ON, but no samples/CMakeLists.txt is present; skipping sample targets")
    endif()
endif()

# === C API ===
if(XLPP_BUILD_CAPI)
    add_subdirectory(bindings/c)
endif()

# === Language bindings ===
# Python links directly against the in-tree native target when requested.
# C# consumes the C ABI and is built by dotnet/NuGet tooling rather than CMake.
if(XLPP_BUILD_BINDINGS)
    add_subdirectory(bindings/python)
endif()

if(XLPP_BUILD_BENCHMARKS)
    add_subdirectory(benchmarks)
endif()

if(XLPP_BUILD_TOOLS)
    add_subdirectory(tools)
endif()

# === Install ===
include(CMakePackageConfigHelpers)

# Export xlpp_static as XLPP::XLPP (the roadmap target naming).
set_target_properties(xlpp_static PROPERTIES EXPORT_NAME XLPP)

install(TARGETS xlpp_static
    EXPORT xlpp-targets
    ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
)
install(DIRECTORY include/ DESTINATION ${CMAKE_INSTALL_INCLUDEDIR}
    FILES_MATCHING PATTERN "*.h"
)
install(EXPORT xlpp-targets
    FILE xlpp-targets.cmake
    NAMESPACE XLPP::
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/xlpp
)

if(XLPP_USE_BUNDLED_ZLIB)
    # The installed XLPP::XLPP target references zlib symbols, so the bundled
    # static archive must be installed and exported too (as XLPP::zlib).
    set_target_properties(zlibstatic PROPERTIES EXPORT_NAME zlib)
    install(TARGETS zlibstatic
        EXPORT xlpp-zlib-targets
        ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
    )
    install(EXPORT xlpp-zlib-targets
        FILE xlpp-zlib-targets.cmake
        NAMESPACE XLPP::
        DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/xlpp
    )
endif()

# === Package ===
write_basic_package_version_file(
    ${CMAKE_CURRENT_BINARY_DIR}/xlpp-config-version.cmake
    VERSION ${PROJECT_VERSION}
    COMPATIBILITY SameMajorVersion
)
configure_package_config_file(
    ${CMAKE_CURRENT_SOURCE_DIR}/cmake/xlpp-config.cmake.in
    ${CMAKE_CURRENT_BINARY_DIR}/xlpp-config.cmake
    INSTALL_DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/xlpp
)
install(FILES
    ${CMAKE_CURRENT_BINARY_DIR}/xlpp-config.cmake
    ${CMAKE_CURRENT_BINARY_DIR}/xlpp-config-version.cmake
    DESTINATION ${CMAKE_INSTALL_LIBDIR}/cmake/xlpp
)
