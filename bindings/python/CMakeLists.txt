cmake_minimum_required(VERSION 3.18)
project(xlpp_python LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Python COMPONENTS Interpreter Development REQUIRED)
find_package(pybind11 CONFIG REQUIRED)

set(XLPP_ROOT "${CMAKE_CURRENT_SOURCE_DIR}/../..")

set(XLPP_LIBRARY "${XLPP_ROOT}/x64/Release/XLPP.lib" CACHE FILEPATH "Path to the XLPP static library")
set(ZLIB_LIBRARY "${XLPP_ROOT}/x64/Release/zlib.lib" CACHE FILEPATH "Path to the zlib static library")

if(NOT EXISTS "${XLPP_LIBRARY}")
    message(FATAL_ERROR "XLPP_LIBRARY does not exist: ${XLPP_LIBRARY}. Build XL++ Release first or pass -DXLPP_LIBRARY=...")
endif()
if(NOT EXISTS "${ZLIB_LIBRARY}")
    message(FATAL_ERROR "ZLIB_LIBRARY does not exist: ${ZLIB_LIBRARY}. Build zlib first or pass -DZLIB_LIBRARY=...")
endif()

pybind11_add_module(xlpp
    src/xlpp_bindings.cpp
)

target_include_directories(xlpp PRIVATE
    "${XLPP_ROOT}/include"
    "${XLPP_ROOT}/third_party/zlib"
)

target_link_libraries(xlpp PRIVATE
    "${XLPP_LIBRARY}"
    "${ZLIB_LIBRARY}"
)

target_compile_definitions(xlpp PRIVATE XLPP_STATIC)

if(MSVC)
    target_compile_options(xlpp PRIVATE /EHsc /bigobj)
endif()
