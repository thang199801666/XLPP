cmake_minimum_required(VERSION 3.18)
project(xlpp_python LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(Python COMPONENTS Interpreter Development REQUIRED)
find_package(pybind11 CONFIG REQUIRED)

set(XLPP_ROOT "${CMAKE_CURRENT_SOURCE_DIR}/..")

pybind11_add_module(xlpp
    src/xlpp_bindings.cpp
)

target_include_directories(xlpp PRIVATE
    "${XLPP_ROOT}/include"
    "${XLPP_ROOT}/third_party/zlib"
)

target_link_libraries(xlpp PRIVATE
    "${XLPP_ROOT}/x64/Release/XLPP.lib"
    "${XLPP_ROOT}/x64/Release/zlib.lib"
)

target_compile_definitions(xlpp PRIVATE XLPP_STATIC)

if(MSVC)
    target_compile_options(xlpp PRIVATE /EHsc /bigobj)
endif()
