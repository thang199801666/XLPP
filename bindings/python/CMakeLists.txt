cmake_minimum_required(VERSION 3.18)
project(xlpp_python LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

find_package(Python COMPONENTS Interpreter Development REQUIRED)
find_package(pybind11 CONFIG REQUIRED)

pybind11_add_module(xlpp
    src/xlpp_bindings.cpp
)

# In-tree builds use the exact native target under test. Standalone binding
# builds consume the installed CMake package instead of hard-coded VS paths.
if(TARGET xlpp_static)
    target_link_libraries(xlpp PRIVATE xlpp_static)
    target_include_directories(xlpp PRIVATE "${CMAKE_SOURCE_DIR}/include")
else()
    find_package(XLPP CONFIG REQUIRED)
    target_link_libraries(xlpp PRIVATE XLPP::XLPP)
endif()

target_compile_definitions(xlpp PRIVATE XLPP_STATIC)

if(MSVC)
    target_compile_options(xlpp PRIVATE /EHsc /bigobj)
endif()
