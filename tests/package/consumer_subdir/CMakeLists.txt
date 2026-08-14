cmake_minimum_required(VERSION 3.18)
project(xlpp_consumer_subdir CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

# Path to the XL++ source tree (passed by the test driver).
set(XLPP_SOURCE_DIR "" CACHE PATH "Path to the XL++ source directory")

set(XLPP_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(XLPP_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(XLPP_BUILD_CAPI OFF CACHE BOOL "" FORCE)

add_subdirectory(${XLPP_SOURCE_DIR} xlpp)

add_executable(subdir_app main.cpp)
target_link_libraries(subdir_app PRIVATE XLPP::xlpp)

add_custom_target(run_subdir_app COMMAND subdir_app DEPENDS subdir_app)
