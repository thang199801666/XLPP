cmake_minimum_required(VERSION 3.18)
project(xlpp_consumer CXX)

set(CMAKE_CXX_STANDARD 20)
set(CMAKE_CXX_STANDARD_REQUIRED ON)

find_package(XLPP 1.0 CONFIG REQUIRED)

add_executable(consumer_app main.cpp)
target_link_libraries(consumer_app PRIVATE XLPP::XLPP)

add_custom_target(run_consumer COMMAND consumer_app DEPENDS consumer_app)
