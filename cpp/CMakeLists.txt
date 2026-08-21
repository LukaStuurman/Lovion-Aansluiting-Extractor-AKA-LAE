cmake_minimum_required(VERSION 3.18)

project(LovionCoordinateToolCpp LANGUAGES CXX)

add_executable(LovionCoordinateToolCpp
    src/LovionCoordinateTool.cpp
)

target_compile_features(LovionCoordinateToolCpp PRIVATE cxx_std_17)

if (MSVC)
    target_compile_options(LovionCoordinateToolCpp PRIVATE /W4 /permissive-)
else()
    target_compile_options(LovionCoordinateToolCpp PRIVATE -Wall -Wextra -pedantic)
endif()

if (WIN32)
    target_link_libraries(LovionCoordinateToolCpp PRIVATE winhttp)
else()
    message(FATAL_ERROR "LovionCoordinateToolCpp gebruikt WinHTTP en is momenteel Windows-only.")
endif()
