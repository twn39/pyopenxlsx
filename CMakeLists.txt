cmake_minimum_required(VERSION 3.15)
project(pyopenxlsx LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# Configure OpenXLSX
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_BENCHMARKS OFF CACHE BOOL "" FORCE)

# Prevent nowide from searching for Boost
if(NOT TARGET Boost::boost)
    add_library(boost_dummy INTERFACE)
    add_library(Boost::boost ALIAS boost_dummy)
    add_library(boost_config_dummy INTERFACE)
    add_library(Boost::config ALIAS boost_config_dummy)
endif()

add_subdirectory(third_party/OpenXLSX EXCLUDE_FROM_ALL)

# Configure Python and pybind11
find_package(Python REQUIRED COMPONENTS Interpreter Development.Module)
find_package(pybind11 REQUIRED)

# Define the python module
pybind11_add_module(_openxlsx 
    src/bindings.cpp
    src/constants.cpp
    src/types.cpp
    src/styles.cpp
    src/document.cpp
    src/workbook.cpp
    src/worksheet.cpp
    src/cell.cpp
)

# Link dependencies
target_link_libraries(_openxlsx PRIVATE OpenXLSX::OpenXLSX)

# Install steps (handled by scikit-build-core)
install(TARGETS _openxlsx DESTINATION pyopenxlsx)
