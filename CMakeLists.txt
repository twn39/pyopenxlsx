cmake_minimum_required(VERSION 3.15)
project(pyopenxlsx LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# Configure OpenXLSX
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_BENCHMARKS OFF CACHE BOOL "" FORCE)
add_subdirectory(third_party/OpenXLSX)

# Configure pybind11
add_subdirectory(third_party/pybind11)

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
