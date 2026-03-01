cmake_minimum_required(VERSION 3.15)
project(pyopenxlsx LANGUAGES CXX)

set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_POSITION_INDEPENDENT_CODE ON)

# Enable Link Time Optimization (LTO/IPO)
include(CheckIPOSupported)
check_ipo_supported(RESULT ipo_supported OUTPUT ipo_output)
if(ipo_supported)
    set(CMAKE_INTERPROCEDURAL_OPTIMIZATION TRUE)
else()
    message(STATUS "IPO / LTO not supported: ${ipo_output}")
endif()

# Compiler specific optimizations for Release builds
if(CMAKE_CXX_COMPILER_ID MATCHES "GNU|Clang")
    set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} -O3 -DNDEBUG")
    set(CMAKE_CXX_FLAGS_RELWITHDEBINFO "${CMAKE_CXX_FLAGS_RELWITHDEBINFO} -O3 -g -DNDEBUG")
elseif(CMAKE_CXX_COMPILER_ID MATCHES "MSVC")
    set(CMAKE_CXX_FLAGS_RELEASE "${CMAKE_CXX_FLAGS_RELEASE} /O2 /Oi /Ob2 /DNDEBUG")
    set(CMAKE_CXX_FLAGS_RELWITHDEBINFO "${CMAKE_CXX_FLAGS_RELWITHDEBINFO} /O2 /Oi /Ob2 /Zi /DNDEBUG")
endif()

# Configure OpenXLSX
set(OPENXLSX_BUILD_SAMPLES OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_TESTS OFF CACHE BOOL "" FORCE)
set(OPENXLSX_BUILD_BENCHMARKS OFF CACHE BOOL "" FORCE)

add_subdirectory(third_party/OpenXLSX EXCLUDE_FROM_ALL)

# Configure Python and nanobind
find_package(Python REQUIRED COMPONENTS Interpreter Development.Module)

add_subdirectory(third_party/nanobind)

# Define the python module
nanobind_add_module(_openxlsx 
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
