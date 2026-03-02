# PyOpenXLSX Project Context

`pyopenxlsx` is a high-performance Python binding for the [OpenXLSX](https://github.com/twn39/OpenXLSX) C++ library, using `nanobind` and `scikit-build-core`.

## Project Overview

- **Purpose**: Provides a fast, Pythonic interface for reading, writing, and manipulating Excel (.xlsx) files.
- **Architecture**:
    - **C++ Core**: Uses `OpenXLSX` (located in `third_party/OpenXLSX`) for low-level XLSX manipulation.
    - **Bindings**: `src/bindings.cpp` and related `.cpp` files in `src/` define the nanobind modules.
    - **Python API**: `src/pyopenxlsx/` contains the Python-side wrappers that provide a more ergonomic and Pythonic API (e.g., properties, context managers, async support).
- **Key Technologies**: C++17, Python 3.11+, nanobind, CMake, scikit-build-core, pytest.

## Building and Running

### Development Setup
The project uses `uv` for dependency management.

```bash
# Install development dependencies and the package in editable mode
uv pip install -e .
```

### Build Commands
Since it uses `scikit-build-core`, standard Python build tools work:

```bash
# Build using uv
uv build
```

### Testing
Tests are located in the `tests/` directory and use `pytest`.

```bash
# Run all tests
uv run pytest

# Run tests with coverage
uv run pytest --cov=pyopenxlsx

# Run benchmarks
uv run pytest tests/test_benchmark.py
```

## Development Conventions

- **Hybrid Implementation**: Low-level performance-critical logic resides in C++, while the high-level user-facing API is implemented in Python in `src/pyopenxlsx/`.
- **Performance Optimization**:
    - **Fast Path**: Use `Worksheet.set_cell_value`, `write_rows`, or `set_cells` for bulk updates. These bypass Python `Cell` object creation and are 10-20x faster.
    - **Bulk Read/Write**: Supports `numpy` and buffer protocols (via `write_range`, `get_range_values`) for high-speed numeric data processing.
- **Memory Safety**:
    - Uses `WeakValueDictionary` for worksheet and cell caching to ensure objects are garbage collected when no longer referenced externally.
    - Utilizes `weakref` for back-references (e.g., `Cell` -> `Worksheet`) to prevent circular reference cycles.
- **Async Support**: Many I/O operations have async counterparts (e.g., `Workbook.save_async`, `load_workbook_async`) implemented using `asyncio` and thread pools.
- **Type Safety**:
    - Type stubs (`.pyi` files) are provided in `src/pyopenxlsx/` for better IDE support and static analysis.
    - The project includes a `py.typed` marker.
- **Code Style**:
    - Python: Follows standard practices, linted/formatted via `ruff`.
    - C++: C++17 standard, formatted via `.clang-format`.
- **CI/CD**:
    - Uses `cibuildwheel` to build wheels for Linux, macOS, and Windows across multiple Python versions (3.11 to 3.14).
    - Note: 32-bit Windows and free-threaded Python versions are currently skipped in CI due to linking/compatibility issues.

## Key Files and Directories

- `src/bindings.cpp`: Entry point for nanobind module definitions.
- `src/pyopenxlsx/`: Python package source code and type stubs (`.pyi`).
- `src/*.cpp`: Modularized C++ binding implementations for Cells, Styles, etc.
- `src/pyopenxlsx/workbook.py`: Main `Workbook` class implementation.
- `CMakeLists.txt`: Root CMake configuration for building the C++ extension.
- `pyproject.toml`: Main project configuration, dependencies, and build metadata.
- `third_party/OpenXLSX/`: Submodule for the underlying C++ library.
- `tests/`: Comprehensive test suite covering cells, styles, workbook/worksheet operations, and performance.
