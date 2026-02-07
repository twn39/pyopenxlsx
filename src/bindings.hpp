#ifndef PYOPENXLSX_BINDINGS_HPP
#define PYOPENXLSX_BINDINGS_HPP

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>

#include <OpenXLSX.hpp>

namespace py = pybind11;
using namespace OpenXLSX;

// 子模块初始化函数声明
void init_constants(py::module& m);
void init_types(py::module& m);
void init_styles(py::module& m);
void init_document(py::module& m);
void init_workbook(py::module& m);
void init_worksheet(py::module& m);
void init_cell(py::module& m);

#endif  // PYOPENXLSX_BINDINGS_HPP
