#ifndef PYOPENXLSX_BINDINGS_HPP
#define PYOPENXLSX_BINDINGS_HPP

#include <nanobind/make_iterator.h>
#include <nanobind/nanobind.h>
#include <nanobind/stl/string.h>
#include <nanobind/stl/vector.h>

#include <OpenXLSX.hpp>

namespace py = nanobind;
using namespace OpenXLSX;

// 子模块初始化函数声明
void init_constants(py::module_& m);
void init_types(py::module_& m);
void init_styles(py::module_& m);
void init_document(py::module_& m);
void init_workbook(py::module_& m);
void init_worksheet(py::module_& m);
void init_cell(py::module_& m);

#endif  // PYOPENXLSX_BINDINGS_HPP
