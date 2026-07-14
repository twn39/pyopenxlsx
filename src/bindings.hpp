#ifndef PYOPENXLSX_BINDINGS_HPP
#define PYOPENXLSX_BINDINGS_HPP

#include <nanobind/make_iterator.h>
#include <nanobind/nanobind.h>
#include <nanobind/stl/string.h>
#include <nanobind/stl/string_view.h>
#include <nanobind/stl/tuple.h>
#include <nanobind/stl/vector.h>

// datetime.h requires Python.h to be included first.
// nanobind/nanobind.h above guarantees Python.h is already included.
#include <datetime.h>

#include <OpenXLSX.hpp>
#include <XLAutoFilter.hpp>
#include <XLChart.hpp>
#include <XLChartsheet.hpp>
#include <XLColor.hpp>
#include <XLColumn.hpp>
#include <XLComments.hpp>
#include <XLConditionalFormatting.hpp>
#include <XLDataValidation.hpp>
#include <XLDateTime.hpp>
#include <XLDrawing.hpp>
#include <XLFormulaEngine.hpp>
#include <XLMergeCells.hpp>
#include <XLNumberFormatter.hpp>
#include <XLPageSetup.hpp>
#include <XLPivotTable.hpp>
#include <XLRichText.hpp>
#include <XLRow.hpp>
#include <XLSheet.hpp>
#include <XLSparkline.hpp>
#include <XLStreamReader.hpp>
#include <XLStreamWriter.hpp>
#include <XLStyles.hpp>
#include <XLTables.hpp>
#include <XLThreadedComments.hpp>
#include <XLWorkbook.hpp>
#include <XLWorksheet.hpp>
#include <gsl/gsl>

namespace py = nanobind;

// 子模块初始化函数声明
void init_constants(py::module_& m);
void init_types(py::module_& m);
void init_styles(py::module_& m);
void init_document(py::module_& m);
void init_workbook(py::module_& m);
void init_worksheet(py::module_& m);
void init_cell(py::module_& m);
void init_data_validation(py::module_& m);
void init_tables(py::module_& m);
void init_page_setup(py::module_& m);
void init_rich_text(py::module_& m);
void init_defined_names(py::module_& m);
void init_autofilter(py::module_& m);
void init_chart(py::module_& m);
void init_comments(py::module_& m);
void init_pivot_table(py::module_& m);
void init_streams(py::module_& m);
void init_conditional_formatting(py::module_& m);
void init_formula_engine(py::module_& m);
void init_exceptions(py::module_& m);
void init_row(py::module_& m);
void init_sheet(py::module_& m);
void init_relationships(py::module_& m);

#endif  // PYOPENXLSX_BINDINGS_HPP
