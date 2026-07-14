#include <headers/XLException.hpp>

#include "bindings.hpp"

using namespace OpenXLSX;

void init_exceptions(py::module_& m) {
    // Map OpenXLSX exception hierarchy onto Python exceptions (no BC aliases).
    py::exception<XLException> xl_exc(m, "XLException", PyExc_RuntimeError);
    py::exception<XLOverflowError> xl_overflow(m, "XLOverflowError", xl_exc);
    py::exception<XLValueTypeError> xl_value(m, "XLValueTypeError", xl_exc);
    py::exception<XLCellAddressError> xl_addr(m, "XLCellAddressError", xl_exc);
    py::exception<XLInputError> xl_input(m, "XLInputError", xl_exc);
    py::exception<XLInternalError> xl_internal(m, "XLInternalError", xl_exc);
    py::exception<XLPropertyError> xl_prop(m, "XLPropertyError", xl_exc);
    py::exception<XLSheetError> xl_sheet(m, "XLSheetError", xl_exc);
    py::exception<XLDateTimeError> xl_dt(m, "XLDateTimeError", xl_exc);
    py::exception<XLFormulaError> xl_formula(m, "XLFormulaError", xl_exc);

    (void)xl_overflow;
    (void)xl_value;
    (void)xl_addr;
    (void)xl_input;
    (void)xl_internal;
    (void)xl_prop;
    (void)xl_sheet;
    (void)xl_dt;
    (void)xl_formula;
}
