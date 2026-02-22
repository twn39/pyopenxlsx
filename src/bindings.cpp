#include "bindings.hpp"

NB_MODULE(_openxlsx, m) {
    m.doc() = "Python bindings for OpenXLSX";
    init_constants(m);
    init_types(m);
    init_styles(m);
    init_document(m);
    init_workbook(m);
    init_worksheet(m);
    init_cell(m);
}
