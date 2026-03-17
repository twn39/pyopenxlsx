#include "bindings.hpp"

NB_MODULE(_openxlsx, m) {
    m.doc() = "Python bindings for OpenXLSX";
    m.attr("__version__") = "0.4.0";
    init_constants(m);
    init_types(m);
    init_styles(m);
    init_document(m);
    init_workbook(m);
    init_worksheet(m);
    init_cell(m);
    init_rich_text(m);
    init_defined_names(m);
}
