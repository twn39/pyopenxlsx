#include "bindings.hpp"

void init_page_setup(py::module_& m) {
    // XLPageOrientation Enum
    py::enum_<XLPageOrientation>(m, "XLPageOrientation")
        .value("Default", XLPageOrientation::Default)
        .value("Portrait", XLPageOrientation::Portrait)
        .value("Landscape", XLPageOrientation::Landscape);

    // XLPageMargins
    py::class_<XLPageMargins>(m, "XLPageMargins")
        .def(py::init<>())
        .def("left", &XLPageMargins::left)
        .def("set_left", py::overload_cast<double>(&XLPageMargins::setLeft))
        .def("right", &XLPageMargins::right)
        .def("set_right", py::overload_cast<double>(&XLPageMargins::setRight))
        .def("top", &XLPageMargins::top)
        .def("set_top", py::overload_cast<double>(&XLPageMargins::setTop))
        .def("bottom", &XLPageMargins::bottom)
        .def("set_bottom", py::overload_cast<double>(&XLPageMargins::setBottom))
        .def("header", &XLPageMargins::header)
        .def("set_header", py::overload_cast<double>(&XLPageMargins::setHeader))
        .def("footer", &XLPageMargins::footer)
        .def("set_footer", py::overload_cast<double>(&XLPageMargins::setFooter));

    // XLPrintOptions
    py::class_<XLPrintOptions>(m, "XLPrintOptions")
        .def(py::init<>())
        .def("grid_lines", &XLPrintOptions::gridLines)
        .def("set_grid_lines", &XLPrintOptions::setGridLines)
        .def("headings", &XLPrintOptions::headings)
        .def("set_headings", &XLPrintOptions::setHeadings)
        .def("horizontal_centered", &XLPrintOptions::horizontalCentered)
        .def("set_horizontal_centered", &XLPrintOptions::setHorizontalCentered)
        .def("vertical_centered", &XLPrintOptions::verticalCentered)
        .def("set_vertical_centered", &XLPrintOptions::setVerticalCentered);

    // XLPageSetup
    py::class_<XLPageSetup>(m, "XLPageSetup")
        .def(py::init<>())
        .def("paper_size", &XLPageSetup::paperSize)
        .def("set_paper_size", &XLPageSetup::setPaperSize)
        .def("orientation", &XLPageSetup::orientation)
        .def("set_orientation", &XLPageSetup::setOrientation)
        .def("scale", &XLPageSetup::scale)
        .def("set_scale", &XLPageSetup::setScale)
        .def("fit_to_width", &XLPageSetup::fitToWidth)
        .def("set_fit_to_width", &XLPageSetup::setFitToWidth)
        .def("fit_to_height", &XLPageSetup::fitToHeight)
        .def("set_fit_to_height", &XLPageSetup::setFitToHeight)
        .def("page_order", &XLPageSetup::pageOrder)
        .def("set_page_order", &XLPageSetup::setPageOrder)
        .def("use_first_page_number", &XLPageSetup::useFirstPageNumber)
        .def("set_use_first_page_number", &XLPageSetup::setUseFirstPageNumber)
        .def("first_page_number", &XLPageSetup::firstPageNumber)
        .def("set_first_page_number", &XLPageSetup::setFirstPageNumber)
        .def("black_and_white", &XLPageSetup::blackAndWhite)
        .def("set_black_and_white", &XLPageSetup::setBlackAndWhite);

    py::class_<XLHeaderFooter>(m, "XLHeaderFooter")
        .def(py::init<>())
        .def("different_first", &XLHeaderFooter::differentFirst)
        .def("set_different_first", &XLHeaderFooter::setDifferentFirst)
        .def("different_odd_even", &XLHeaderFooter::differentOddEven)
        .def("set_different_odd_even", &XLHeaderFooter::setDifferentOddEven)
        .def("scale_with_doc", &XLHeaderFooter::scaleWithDoc)
        .def("set_scale_with_doc", &XLHeaderFooter::setScaleWithDoc)
        .def("align_with_margins", &XLHeaderFooter::alignWithMargins)
        .def("set_align_with_margins", &XLHeaderFooter::setAlignWithMargins)
        .def("odd_header", &XLHeaderFooter::oddHeader)
        .def("set_odd_header", &XLHeaderFooter::setOddHeader)
        .def("odd_footer", &XLHeaderFooter::oddFooter)
        .def("set_odd_footer", &XLHeaderFooter::setOddFooter)
        .def("even_header", &XLHeaderFooter::evenHeader)
        .def("set_even_header", &XLHeaderFooter::setEvenHeader)
        .def("even_footer", &XLHeaderFooter::evenFooter)
        .def("set_even_footer", &XLHeaderFooter::setEvenFooter)
        .def("first_header", &XLHeaderFooter::firstHeader)
        .def("set_first_header", &XLHeaderFooter::setFirstHeader)
        .def("first_footer", &XLHeaderFooter::firstFooter)
        .def("set_first_footer", &XLHeaderFooter::setFirstFooter);
}
