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
        .def("set_left", &XLPageMargins::setLeft)
        .def("right", &XLPageMargins::right)
        .def("set_right", &XLPageMargins::setRight)
        .def("top", &XLPageMargins::top)
        .def("set_top", &XLPageMargins::setTop)
        .def("bottom", &XLPageMargins::bottom)
        .def("set_bottom", &XLPageMargins::setBottom)
        .def("header", &XLPageMargins::header)
        .def("set_header", &XLPageMargins::setHeader)
        .def("footer", &XLPageMargins::footer)
        .def("set_footer", &XLPageMargins::setFooter);

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
        .def("black_and_white", &XLPageSetup::blackAndWhite)
        .def("set_black_and_white", &XLPageSetup::setBlackAndWhite);
}
