#include <nanobind/stl/optional.h>

#include "bindings.hpp"

void init_chart(py::module_& m) {
    py::enum_<XLChartType>(m, "XLChartType")
        .value("Bar", XLChartType::Bar)
        .value("BarStacked", XLChartType::BarStacked)
        .value("BarPercentStacked", XLChartType::BarPercentStacked)
        .value("Bar3D", XLChartType::Bar3D)
        .value("Bar3DStacked", XLChartType::Bar3DStacked)
        .value("Bar3DPercentStacked", XLChartType::Bar3DPercentStacked)
        .value("Line", XLChartType::Line)
        .value("LineStacked", XLChartType::LineStacked)
        .value("LinePercentStacked", XLChartType::LinePercentStacked)
        .value("Line3D", XLChartType::Line3D)
        .value("Pie", XLChartType::Pie)
        .value("Pie3D", XLChartType::Pie3D)
        .value("Scatter", XLChartType::Scatter)
        .value("Area", XLChartType::Area)
        .value("AreaStacked", XLChartType::AreaStacked)
        .value("AreaPercentStacked", XLChartType::AreaPercentStacked)
        .value("Area3D", XLChartType::Area3D)
        .value("Area3DStacked", XLChartType::Area3DStacked)
        .value("Area3DPercentStacked", XLChartType::Area3DPercentStacked)
        .value("Doughnut", XLChartType::Doughnut)
        .value("Radar", XLChartType::Radar)
        .value("RadarFilled", XLChartType::RadarFilled)
        .value("RadarMarkers", XLChartType::RadarMarkers);

    py::enum_<XLLegendPosition>(m, "XLLegendPosition")
        .value("Bottom", XLLegendPosition::Bottom)
        .value("Left", XLLegendPosition::Left)
        .value("Right", XLLegendPosition::Right)
        .value("Top", XLLegendPosition::Top)
        .value("TopRight", XLLegendPosition::TopRight)
        .value("Hidden", XLLegendPosition::Hidden);

    py::enum_<XLMarkerStyle>(m, "XLMarkerStyle")
        .value("None", XLMarkerStyle::None)
        .value("Circle", XLMarkerStyle::Circle)
        .value("Dash", XLMarkerStyle::Dash)
        .value("Diamond", XLMarkerStyle::Diamond)
        .value("Dot", XLMarkerStyle::Dot)
        .value("Picture", XLMarkerStyle::Picture)
        .value("Plus", XLMarkerStyle::Plus)
        .value("Square", XLMarkerStyle::Square)
        .value("Star", XLMarkerStyle::Star)
        .value("Triangle", XLMarkerStyle::Triangle)
        .value("X", XLMarkerStyle::X)
        .value("Default", XLMarkerStyle::Default);

    py::class_<XLChartSeries>(m, "XLChartSeries")
        .def("set_title", &XLChartSeries::setTitle)
        .def("set_smooth", &XLChartSeries::setSmooth)
        .def("set_marker_style", &XLChartSeries::setMarkerStyle)
        .def("set_data_labels", &XLChartSeries::setDataLabels, py::arg("show_value"),
             py::arg("show_category_name") = false, py::arg("show_percent") = false);

    py::class_<XLAxis>(m, "XLAxis")
        .def("set_title", &XLAxis::setTitle)
        .def("set_min_bounds", &XLAxis::setMinBounds)
        .def("clear_min_bounds", &XLAxis::clearMinBounds)
        .def("set_max_bounds", &XLAxis::setMaxBounds)
        .def("clear_max_bounds", &XLAxis::clearMaxBounds)
        .def("set_major_gridlines", &XLAxis::setMajorGridlines)
        .def("set_minor_gridlines", &XLAxis::setMinorGridlines);

    py::class_<XLChartAnchor>(m, "XLChartAnchor")
        .def(py::init<std::string_view, uint32_t, uint32_t, XLDistance, XLDistance>())
        .def_rw("name", &XLChartAnchor::name)
        .def_rw("row", &XLChartAnchor::row)
        .def_rw("col", &XLChartAnchor::col)
        .def_rw("width", &XLChartAnchor::width)
        .def_rw("height", &XLChartAnchor::height);

    py::class_<XLChart>(m, "XLChart")
        .def("add_series",
             py::overload_cast<const XLWorksheet&, const XLCellRange&, std::string_view,
                               std::optional<XLChartType>, bool>(&XLChart::addSeries),
             py::arg("wks"), py::arg("values"), py::arg("title") = "",
             py::arg("target_chart_type") = py::none(), py::arg("use_secondary_axis") = false)
        .def("add_series",
             py::overload_cast<const XLWorksheet&, const XLCellRange&, const XLCellRange&,
                               std::string_view, std::optional<XLChartType>, bool>(
                 &XLChart::addSeries),
             py::arg("wks"), py::arg("values"), py::arg("categories"), py::arg("title") = "",
             py::arg("target_chart_type") = py::none(), py::arg("use_secondary_axis") = false)
        .def("add_series_ref",
             py::overload_cast<std::string_view, std::string_view, std::string_view,
                               std::optional<XLChartType>, bool>(&XLChart::addSeries),
             py::arg("values_ref"), py::arg("title") = "", py::arg("categories_ref") = "",
             py::arg("target_chart_type") = py::none(), py::arg("use_secondary_axis") = false)
        .def("set_title", &XLChart::setTitle)
        .def("set_style", &XLChart::setStyle)
        .def("set_legend_position", &XLChart::setLegendPosition)
        .def("x_axis", &XLChart::xAxis)
        .def("y_axis", &XLChart::yAxis)
        .def("axis", &XLChart::axis)
        .def("set_show_data_labels", &XLChart::setShowDataLabels, py::arg("show_value"),
             py::arg("show_category") = false, py::arg("show_percent") = false)
        .def("set_series_smooth", &XLChart::setSeriesSmooth)
        .def("set_series_marker", &XLChart::setSeriesMarker);
}
