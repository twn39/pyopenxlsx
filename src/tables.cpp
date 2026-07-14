#include <headers/XLSlicer.hpp>
#include <headers/XLSlicerCollection.hpp>

#include "bindings.hpp"

using namespace OpenXLSX;
using namespace nanobind::literals;


void init_tables(py::module_& m) {
    py::enum_<XLTotalsRowFunction>(m, "XLTotalsRowFunction")
        .value("None", XLTotalsRowFunction::None)
        .value("Sum", XLTotalsRowFunction::Sum)
        .value("Min", XLTotalsRowFunction::Min)
        .value("Max", XLTotalsRowFunction::Max)
        .value("Average", XLTotalsRowFunction::Average)
        .value("Count", XLTotalsRowFunction::Count)
        .value("CountNums", XLTotalsRowFunction::CountNums)
        .value("StdDev", XLTotalsRowFunction::StdDev)
        .value("Var", XLTotalsRowFunction::Var)
        .value("Custom", XLTotalsRowFunction::Custom);

    py::class_<XLTableColumn>(m, "XLTableColumn")
        .def("id", &XLTableColumn::id)
        .def("name", &XLTableColumn::name)
        .def("set_name", &XLTableColumn::setName)
        .def("totals_row_function", &XLTableColumn::totalsRowFunction)
        .def("set_totals_row_function", &XLTableColumn::setTotalsRowFunction)
        .def("totals_row_label", &XLTableColumn::totalsRowLabel)
        .def("set_totals_row_label", &XLTableColumn::setTotalsRowLabel)
        .def("calculated_column_formula", &XLTableColumn::calculatedColumnFormula)
        .def("set_calculated_column_formula", &XLTableColumn::setCalculatedColumnFormula)
        .def("totals_row_formula", &XLTableColumn::totalsRowFormula)
        .def("set_totals_row_formula", &XLTableColumn::setTotalsRowFormula);

    py::class_<XLTable>(m, "XLTable")
        .def(py::init<>())
        .def("name", &XLTable::name)
        .def("set_name", &XLTable::setName)
        .def("display_name", &XLTable::displayName)
        .def("set_display_name", &XLTable::setDisplayName)
        .def("range_reference", &XLTable::rangeReference)
        .def("set_range_reference", &XLTable::setRangeReference)
        .def("style_name", &XLTable::styleName)
        .def("set_style_name", &XLTable::setStyleName)
        .def("comment", &XLTable::comment)
        .def("set_comment", &XLTable::setComment)
        .def("show_row_stripes", &XLTable::showRowStripes)
        .def("set_show_row_stripes", &XLTable::setShowRowStripes)
        .def("show_column_stripes", &XLTable::showColumnStripes)
        .def("set_show_column_stripes", &XLTable::setShowColumnStripes)
        .def("show_first_column", &XLTable::showFirstColumn)
        .def("set_show_first_column", &XLTable::setShowFirstColumn)
        .def("show_last_column", &XLTable::showLastColumn)
        .def("set_show_last_column", &XLTable::setShowLastColumn)
        .def("show_header_row", &XLTable::showHeaderRow)
        .def("set_show_header_row", &XLTable::setShowHeaderRow)
        .def("show_totals_row", &XLTable::showTotalsRow)
        .def("set_show_totals_row", &XLTable::setShowTotalsRow)
        .def("append_column", &XLTable::appendColumn)
        .def("column", py::overload_cast<std::string_view>(&XLTable::column, py::const_))
        .def("column", py::overload_cast<uint32_t>(&XLTable::column, py::const_));

    py::class_<XLSlicerOptions>(m, "XLSlicerOptions")
        .def(py::init<>())
        .def_rw("name", &XLSlicerOptions::name)
        .def_rw("caption", &XLSlicerOptions::caption)
        .def_rw("width", &XLSlicerOptions::width)
        .def_rw("height", &XLSlicerOptions::height)
        .def_rw("offset_x", &XLSlicerOptions::offsetX)
        .def_rw("offset_y", &XLSlicerOptions::offsetY);

    py::enum_<XLSlicerStyle>(m, "XLSlicerStyle")
        .value("Light1", XLSlicerStyle::Light1)
        .value("Light2", XLSlicerStyle::Light2)
        .value("Light3", XLSlicerStyle::Light3)
        .value("Light4", XLSlicerStyle::Light4)
        .value("Light5", XLSlicerStyle::Light5)
        .value("Light6", XLSlicerStyle::Light6)
        .value("Dark1", XLSlicerStyle::Dark1)
        .value("Dark2", XLSlicerStyle::Dark2)
        .value("Dark3", XLSlicerStyle::Dark3)
        .value("Dark4", XLSlicerStyle::Dark4)
        .value("Dark5", XLSlicerStyle::Dark5)
        .value("Dark6", XLSlicerStyle::Dark6)
        .value("Other1", XLSlicerStyle::Other1)
        .value("Other2", XLSlicerStyle::Other2)
        .value("Custom", XLSlicerStyle::Custom);

    py::class_<XLSlicer>(m, "XLSlicer")
        .def(py::init<>())
        .def_prop_ro("valid", &XLSlicer::valid)
        .def_prop_ro("name", &XLSlicer::name)
        .def_prop_ro("caption", &XLSlicer::caption)
        .def_prop_ro("cache", &XLSlicer::cache)
        .def_prop_ro("style", &XLSlicer::style)
        .def_prop_ro("style_raw", &XLSlicer::styleRaw)
        .def_prop_ro("show_caption", &XLSlicer::showCaption)
        .def_prop_ro("column_count", &XLSlicer::columnCount)
        .def_prop_ro("locked_position", &XLSlicer::lockedPosition)
        .def_prop_ro("cell_ref", &XLSlicer::cellRef)
        .def_prop_ro("width", &XLSlicer::width)
        .def_prop_ro("height", &XLSlicer::height)
        .def_prop_ro("items", &XLSlicer::items)
        .def_prop_ro("selected_items", &XLSlicer::selectedItems)
        .def_prop_ro("is_sort_descending", &XLSlicer::isSortDescending)
        .def(
            "set_caption",
            [](XLSlicer& self, std::string_view caption) -> XLSlicer& {
                return self.setCaption(caption);
            },
            "caption"_a, py::rv_policy::reference_internal)
        .def(
            "set_style",
            [](XLSlicer& self, XLSlicerStyle style) -> XLSlicer& { return self.setStyle(style); },
            "style"_a, py::rv_policy::reference_internal)
        .def(
            "set_style_raw",
            [](XLSlicer& self, std::string_view raw) -> XLSlicer& { return self.setStyleRaw(raw); },
            "raw_style_name"_a, py::rv_policy::reference_internal)
        .def(
            "show_only",
            [](XLSlicer& self, const std::vector<std::string>& items) -> XLSlicer& {
                return self.showOnly(items);
            },
            "items"_a, py::rv_policy::reference_internal)
        .def(
            "show_all", [](XLSlicer& self) -> XLSlicer& { return self.showAll(); },
            py::rv_policy::reference_internal)
        .def(
            "move_to",
            [](XLSlicer& self, std::string_view cell_ref) -> XLSlicer& {
                return self.moveTo(cell_ref);
            },
            "cell_ref"_a, py::rv_policy::reference_internal)
        .def(
            "resize",
            [](XLSlicer& self, uint32_t width_px, uint32_t height_px) -> XLSlicer& {
                return self.resize(width_px, height_px);
            },
            "width_px"_a, "height_px"_a, py::rv_policy::reference_internal);

    py::class_<XLSlicerBuilder>(m, "XLSlicerBuilder")
        .def(
            "name",
            [](XLSlicerBuilder& self, std::string_view n) -> XLSlicerBuilder& {
                return self.name(n);
            },
            "n"_a, py::rv_policy::reference_internal)
        .def(
            "caption",
            [](XLSlicerBuilder& self, std::string_view c) -> XLSlicerBuilder& {
                return self.caption(c);
            },
            "c"_a, py::rv_policy::reference_internal)
        .def(
            "style",
            [](XLSlicerBuilder& self, XLSlicerStyle s) -> XLSlicerBuilder& {
                return self.style(s);
            },
            "s"_a, py::rv_policy::reference_internal)
        .def(
            "style_raw",
            [](XLSlicerBuilder& self, std::string_view raw) -> XLSlicerBuilder& {
                return self.styleRaw(raw);
            },
            "raw_name"_a, py::rv_policy::reference_internal)
        .def(
            "size",
            [](XLSlicerBuilder& self, uint32_t w, uint32_t h) -> XLSlicerBuilder& {
                return self.size(w, h);
            },
            "width_px"_a, "height_px"_a, py::rv_policy::reference_internal)
        .def(
            "show_only",
            [](XLSlicerBuilder& self, const std::vector<std::string>& items) -> XLSlicerBuilder& {
                return self.showOnly(items);
            },
            "items"_a, py::rv_policy::reference_internal)
        .def(
            "column_count",
            [](XLSlicerBuilder& self, int cols) -> XLSlicerBuilder& {
                return self.columnCount(cols);
            },
            "cols"_a, py::rv_policy::reference_internal)
        .def(
            "sort_descending",
            [](XLSlicerBuilder& self, bool desc) -> XLSlicerBuilder& {
                return self.sortDescending(desc);
            },
            "desc"_a = true, py::rv_policy::reference_internal)
        .def(
            "locked_position",
            [](XLSlicerBuilder& self, bool locked) -> XLSlicerBuilder& {
                return self.lockedPosition(locked);
            },
            "locked"_a = true, py::rv_policy::reference_internal)
        .def(
            "offset",
            [](XLSlicerBuilder& self, int32_t dx, int32_t dy) -> XLSlicerBuilder& {
                return self.offset(dx, dy);
            },
            "dx"_a, "dy"_a, py::rv_policy::reference_internal)
        .def("build", &XLSlicerBuilder::build);

    py::class_<XLSlicerCollection>(m, "XLSlicerCollection")
        .def_prop_ro("count", &XLSlicerCollection::count)
        .def("__len__", &XLSlicerCollection::count)
        .def_prop_ro("empty", &XLSlicerCollection::empty)
        .def_prop_ro("valid", &XLSlicerCollection::valid)
        .def("contains", &XLSlicerCollection::contains, "name"_a)
        .def("find", &XLSlicerCollection::find, "name"_a)
        .def(
            "__getitem__",
            [](XLSlicerCollection& self, py::handle key) -> XLSlicer& {
                if (py::isinstance<py::int_>(key)) {
                    auto index = py::cast<size_t>(key);
                    if (index >= self.count()) throw py::index_error();
                    return self[index];
                }
                return self[py::cast<std::string>(key)];
            },
            py::rv_policy::reference_internal)
        .def(
            "add",
            [](XLSlicerCollection& self, std::string_view cell_ref, const XLTable& table,
               std::string_view column_name) { return self.add(cell_ref, table, column_name); },
            "cell_ref"_a, "table"_a, "column_name"_a)
        .def(
            "add_pivot",
            [](XLSlicerCollection& self, std::string_view cell_ref, const XLPivotTable& pivot,
               std::string_view field_name) { return self.add(cell_ref, pivot, field_name); },
            "cell_ref"_a, "pivot_table"_a, "field_name"_a)
        .def(
            "remove",
            [](XLSlicerCollection& self, py::handle key) {
                if (py::isinstance<py::int_>(key)) {
                    self.remove(py::cast<size_t>(key));
                } else {
                    self.remove(py::cast<std::string>(key));
                }
            },
            "key"_a)
        .def(
            "__iter__",
            [](XLSlicerCollection& self) {
                return py::make_iterator(py::type<XLSlicerCollection>(), "iterator", self.begin(),
                                         self.end());
            },
            py::keep_alive<0, 1>());

    py::class_<XLTables>(m, "XLTables")
        .def(py::init<>())
        .def("count", &XLTables::count)
        .def("__len__", &XLTables::count)
        .def("__getitem__",
             [](const XLTables& self, size_t index) {
                 if (index >= self.count()) throw py::index_error();
                 return self[index];
             })
        .def("get_table", &XLTables::table)
        .def("add", py::overload_cast<std::string_view, std::string_view>(&XLTables::add),
             "name"_a, "range"_a)
        .def("add_range", py::overload_cast<std::string_view, const XLCellRange&>(&XLTables::add),
             "name"_a, "range"_a);
}
