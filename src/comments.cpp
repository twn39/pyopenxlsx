#include "bindings.hpp"

void init_comments(py::module_& m) {
    py::class_<XLThreadedComment>(m, "XLThreadedComment")
        .def(py::init<>())
        .def_prop_ro("valid", &XLThreadedComment::valid)
        .def_prop_ro("ref", &XLThreadedComment::ref)
        .def_prop_ro("id", &XLThreadedComment::id)
        .def_prop_ro("parent_id", &XLThreadedComment::parentId)
        .def_prop_ro("person_id", &XLThreadedComment::personId)
        .def_prop_ro("text", &XLThreadedComment::text)
        .def_prop_rw("is_resolved", &XLThreadedComment::isResolved,
                     &XLThreadedComment::setResolved);

    py::class_<XLThreadedComments>(m, "XLThreadedComments")
        .def(py::init<>())
        .def("comment", &XLThreadedComments::comment, py::arg("ref"))
        .def("replies", &XLThreadedComments::replies, py::arg("parent_id"))
        .def("add_comment", &XLThreadedComments::addComment, py::arg("ref"), py::arg("person_id"),
             py::arg("text"))
        .def("add_reply", &XLThreadedComments::addReply, py::arg("parent_id"), py::arg("person_id"),
             py::arg("text"))
        .def("delete_comment", &XLThreadedComments::deleteComment, py::arg("ref"));

    py::class_<XLPerson>(m, "XLPerson")
        .def(py::init<>())
        .def_prop_ro("valid", &XLPerson::valid)
        .def_prop_ro("id", &XLPerson::id)
        .def_prop_ro("display_name", &XLPerson::displayName);

    py::class_<XLPersons>(m, "XLPersons")
        .def(py::init<>())
        .def("person", &XLPersons::person, py::arg("id"))
        .def("add_person", &XLPersons::addPerson, py::arg("display_name"));
}
