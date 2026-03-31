
#include "internal_access.hpp"

// Structure to hold image info
struct ImageInfo {
    std::string name;       // e.g., "image1.png"
    std::string path;       // e.g., "xl/media/image1.png"
    std::string extension;  // e.g., "png"
};

// Get list of images embedded in the document
std::vector<ImageInfo> get_embedded_images(XLDocument& doc) {
    auto& archive = get_archive(doc);
    std::vector<ImageInfo> images;

    const std::string prefix = "xl/media/";
    // Image extensions to recognize
    static const std::vector<std::string> imageExts = {".png", ".jpg", ".jpeg", ".gif",
                                                       ".bmp", ".emf", ".wmf",  ".tiff"};

    // Single pass over all entries — O(n) instead of O(8000) hasEntry() probes
    for (const auto& entryName : archive.entryNames()) {
        if (entryName.size() <= prefix.size() || entryName.compare(0, prefix.size(), prefix) != 0)
            continue;

        std::string filename = entryName.substr(prefix.size());
        auto dotPos = filename.rfind('.');
        if (dotPos == std::string::npos) continue;

        std::string ext = filename.substr(dotPos);
        // Convert to lowercase for comparison
        for (auto& ch : ext) ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));

        bool isImage = false;
        for (const auto& imageExt : imageExts) {
            if (ext == imageExt) {
                isImage = true;
                break;
            }
        }
        if (!isImage) continue;

        ImageInfo info;
        info.path = entryName;
        info.name = filename;
        info.extension = ext.substr(1);  // Remove leading dot
        images.push_back(std::move(info));
    }

    return images;
}

// Get image data as bytes
py::bytes get_image_data(XLDocument& doc, const std::string& imagePath) {
    auto& archive = get_archive(doc);

    std::string fullPath = imagePath;
    if (imagePath.find('/') == std::string::npos) {
        fullPath = "xl/media/" + imagePath;
    }

    if (!archive.hasEntry(fullPath)) {
        throw std::runtime_error("Image not found in archive: " + fullPath);
    }

    std::string data;
    {
        py::gil_scoped_release release;
        data = archive.getEntry(fullPath);
    }

    return py::bytes(data.data(), data.size());
}

void init_document(py::module_& m) {
    // Bind ImageInfo struct
    py::class_<ImageInfo>(m, "ImageInfo")
        .def_ro("name", &ImageInfo::name, "Image filename (e.g., 'image1.png')")
        .def_ro("path", &ImageInfo::path, "Full path in archive (e.g., 'xl/media/image1.png')")
        .def_ro("extension", &ImageInfo::extension, "File extension (e.g., 'png')")
        .def("__repr__", [](const ImageInfo& self) {
            return "<ImageInfo name='" + self.name + "' path='" + self.path + "'>";
        });

    // Bind XLProperties
    py::class_<XLProperties>(m, "XLProperties")
        .def("set_property",
             [](XLProperties& self, const std::string& name, const std::string& value) {
                 py::gil_scoped_release release;
                 self.setProperty(name, value);
             })
        .def("set_property",
             [](XLProperties& self, const std::string& name, int value) {
                 py::gil_scoped_release release;
                 self.setProperty(name, value);
             })
        .def("set_property",
             [](XLProperties& self, const std::string& name, double value) {
                 py::gil_scoped_release release;
                 self.setProperty(name, value);
             })
        .def("property",
             [](const XLProperties& self, const std::string& name) {
                 py::gil_scoped_release release;
                 return self.property(name);
             })
        .def("delete_property", [](XLProperties& self, const std::string& name) {
            py::gil_scoped_release release;
            self.deleteProperty(name);
        });

    // Bind XLAppProperties
    py::class_<XLAppProperties>(m, "XLAppProperties")
        .def("increment_sheet_count",
             [](XLAppProperties& self, int16_t increment) {
                 py::gil_scoped_release release;
                 self.incrementSheetCount(increment);
             })
        .def("align_worksheets",
             [](XLAppProperties& self, const std::vector<std::string>& names) {
                 py::gil_scoped_release release;
                 self.alignWorksheets(names);
             })
        .def("add_sheet_name",
             [](XLAppProperties& self, const std::string& title) {
                 py::gil_scoped_release release;
                 self.addSheetName(title);
             })
        .def("delete_sheet_name",
             [](XLAppProperties& self, const std::string& title) {
                 py::gil_scoped_release release;
                 self.deleteSheetName(title);
             })
        .def("set_sheet_name",
             [](XLAppProperties& self, const std::string& oldTitle, const std::string& newTitle) {
                 py::gil_scoped_release release;
                 self.setSheetName(oldTitle, newTitle);
             })
        .def("add_heading_pair",
             [](XLAppProperties& self, const std::string& name, int value) {
                 py::gil_scoped_release release;
                 self.addHeadingPair(name, value);
             })
        .def("delete_heading_pair",
             [](XLAppProperties& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.deleteHeadingPair(name);
             })
        .def("set_heading_pair",
             [](XLAppProperties& self, const std::string& name, int newValue) {
                 py::gil_scoped_release release;
                 self.setHeadingPair(name, newValue);
             })
        .def("set_property",
             [](XLAppProperties& self, const std::string& name, const std::string& value) {
                 py::gil_scoped_release release;
                 auto& doc = get_xml_doc(self);
                 auto property = doc.document_element().child(name.c_str());
                 if (property.empty()) property = doc.document_element().append_child(name.c_str());
                 property.text().set(value.c_str());
             })
        .def("property",
             [](const XLAppProperties& self, const std::string& name) {
                 py::gil_scoped_release release;
                 return self.property(name);
             })
        .def("delete_property",
             [](XLAppProperties& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.deleteProperty(name);
             })
        .def("append_sheet_name", &XLAppProperties::appendSheetName)
        .def("prepend_sheet_name", &XLAppProperties::prependSheetName)
        .def("insert_sheet_name", &XLAppProperties::insertSheetName);

    // Bind XLDocument
    py::class_<XLDocument>(m, "XLDocument")
        .def(py::init<>())
        .def(py::init<const std::string&>())
        .def("open",
             [](XLDocument& self, const std::string& path) {
                 py::gil_scoped_release release;
                 self.open(path);
             })
        .def(
            "create",
            [](XLDocument& self, const std::string& name, bool forceOverwrite) {
                py::gil_scoped_release release;
                self.create(name, forceOverwrite);
            },
            py::arg("name"), py::arg("force_overwrite") = true)
        .def("close",
             [](XLDocument& self) {
                 py::gil_scoped_release release;
                 self.close();
             })
        .def("is_open", &XLDocument::isOpen)
        .def("has_macro", &XLDocument::hasMacro)
        .def("name", &XLDocument::name)
        .def("path", &XLDocument::path)
        .def("save",
             [](XLDocument& self) {
                 py::gil_scoped_release release;
                 self.save();
             })
        .def(
            "save_as",
            [](XLDocument& self, const std::string& name, bool forceOverwrite) {
                py::gil_scoped_release release;
                self.saveAs(name, forceOverwrite);
            },
            py::arg("name"), py::arg("force_overwrite") = true)
        .def("workbook", &XLDocument::workbook, py::keep_alive<0, 1>())
        .def(
            "content_types", [](XLDocument& self) { return &self.contentTypes(); },
            py::rv_policy::reference_internal)
        .def(
            "app_properties", [](XLDocument& self) { return &get_app_properties(self); },
            py::rv_policy::reference_internal)
        .def(
            "core_properties", [](XLDocument& self) { return &get_core_properties(self); },
            py::rv_policy::reference_internal)
        .def("property", &XLDocument::property)
        .def("set_property", &XLDocument::setProperty)
        .def("delete_property", &XLDocument::deleteProperty)
        .def("custom_property",
             [](XLDocument& self, const std::string& name) {
                 return self.customProperties().property(name);
             },
             py::arg("name"), "Get a custom document property by name")
        .def("set_custom_property",
             [](XLDocument& self, const std::string& name, py::object value) {
                 if (py::isinstance<py::str>(value)) {
                     self.customProperties().setProperty(name, py::cast<std::string>(value));
                 } else if (py::isinstance<py::int_>(value)) {
                     self.customProperties().setProperty(name, py::cast<int>(value));
                 } else if (py::isinstance<py::float_>(value)) {
                     self.customProperties().setProperty(name, py::cast<double>(value));
                 } else if (py::isinstance<py::bool_>(value)) {
                     self.customProperties().setProperty(name, py::cast<bool>(value));
                 } else {
                     self.customProperties().setProperty(name, py::cast<std::string>(py::str(value)));
                 }
             },
             py::arg("name"), py::arg("value"), "Set a custom document property")
        .def("delete_custom_property",
             [](XLDocument& self, const std::string& name) {
                 self.customProperties().deleteProperty(name);
             },
             py::arg("name"), "Delete a custom document property by name")
        .def("styles", &XLDocument::styles, py::rv_policy::reference_internal)
        .def(
            "add_image",
            [](XLDocument& self, const std::string& name, py::bytes data) {
                std::string imgData(static_cast<const char*>(data.data()), data.size());
                py::gil_scoped_release release;
                return self.addImage(name, std::move(imgData));
            },
            py::arg("name"), py::arg("data"),
            "Add an image to the document archive. Returns the path in the archive.")
        .def(
            "get_image",
            [](XLDocument& self, const std::string& path) {
                py::gil_scoped_release release;
                std::string data = self.getImage(path);
                return py::bytes(data.data(), data.size());
            },
            py::arg("path"), "Get image data as bytes from the document archive.")
        .def(
            "get_embedded_images",
            [](XLDocument& self) {
                py::gil_scoped_release release;
                return get_embedded_images(self);
            },
            "Get list of embedded images in the document. Returns list of dicts with 'name', "
            "'path', 'extension' keys.")
        .def("get_image_data", &get_image_data, py::arg("image_path"),
             "Get image data as bytes. image_path can be full path (e.g., 'xl/media/image1.png') "
             "or just filename (e.g., 'image1.png').")
        .def("get_archive_entries",
             [](XLDocument& self) {
                 auto& archive = get_archive(self);
                 return archive.entryNames();
             },
             "Get a list of all entries (files/directories) in the underlying zip archive.")
        .def("has_archive_entry",
             [](XLDocument& self, const std::string& path) {
                 auto& archive = get_archive(self);
                 return archive.hasEntry(path);
             },
             py::arg("path"),
             "Check if the underlying zip archive contains an entry with the given path.")
        .def("get_archive_entry",
             [](XLDocument& self, const std::string& path) {
                 auto& archive = get_archive(self);
                 if (!archive.hasEntry(path)) {
                     throw std::runtime_error("Entry not found in archive: " + path);
                 }
                 std::string data;
                 {
                     py::gil_scoped_release release;
                     data = archive.getEntry(path);
                 }
                 return py::bytes(data.data(), data.size());
             },
             py::arg("path"),
             "Get the raw bytes of an entry from the underlying zip archive.")
        .def(
            "__enter__", [](XLDocument& self) -> XLDocument& { return self; },
            py::rv_policy::reference)
        .def(
            "__exit__",
            [](XLDocument& self, py::handle exc_type, py::handle exc_value, py::handle traceback) {
                self.close();
            },
            py::arg("exc_type").none(), py::arg("exc_value").none(), py::arg("traceback").none());
}
