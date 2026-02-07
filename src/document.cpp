#include <headers/XLContentTypes.hpp>
#include <regex>

#include "bindings.hpp"

// Helper to access protected members
class XLAppPropertiesPublic : public XLAppProperties {
   public:
    XMLDocument& getXmlDocument() { return xmlDocument(); }
};

namespace {
// Template trick to access private members of a class, even if it's final.
template <typename Tag, typename Tag::type M>
struct Rob {
    friend typename Tag::type get(Tag) { return M; }
};

struct XLDocumentContentTypes {
    typedef XLContentTypes XLDocument::* type;
};

template struct Rob<XLDocumentContentTypes, &XLDocument::m_contentTypes>;
XLContentTypes XLDocument::* get(XLDocumentContentTypes);

struct XLDocumentAppProperties {
    typedef XLAppProperties XLDocument::* type;
};

template struct Rob<XLDocumentAppProperties, &XLDocument::m_appProperties>;
XLAppProperties XLDocument::* get(XLDocumentAppProperties);

struct XLDocumentCoreProperties {
    typedef XLProperties XLDocument::* type;
};

template struct Rob<XLDocumentCoreProperties, &XLDocument::m_coreProperties>;
XLProperties XLDocument::* get(XLDocumentCoreProperties);

struct XLDocumentArchive {
    typedef IZipArchive XLDocument::* type;
};

template struct Rob<XLDocumentArchive, &XLDocument::m_archive>;
IZipArchive XLDocument::* get(XLDocumentArchive);
}  // namespace

// Structure to hold image info
struct ImageInfo {
    std::string name;       // e.g., "image1.png"
    std::string path;       // e.g., "xl/media/image1.png"
    std::string extension;  // e.g., "png"
};

// Get list of images embedded in the document
std::vector<ImageInfo> get_embedded_images(XLDocument& doc) {
    auto& archive = doc.*get(XLDocumentArchive());
    std::vector<ImageInfo> images;

    // Excel stores images in xl/media/ directory
    // We need to iterate through the archive to find them
    // Since IZipArchive doesn't have a list entries method, we try common patterns

    // Check for common image formats
    const std::vector<std::string> extensions = {"png", "jpg", "jpeg", "gif",
                                                 "bmp", "emf", "wmf",  "tiff"};

    for (const auto& ext : extensions) {
        for (int i = 1; i <= 1000; ++i) {  // Check up to 1000 images per type
            std::string path = "xl/media/image" + std::to_string(i) + "." + ext;
            if (archive.hasEntry(path)) {
                ImageInfo info;
                info.path = path;
                info.name = "image" + std::to_string(i) + "." + ext;
                info.extension = ext;
                images.push_back(info);
            } else if (i > 10) {
                // If we haven't found any in the first 10, skip to next extension
                // to avoid checking all 1000 for each type
                break;
            }
        }
    }

    return images;
}

// Get image data as bytes
py::bytes get_image_data(XLDocument& doc, const std::string& imagePath) {
    auto& archive = doc.*get(XLDocumentArchive());

    std::string fullPath = imagePath;
    // If only image name is provided, prepend xl/media/
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

    return py::bytes(data);
}

void init_document(py::module& m) {
    // Bind ImageInfo struct
    py::class_<ImageInfo>(m, "ImageInfo")
        .def_readonly("name", &ImageInfo::name, "Image filename (e.g., 'image1.png')")
        .def_readonly("path", &ImageInfo::path,
                      "Full path in archive (e.g., 'xl/media/image1.png')")
        .def_readonly("extension", &ImageInfo::extension, "File extension (e.g., 'png')")
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
                 auto& public_self = static_cast<XLAppPropertiesPublic&>(self);
                 auto& doc = public_self.getXmlDocument();
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
        .def("create",
             [](XLDocument& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.create(name);
             })
        .def("close",
             [](XLDocument& self) {
                 py::gil_scoped_release release;
                 self.close();
             })
        .def("is_open", &XLDocument::isOpen)
        .def("name", &XLDocument::name)
        .def("path", &XLDocument::path)
        .def("save",
             [](XLDocument& self) {
                 py::gil_scoped_release release;
                 self.save();
             })
        .def("save_as",
             [](XLDocument& self, const std::string& name) {
                 py::gil_scoped_release release;
                 self.saveAs(name);
             })
        .def("workbook", &XLDocument::workbook, py::keep_alive<0, 1>())
        .def(
            "content_types",
            [](XLDocument& self) { return &(self.*get(XLDocumentContentTypes())); },
            py::return_value_policy::reference_internal)
        .def(
            "app_properties",
            [](XLDocument& self) { return &(self.*get(XLDocumentAppProperties())); },
            py::return_value_policy::reference_internal)
        .def(
            "core_properties",
            [](XLDocument& self) { return &(self.*get(XLDocumentCoreProperties())); },
            py::return_value_policy::reference_internal)
        .def("property", &XLDocument::property)
        .def("set_property", &XLDocument::setProperty)
        .def("delete_property", &XLDocument::deleteProperty)
        .def("styles", &XLDocument::styles, py::return_value_policy::reference_internal)
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
        .def(
            "__enter__", [](XLDocument& self) -> XLDocument& { return self; },
            py::return_value_policy::reference)
        .def("__exit__", [](XLDocument& self, py::object exc_type, py::object exc_value,
                            py::object traceback) { self.close(); });
}
