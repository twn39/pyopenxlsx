"""Document core/app and custom property façades for Workbook."""

from __future__ import annotations

from ._openxlsx import XLProperty


class DocumentProperties:
    """
    High-level wrapper for Excel document properties.
    Supports both Core and App/Extended properties.
    """

    _PROPERTY_MAP = {
        "title": XLProperty.Title,
        "subject": XLProperty.Subject,
        "creator": XLProperty.Creator,
        "keywords": XLProperty.Keywords,
        "description": XLProperty.Description,
        "last_modified_by": XLProperty.LastModifiedBy,
        "last_printed": XLProperty.LastPrinted,
        "created": XLProperty.CreationDate,
        "modified": XLProperty.ModificationDate,
        "category": XLProperty.Category,
        "application": XLProperty.Application,
        "doc_security": XLProperty.DocSecurity,
        "scale_crop": XLProperty.ScaleCrop,
        "manager": XLProperty.Manager,
        "company": XLProperty.Company,
        "links_up_to_date": XLProperty.LinksUpToDate,
        "shared_doc": XLProperty.SharedDoc,
        "hyperlink_base": XLProperty.HyperlinkBase,
        "hyperlinks_changed": XLProperty.HyperlinksChanged,
        "app_version": XLProperty.AppVersion,
    }

    def __init__(self, doc):
        self._doc = doc

    def __getitem__(self, key):
        if isinstance(key, XLProperty):
            return self._doc.property(key)

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            return self._doc.property(prop)

        # Fallback to string-based lookup in AppProperties (most flexible)
        return self._doc.app_properties().property(key)

    def __setitem__(self, key, value):
        if isinstance(key, XLProperty):
            self._doc.set_property(key, str(value))
            return

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            self._doc.set_property(prop, str(value))
        else:
            # Fallback to string-based set in AppProperties
            self._doc.app_properties().set_property(key, str(value))

    def __delitem__(self, key):
        if isinstance(key, XLProperty):
            self._doc.delete_property(key)
            return

        prop = self._PROPERTY_MAP.get(key.lower().replace(" ", "_"))
        if prop is not None:
            self._doc.delete_property(prop)
        else:
            self._doc.app_properties().delete_property(key)

    @property
    def title(self):
        return self[XLProperty.Title]

    @title.setter
    def title(self, value):
        self[XLProperty.Title] = value

    @property
    def creator(self):
        return self[XLProperty.Creator]

    @creator.setter
    def creator(self, value):
        self[XLProperty.Creator] = value

    @property
    def last_modified_by(self):
        return self[XLProperty.LastModifiedBy]

    @last_modified_by.setter
    def last_modified_by(self, value):
        self[XLProperty.LastModifiedBy] = value

    @property
    def subject(self):
        return self[XLProperty.Subject]

    @subject.setter
    def subject(self, value):
        self[XLProperty.Subject] = value

    @property
    def description(self):
        return self[XLProperty.Description]

    @description.setter
    def description(self, value):
        self[XLProperty.Description] = value

    @property
    def keywords(self):
        return self[XLProperty.Keywords]

    @keywords.setter
    def keywords(self, value):
        self[XLProperty.Keywords] = value

    @property
    def category(self):
        return self[XLProperty.Category]

    @category.setter
    def category(self, value):
        self[XLProperty.Category] = value

    @property
    def company(self):
        return self[XLProperty.Company]

    @company.setter
    def company(self, value):
        self[XLProperty.Company] = value


class CustomProperties:
    """
    Proxy for custom document properties.
    """

    def __init__(self, doc):
        self._doc = doc

    def __getitem__(self, name):
        return self._doc.custom_property(name)

    def __setitem__(self, name, value):
        self._doc.set_custom_property(name, str(value))

    def __delitem__(self, name):
        self._doc.delete_custom_property(name)

    def __contains__(self, name):
        try:
            val = self._doc.custom_property(name)
            return val != ""
        except Exception:
            return False
