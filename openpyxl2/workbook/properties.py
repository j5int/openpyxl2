from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import datetime

from openpyxl2.compat import safe_string, unicode
from openpyxl2.date_time import CALENDAR_WINDOWS_1900, datetime_to_W3CDTF, W3CDTF_to_datetime
from openpyxl2.descriptors import Strict, String, Typed, Sequence, Alias
from openpyxl2.xml.functions import ElementTree, Element, SubElement, tostring, fromstring
from openpyxl2.xml.constants import COREPROPS_NS, DCORE_NS, XSI_NS, DCTERMS_NS, DCTERMS_PREFIX



class W3CDateTime(Typed):

    expected_type = datetime.datetime

    def __set__(self, instance, value):
        if value is not None and isinstance(value, unicode):
            try:
                value = W3CDTF_to_datetime(value)
            except ValueError:
                raise ValueError("Value must be W3C datetime format")
        super(W3CDateTime, self).__set__(instance, value)


class DocumentProperties(Strict):
    """High-level properties of the document.
    Defined in ECMA-376 Par2 Annex D
    """

    category = String(allow_none=True)
    contentStatus = String(allow_none=True)
    keywords = Sequence(expected_type=str)
    lastModifiedBy = String(allow_none=True)
    lastPrinted = String(allow_none=True)
    revision = String(allow_none=True)
    version = String(allow_none=True)
    last_modified_by = Alias("lastModifiedBy")

    # Dublin Core Properties
    subject = String(allow_none=True)
    title = String(allow_none=True)
    creator = String(allow_none=True)
    description = String(allow_none=True)
    identifier = String(allow_none=True)
    language = String(allow_none=True)
    created = W3CDateTime(expected_type=datetime.datetime, allow_none=True)
    modified = W3CDateTime(expected_type=datetime.datetime, allow_none=True)

    __fields__ = ("category", "contentStatus", "lastModifiedBy",
                "lastPrinted", "revision", "version", "created", "creator", "description",
                "identifier", "language", "modified", "subject", "title")

    def __init__(self,
                 category=None,
                 contentStatus=None,
                 keywords=[],
                 lastModifiedBy=None,
                 lastPrinted=None,
                 revision=None,
                 version=None,
                 created=datetime.datetime.now(),
                 creator="openpyxl",
                 description=None,
                 identifier=None,
                 language=None,
                 modified=datetime.datetime.now(),
                 subject=None,
                 title=None,
                 ):
        self.contentStatus = contentStatus
        self.lastPrinted = lastPrinted
        self.revision = revision
        self.version = version
        self.creator = creator
        self.lastModifiedBy = lastModifiedBy
        self.creator = creator
        self.modified = modified
        self.created = created
        self.title = title
        self.subject = subject
        self.description = description
        self.identifier = identifier
        self.language = language
        self.keywords = keywords
        self.category = category

    def __iter__(self):
        for attr in self.__fields__:
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


def write_properties(props):
    """Write the core properties to xml."""
    root = Element('{%s}coreProperties' % COREPROPS_NS)
    for attr in ("creator", "title", "description", "subject", "identifier",
                 "language"):
        SubElement(root, '{%s}%s' % (DCORE_NS, attr)).text = getattr(props, attr)

    for attr in ("created", "modified"):
        value = datetime_to_W3CDTF(getattr(props, attr))
        SubElement(root, '{%s}%s' % (DCTERMS_NS, attr),
                   {'{%s}type' % XSI_NS:'%s:W3CDTF' % DCTERMS_PREFIX}).text = value

    for attr in ("lastModifiedBy", "category", "contentStatus",
                 "lastPrinted", "version", "revision"):
        SubElement(root, '{%s}%s' % (COREPROPS_NS, attr)).text = getattr(props, attr)

    node = SubElement(root, '{%s}keywords' % COREPROPS_NS)
    for kw in props.keywords:
        SubElement(node, "{%s}keyword").text = kw
    return tostring(root)


def read_properties(xml_source):
    properties = DocumentProperties()
    root = fromstring(xml_source)
    properties.creator = root.findtext('{%s}creator' % DCORE_NS)
    properties.last_modified_by = root.findtext('{%s}lastModifiedBy' % COREPROPS_NS)

    created_node = root.find('{%s}created' % DCTERMS_NS)
    if created_node is not None:
        properties.created = created_node.text

    modified_node = root.find('{%s}modified' % DCTERMS_NS)
    if modified_node is not None:
        properties.modified = modified_node.text

    return properties



class DocumentSecurity(object):
    """Security information about the document."""

    def __init__(self):
        self.lock_revision = False
        self.lock_structure = False
        self.lock_windows = False
        self.revision_password = ''
        self.workbook_password = ''
