from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read in global settings to be maintained by the workbook object."""

# package imports
from openpyxl2.xml.functions import fromstring, safe_iterator
from openpyxl2.xml.constants import (
    REL_NS,
    SHEET_MAIN_NS,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    WORKSHEET_TYPE,
    EXTERNAL_LINK,
)
from openpyxl2.utils.datetime  import (
    CALENDAR_WINDOWS_1900,
    CALENDAR_MAC_1904
    )
from openpyxl2.workbook.names.named_range import (
    NamedRange,
    NamedValue,
    split_named_range,
    refers_to_range,
    external_range,
    )

from openpyxl2.packaging.manifest import Manifest
from openpyxl2.packaging.relationship import RelationshipList
from openpyxl2.packaging.workbook import WorkbookPackage

# constants
VALID_WORKSHEET = WORKSHEET_TYPE


def read_content_types(archive):
    """Read content types."""
    xml_source = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(xml_source)
    package = Manifest.from_tree(root)
    for typ in package.Override:
        yield typ.ContentType, typ.PartName


def read_rels(archive):
    """Read relationships for a workbook"""
    xml_source = archive.read(ARC_WORKBOOK_RELS)
    tree = fromstring(xml_source)
    rels = RelationshipList.from_tree(tree)
    for r in rels.Relationship:
        # normalise path
        pth = r.Target
        if pth.startswith("/xl"):
            pth = pth.replace("/xl", "xl")
        elif not pth.startswith("xl") and not pth.startswith(".."):
            pth = "xl/" + pth
        r.Target = pth
        yield r.Id, {'path':r.Target, 'type':r.Type}


def read_sheets(archive):
    """Read worksheet titles and ids for a workbook"""
    xml_source = archive.read(ARC_WORKBOOK)
    tree = fromstring(xml_source)
    for element in safe_iterator(tree, '{%s}sheet' % SHEET_MAIN_NS):
        attrib = element.attrib
        attrib['id'] = attrib["{%s}id" % REL_NS]
        del attrib["{%s}id" % REL_NS]
        if attrib['id']:
            yield attrib
