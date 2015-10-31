from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

from openpyxl2.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)
from openpyxl2.xml.functions import fromstring

from openpyxl2.packaging.relationship import RelationshipList
from openpyxl2.packaging.manifest import Manifest
from .parser import WorkbookPackage

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def reader(archive):
    src = archive.read(ARC_WORKBOOK)
    wb = WorkbookPackage.from_tree(fromstring(src))

    src = archive.read(ARC_WORKBOOK_RELS)
    rels = RelationshipList.from_tree(fromstring(src))

    rels = dict([(r.id, r) for r in rels.Relationship])

    for sheet in wb.sheets:
        yield sheet, rels[sheet.id]
