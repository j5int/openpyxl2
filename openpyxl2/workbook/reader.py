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
from .workbook import Workbook
from openpyxl2.utils.datetime import CALENDAR_MAC_1904

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


def reader(archive):
    src = archive.read(ARC_WORKBOOK)
    package = WorkbookPackage.from_tree(fromstring(src))
    wb = Workbook()
    if package.properties.date1904:
        wb.excel_base_date = CALENDAR_MAC_1904
    wb.code_name = package.fileVersion.codeName
    wb.active = package.active

    src = archive.read(ARC_WORKBOOK_RELS)
    rels = RelationshipList.from_tree(fromstring(src))

    for sheet in package.sheets:
        yield sheet, rels[sheet.id]
