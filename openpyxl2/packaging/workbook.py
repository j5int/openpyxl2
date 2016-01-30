from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

import posixpath

from openpyxl2.xml.constants import (
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
)
from openpyxl2.xml.functions import fromstring

from openpyxl2.packaging.relationship import get_dependents
from openpyxl2.packaging.manifest import Manifest
from openpyxl2.workbook.parser import WorkbookPackage
from openpyxl2.workbook.workbook import Workbook
from openpyxl2.workbook.defined_name import (
    _unpack_print_area,
    _unpack_print_titles,
)
from openpyxl2.workbook.external_link.external import read_external_link

from openpyxl2.utils.datetime import CALENDAR_MAC_1904

chart_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet"
worksheet_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"


class WorkbookParser:

    def __init__(self, archive):
        self.archive = archive
        self.wb = Workbook()
        self.sheets = []
        self.rels = get_dependents(self.archive, ARC_WORKBOOK_RELS)


    def parse(self):
        src = self.archive.read(ARC_WORKBOOK)
        node = fromstring(src)
        package = WorkbookPackage.from_tree(node)
        if package.properties.date1904:
            wb.excel_base_date = CALENDAR_MAC_1904
        self.wb.code_name = package.properties.codeName
        self.wb.active = package.active
        self.sheets = package.sheets

        for ext_ref in package.externalReferences:
            rel = self.rels[ext_ref.id]
            self.wb._external_links.append(read_external_link(self.archive,
                                                              rel.Target))

        if package.definedNames:
            self.wb.defined_names = package.definedNames


    def find_sheets(self):

        for sheet in self.sheets:
            yield sheet, self.rels[sheet.id]


    def assign_names(self):
        """
        Bind reserved names to parsed worksheets
        """
        defns = []
        for defn in self.wb.defined_names.definedName:
            reserved = defn.is_reserved
            if reserved:
                sheet = self.wb._sheets[defn.localSheetId]
                if reserved == "Print_Titles":
                    rows, cols = _unpack_print_titles(defn)
                    sheet.print_title_rows = rows
                    sheet.print_title_cols = cols
                elif reserved == "Print_Area":
                    sheet.print_area = _unpack_print_area(defn)
                else:
                    defns.append(defn)
                continue
            else:
                defns.append(defn)
        self.wb.defined_names.definedName = defns