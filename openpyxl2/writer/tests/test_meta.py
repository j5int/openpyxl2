from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

# package imports
from openpyxl2.tests.helper import compare_xml
from openpyxl2.writer.workbook import write_content_types
from openpyxl2.workbook import Workbook


def test_write_content_types(datadir):
    datadir.chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_content_types(wb)
    with open('[Content_Types].xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
