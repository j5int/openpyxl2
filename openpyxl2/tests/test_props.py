# Copyright (c) 2010-2016 openpyxl

# Python stdlib imports
from zipfile import ZipFile
from datetime import datetime

# test imports
import pytest

# package imports
from openpyxl2.tests.helper import compare_xml
from openpyxl2.writer.workbook import (
    write_properties_app
)
from openpyxl2.xml.constants import ARC_CORE
from openpyxl2.utils.datetime  import CALENDAR_WINDOWS_1900
from openpyxl2.workbook import DocumentProperties, Workbook


@pytest.mark.parametrize("filename", ['empty.xlsx', 'empty_libre.xlsx'])
def test_read_sheets_titles(datadir, filename):
    from openpyxl2.reader.workbook import read_sheets

    datadir.join("genuine").chdir()
    archive = ZipFile(filename)
    sheet_titles = [s['name'] for s in read_sheets(archive)]
    assert sheet_titles == ['Sheet1 - Text', 'Sheet2 - Numbers', 'Sheet3 - Formulas', 'Sheet4 - Dates']


def test_write_properties_app(datadir):
    datadir.join("writer").chdir()
    wb = Workbook()
    wb.create_sheet()
    wb.create_sheet()
    content = write_properties_app(wb)
    with open('app.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff
