# Copyright (c) 2010-2014 openpyxl

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
from openpyxl2.date_time import CALENDAR_WINDOWS_1900
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


def test_read_workbook_with_no_core_properties(datadir):
    from openpyxl2.workbook import DocumentProperties
    from openpyxl2.reader.excel import _load_workbook

    datadir.join('genuine').chdir()
    archive = ZipFile('empty_with_no_properties.xlsx')
    wb = Workbook()
    default_props = DocumentProperties()
    _load_workbook(wb, archive, None, False, False)
    prop = wb.properties
    assert prop.creator == default_props.creator
    assert prop.lastModifiedBy == default_props.lastModifiedBy
    assert prop.title == default_props.title
    assert prop.subject == default_props.subject
    assert prop.description == default_props.description
    assert prop.category == default_props.category
    assert prop.keywords == default_props.keywords
    assert prop.created.timetuple()[:9] == default_props.created.timetuple()[:9] # might break if tests run on the stoke of midnight
    assert prop.modified.timetuple()[:9] == prop.created.timetuple()[:9]
