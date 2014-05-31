from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

#stdlib
from io import BytesIO
import os

# package
from openpyxl2 import Workbook, load_workbook

# test
import pytest
from openpyxl2.tests.helper import compare_xml

from openpyxl2.writer.excel import (
    save_workbook,
    save_virtual_workbook,
    )
from openpyxl2.writer.worksheet import write_worksheet

from .. workbook import write_workbook, write_workbook_rels


def test_write_auto_filter(datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.active
    ws.cell('F42').value = 'hello'
    ws.auto_filter.ref = 'A1:F1'

    content = write_workbook(wb)
    with open('workbook_auto_filter.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


@pytest.mark.xfail
# This should actually fail because you cannot hide the only sheet of a workbook
def test_write_hidden_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    xml = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
    <workbookPr codeName="ThisWorkbook" defaultThemeVersion="124226"/>
    <bookViews>
      <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" state="hidden" r:id="rId1"/>
    </sheets>
      <definedNames/>
      <calcPr calcId="124519" calcMode="auto" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_empty_workbook(tmpdir):
    tmpdir.chdir()
    wb = Workbook()
    dest_filename = 'empty_book.xlsx'
    save_workbook(wb, dest_filename)
    assert os.path.isfile(dest_filename)


def test_write_virtual_workbook():
    old_wb = Workbook()
    saved_wb = save_virtual_workbook(old_wb)
    new_wb = load_workbook(BytesIO(saved_wb))
    assert new_wb


def test_write_workbook_rels(datadir):
    datadir.chdir()
    wb = Workbook()
    content = write_workbook_rels(wb)
    with open('workbook.xml.rels') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff


def test_write_workbook(datadir):
    datadir.chdir()
    wb = Workbook()
    content = write_workbook(wb)
    with open('workbook.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
