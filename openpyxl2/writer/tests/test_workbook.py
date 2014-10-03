from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

#stdlib
from io import BytesIO
import os

# test
import pytest
from openpyxl2.tests.helper import compare_xml

# package
from openpyxl2 import Workbook, load_workbook
from openpyxl2.workbook.names.named_range import NamedRange
from openpyxl2.xml.functions import Element, tostring
from openpyxl2.xml.constants import XLTX, XLSX, XLSM, XLTM
from .. excel import (
    save_workbook,
    save_virtual_workbook,
    )
from .. workbook import (
    write_workbook,
    write_workbook_rels,
    write_content_types,
)


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


def test_write_hidden_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    wb.create_sheet()
    xml = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>
    <workbookPr/>
    <bookViews>
      <workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" state="hidden" r:id="rId1"/>
      <sheet name="Sheet1" sheetId="2" r:id="rId2"/>
    </sheets>
      <definedNames/>
      <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_hidden_single_worksheet():
    wb = Workbook()
    ws = wb.active
    ws.sheet_state = ws.SHEETSTATE_HIDDEN
    with pytest.raises(ValueError):
        write_workbook(wb)


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


def test_write_named_range():
    from openpyxl2.writer.workbook import _write_defined_names
    wb = Workbook()
    ws = wb.active
    xlrange = NamedRange('test_range', [(ws, "A1:B5")])
    wb._named_ranges.append(xlrange)
    root = Element("root")
    _write_defined_names(wb, root)
    xml = tostring(root)
    expected = """
    <root>
     <s:definedName xmlns:s="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="test_range">'Sheet'!$A$1:$B$5</s:definedName>
    </root>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def check_content_type_workbook(wb, wb_type, as_template):
    assert wb_type in write_content_types(wb, as_template=as_template)


@pytest.mark.parametrize('tmpl, keep_vba, wb_type, as_template', [
    ('empty.xlsx', False, XLTX),
    ('empty.xlsm', True, XLTM),
    ('empty.xltx', False, XLTX),
    ('empty.xltm', True, XLTM)
])
def write_content_types_as_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = load_workbook(tmpl, keep_vba=keep_vba)
    check_content_type_workbook(wb, wb_type, True)


@pytest.mark.parametrize('tmpl, keep_vba, wb_type, as_template', [
    ('empty.xlsx', False, XLSX),
    ('empty.xlsm', True, XLSM),
    ('empty.xltx', False, XLSX),
    ('empty.xltm', True, XLSM)
])
def write_content_types_as_no_template(datadir, tmpl, keep_vba, wb_type):
    datadir.chdir()

    wb = load_workbook(tmpl, keep_vba=keep_vba)
    check_content_type_workbook(wb, wb_type, False)
