from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

#stdlib
from io import BytesIO
import os

# test
import pytest
from openpyxl2.tests.helper import compare_xml

# package
from openpyxl2 import Workbook, load_workbook
from openpyxl2.workbook.defined_name.named_range import NamedRange
from openpyxl2.xml.functions import Element, tostring, fromstring
from openpyxl2.xml.constants import XLTX, XLSX, XLSM, XLTM
from .. excel import (
    save_workbook,
    save_virtual_workbook,
    )
from .. workbook import (
    write_workbook,
    write_workbook_rels,
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
    <workbookPr/>
    <bookViews>
      <workbookView activeTab="0"/>
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
    wb.defined_names.append(xlrange)
    root = Element("root")
    _write_defined_names(wb, root)
    xml = tostring(root)
    expected = """
    <root>
     <definedName name="test_range">'Sheet'!$A$1:$B$5</definedName>
    </root>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_workbook_code_name():
    wb = Workbook()
    wb.code_name = u'MyWB'

    content = write_workbook(wb)
    expected = """
    <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr codeName="MyWB"/>
    <bookViews>
      <workbookView activeTab="0"/>
    </bookViews>
    <sheets>
      <sheet name="Sheet" sheetId="1" r:id="rId1"/>
    </sheets>
    <definedNames/>
    <calcPr calcId="124519" fullCalcOnLoad="1"/>
    </workbook>
    """
    diff = compare_xml(content, expected)
    assert diff is None, diff


def test_write_root_rels():
    from ..workbook import write_root_rels

    wb = Workbook()
    xml = write_root_rels(wb)
    expected = """
    <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
      <Relationship Id="rId1" Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"/>
      <Relationship Id="rId2" Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"/>
      <Relationship Id="rId3" Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"/>
    </Relationships>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
