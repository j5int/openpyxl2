from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

import pytest

from . test_dump import DummyWorkbook

from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def LXMLWorksheet():
    from .. dump_lxml import LXMLWorksheet
    return LXMLWorksheet(DummyWorkbook(), title="TestWorksheet")


def test_write_header(LXMLWorksheet):
    ws = LXMLWorksheet
    doc = ws.write_header()
    header = open(ws.filename)
    xml = header.read()
    expected = """<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetPr>
      <outlinePr summaryRight="1" summaryBelow="1"/>
    </sheetPr>
    <dimension ref="A1:A1"/>
    <sheetViews>
      <sheetView workbookViewId="0">
        <selection sqref="A1" activeCell="A1"/>
      </sheetView>
    </sheetViews>
    <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_close_content(LXMLWorksheet):
    pass


def test_append(LXMLWorksheet):
    ws = LXMLWorksheet
    ws.append([1])
    with open(ws._fileobj_content_name) as rows:
        xml = rows.read()
    expected = """
    <sheetData><row spans="1:1" r="1"></row></sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
