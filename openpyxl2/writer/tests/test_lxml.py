from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# stdlib
import datetime
import decimal
from io import BytesIO

# package
from openpyxl2 import Workbook
from lxml.etree import xmlfile

# test imports
import pytest
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def worksheet():
    from openpyxl2 import Workbook
    wb = Workbook()
    return wb.active

@pytest.fixture
def out():
    return BytesIO()


@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="s"></c>"""),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                         ])
def test_write_cell(out, value, expected):
    from .. lxml_worksheet import write_cell

    wb = Workbook()
    ws = wb.active
    ws['A1'] = value
    with xmlfile(out) as xf:
        write_cell(xf, ws, ws['A1'], ["Hello"])
    xml = out.getvalue()
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def write_rows():
    from .. lxml_worksheet import write_worksheet_data
    return write_worksheet_data


def test_write_sheetdata(out, worksheet, write_rows):
    ws = worksheet
    ws['A1'] = 10
    with xmlfile(out) as xf:
        write_rows(xf, ws, [])
    xml = out.getvalue()
    expected = """<sheetData><row r="1" spans="1:1"><c t="n" r="A1"><v>10</v></c></row></sheetData>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_formula(out, worksheet, write_rows):
    ws = worksheet

    ws.cell('F1').value = 10
    ws.cell('F2').value = 32
    ws.cell('F3').value = '=F1+F2'
    ws.cell('A4').value = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws.cell('B4').value = '='
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws.cell('C4').value = '='
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}

    with xmlfile(out) as xf:
        write_rows(xf, ws, [])

    xml = out.getvalue()
    expected = """
    <sheetData>
      <row r="1" spans="1:6">
        <c r="F1" t="n">
          <v>10</v>
        </c>
      </row>
      <row r="2" spans="1:6">
        <c r="F2" t="n">
          <v>32</v>
        </c>
      </row>
      <row r="3" spans="1:6">
        <c r="F3">
          <f>F1+F2</f>
          <v></v>
        </c>
      </row>
      <row r="4" spans="1:6">
        <c r="A4">
          <f ref="A4:C4" si="0" t="shared">A1+A2+A3</f>
          <v></v>
        </c>
        <c r="B4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
        <c r="C4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_row_height(out, worksheet, write_rows):
    ws = worksheet
    ws.cell('F1').value = 10
    ws.row_dimensions[ws.cell('F1').row].height = 30

    with xmlfile(out) as xf:
        write_rows(xf, ws, {})
    xml = out.getvalue()
    expected = """
     <sheetData>
     <row customHeight="1" ht="30" r="1" spans="1:6">
     <c r="F1" t="n">
       <v>10</v>
     </c>
   </row>
   </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def DummyWorksheet():
    class DummyWorksheet:

        def __init__(self):
            self._styles = {}
            self.column_dimensions = {}
    return DummyWorksheet()


@pytest.fixture
def write_cols():
    from .. lxml_worksheet import write_cols
    return write_cols


@pytest.fixture
def ColumnDimension():
    from openpyxl2.worksheet.dimensions import ColumnDimension
    return ColumnDimension


@pytest.mark.xfail
def test_write_no_cols(out, write_cols, DummyWorksheet):
    with xmlfile(out) as xf:
        write_cols(xf, DummyWorksheet)
    assert out.getvalue() == b""


def test_write_col_widths(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    ws.column_dimensions['A'] = ColumnDimension(width=4)
    with xmlfile(out) as xf:
        write_cols(xf, ws)
    xml = out.getvalue()
    expected = """<cols><col width="4" min="1" max="1" customWidth="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_cols_style(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    ws.column_dimensions['A'] = ColumnDimension()
    ws._styles['A'] = 1
    with xmlfile(out) as xf:
        write_cols(xf, ws)
    xml = out.getvalue()
    expected = """<cols><col max="1" min="1" style="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_lots_cols(out, write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    from openpyxl2.cell import get_column_letter
    for i in range(1, 15):
        label = get_column_letter(i)
        ws._styles[label] = i
        ws.column_dimensions[label] = ColumnDimension()
    with xmlfile(out) as xf:
        write_cols(xf, ws)
    xml = out.getvalue()
    expected = """<cols>
   <col max="1" min="1" style="1"></col>
   <col max="2" min="2" style="2"></col>
   <col max="3" min="3" style="3"></col>
   <col max="4" min="4" style="4"></col>
   <col max="5" min="5" style="5"></col>
   <col max="6" min="6" style="6"></col>
   <col max="7" min="7" style="7"></col>
   <col max="8" min="8" style="8"></col>
   <col max="9" min="9" style="9"></col>
   <col max="10" min="10" style="10"></col>
   <col max="11" min="11" style="11"></col>
   <col max="12" min="12" style="12"></col>
   <col max="13" min="13" style="13"></col>
   <col max="14" min="14" style="14"></col>
 </cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def write_sheet_format():
    from .. worksheet import write_format
    return write_format
