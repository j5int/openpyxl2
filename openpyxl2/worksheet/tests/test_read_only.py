from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from io import BytesIO

import pytest

from openpyxl2.cell.read_only import EMPTY_CELL
from openpyxl2.styles.styleable import StyleArray
from openpyxl2.xml.functions import fromstring

@pytest.fixture
def DummyWorkbook():
    class Workbook:
        epoch = None
        _cell_styles = [StyleArray([0, 0, 0, 0, 0, 0, 0, 0, 0])]
        data_only = False

        def __init__(self):
            self.sheetnames = []

    return Workbook()


@pytest.fixture
def ReadOnlyWorksheet():
    from ..read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


class TestReadOnlyWorksheet:

    def test_from_xml(self, datadir, ReadOnlyWorksheet):

        datadir.chdir()

        ws = ReadOnlyWorksheet(DummyWorkbook(), "Sheet", "", "sheet_inline_strings.xml", [])
        cells = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"


    def test_read_row(self, DummyWorkbook, ReadOnlyWorksheet):

        src = b"""
        <sheetData  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
        <row r="1" spans="4:27">
          <c r="D1">
            <v>1</v>
          </c>
          <c r="K1">
            <v>0.01</v>
          </c>
          <c r="AA1">
            <v>100</v>
          </c>
        </row>
        </sheetData>
        """

        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

        xml = fromstring(src)
        row = tuple(ws._get_row(xml, 11, 11))
        values = [c.value for c in row]
        assert values == [0.01]

        row = tuple(ws._get_row(xml, 1, 11))
        values = [c.value for c in row]
        assert values == [None, None, None, 1, None, None, None, None, None, None, 0.01]


    def test_read_empty_row(self, DummyWorkbook, ReadOnlyWorksheet):

        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

        src = """
        <row r="2" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" />
        """
        element = fromstring(src)
        row = ws._get_row(element, max_col=10)
        row = tuple(row)
        assert len(row) == 10


    def test_get_empty_cells_nonempty_row(self, DummyWorkbook, ReadOnlyWorksheet):
        """Fix for issue #908.

        Get row slice which only contains empty cells in a row containing non-empty
        cells earlier in the row.
        """

        src = b"""
        <sheetData  xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" >
        <row r="1" spans="4:27">
          <c r="A4">
            <v>1</v>
          </c>
        </row>
        </sheetData>
        """

        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])

        xml = fromstring(src)

        min_col = 8
        max_col = 9
        row = tuple(ws._get_row(xml, min_col=min_col, max_col=max_col))

        assert len(row) == 2
        assert all(cell is EMPTY_CELL for cell in row)
        values = [cell.value for cell in row]
        assert values == [None, None]



    def test_read_without_coordinates(self, DummyWorkbook, ReadOnlyWorksheet):

        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", ["Whatever"]*10)
        src = """
        <row xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <c t="s">
            <v>2</v>
          </c>
          <c t="s">
            <v>4</v>
          </c>
          <c t="s">
            <v>3</v>
          </c>
          <c t="s">
            <v>6</v>
          </c>
          <c t="s">
            <v>9</v>
          </c>
        </row>
        """

        element = fromstring(src)
        row = tuple(ws._get_row(element, min_col=1, max_col=None, row_counter=1))
        assert row[0].value == "Whatever"

    @pytest.mark.parametrize("row, column",
                             [
                                 (2, 1),
                                 (3, 1),
                                 (5, 1),
                             ]
                             )
    def test_read_cell_from_empty_row(self, DummyWorkbook, ReadOnlyWorksheet, row, column):
        src = b"""<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="2" />
          <row r="4" />
        </sheetData>
        </worksheet>
        """

        ws = ReadOnlyWorksheet(DummyWorkbook, "Sheet", "", "", [])
        ws._xml = BytesIO(src)
        cell = ws._get_cell(row, column)
        assert cell is EMPTY_CELL
