from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from copy import copy

from openpyxl2.xml.functions import tostring, fromstring
from openpyxl2.tests.helper import compare_xml

import pytest
from openpyxl2.styles import Border, Side
from ..cell_range import CellRange
from ..worksheet import Worksheet
from openpyxl2 import Workbook


@pytest.fixture
def MergeCell():
    from ..merge import MergeCell
    return MergeCell


class TestMergeCell:


    def test_ctor(self, MergeCell):
        cell = MergeCell("A1")
        node = cell.to_tree()
        xml = tostring(node)
        expected = "<mergeCell ref='A1' />"
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, MergeCell):
        xml = "<mergeCell ref='A1' />"
        node = fromstring(xml)
        cell = MergeCell.from_tree(node)
        assert cell == CellRange("A1")


    def test_copy(self, MergeCell):
        cell = MergeCell("A1")
        cp = copy(cell)
        assert cp == cell


@pytest.fixture
def MergedCellRange():
    from ..merge import MergedCellRange
    return MergedCellRange


@pytest.fixture
def default_border():
    return Side(border_style=None, color=None)

@pytest.fixture
def thin_border():
    return Side(border_style="thin", color="000000")

@pytest.fixture
def double_border():
    return Side(border_style="double", color="000000")

@pytest.fixture
def thick_border():
    return Side(border_style="thick", color="000000")

@pytest.fixture
def start_border():
    return Border(top=thick_border(), left=thick_border(),
                  right=thin_border(), bottom=double_border())


class TestMergedCellRange:


    def test_ctor(self, MergedCellRange):
        ws = Worksheet(Workbook())
        cells = MergedCellRange(ws, "A1:E4")
        assert cells.start_cell == ws['A1']


    @pytest.mark.parametrize("end",
                             [
                                 ("C1"),
                                 ("A3"),
                                 ("C3"),
                             ]
                             )

    def test_get_borders(self,  MergedCellRange, end):
        ws = Worksheet(Workbook())
        ws['A1'].border = Border(top=thick_border(), left=thick_border())
        ws[end].border = Border(right=thin_border(), bottom=double_border())

        mcr = MergedCellRange(ws, 'A1:' + end)
        assert mcr.start_cell.coordinate == 'A1'
        assert mcr.start_cell.border == start_border()


    @pytest.mark.parametrize("edge_name, edge, border",
                             [
                    ('top', [(1,1),(1,2), (1,3)],
                     Border(top=thick_border())),
                    ('bottom', [(3,1), (3,2), (3,3)],
                     Border(bottom=double_border())),
                    ('left', [(1,1), (2,1), (3,1)],
                     Border(left=thick_border())),
                    ('right', [(1,3), (2,3), (3,3)],
                     Border(right=thin_border())),
                ]
                )


    def test_side_for_edge(self, MergedCellRange, edge_name, edge, border):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C3')
        mcr.start_cell.border = start_border()

        e, b = mcr._side_for_edge(edge_name,)
        assert e == edge
        assert b == border


    def test_format_1x3(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C1')
        mcr.start_cell.border = start_border()

        mcr.format()

        b1_border = Border(
            top=thick_border(),
                left=default_border(),
                right=default_border(),
                bottom=double_border())
        assert ws['B1'].border == b1_border

        c1_border = Border(
            top=thick_border(),
                left=default_border(),
                right=thin_border(),
                bottom=double_border())
        assert ws['C1'].border == c1_border


    def test_format_3x1(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:A3')
        mcr.start_cell.border = start_border()

        mcr.format()

        a2_border = Border(
            top=default_border(),
                left=thick_border(),
                right=thin_border(),
                bottom=default_border())
        assert ws['A2'].border == a2_border

        a3_border = Border(
            top=default_border(),
                left=thick_border(),
                right=thin_border(),
                bottom=double_border())
        assert ws['A3'].border == a3_border


    def test_format_3x3(self, MergedCellRange):
        ws = Worksheet(Workbook())
        mcr = MergedCellRange(ws, 'A1:C3')
        mcr.start_cell.border = start_border()

        mcr.format()

        for coord in mcr._top:
            cell = ws._cells.get(coord)
            assert cell.border.top == thick_border()

        for coord in mcr._bottom:
            cell = ws._cells.get(coord)
            assert cell.border.bottom== double_border()

        for coord in mcr._left:
            cell = ws._cells.get(coord)
            assert cell.border.left == thick_border()

        for coord in mcr._right:
            cell = ws._cells.get(coord)
            assert cell.border.right == thin_border()

        b2_border = Border(
            top=default_border(),
                left=default_border(),
                right=default_border(),
                bottom=default_border())
        assert ws['B2'].border == b2_border
