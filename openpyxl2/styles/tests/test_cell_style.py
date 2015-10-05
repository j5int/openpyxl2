from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def CellStyle():
    from ..cell_style import CellStyle
    return CellStyle


class TestCellStyle:

    def test_ctor(self, CellStyle):
        cell_style = CellStyle(xfId=0)
        xml = tostring(cell_style.to_tree())
        expected = """
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CellStyle):
        from ..alignment import Alignment
        src = """
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
          <alignment horizontal="center"/>
        </xf>
        """
        node = fromstring(src)
        cell_style = CellStyle.from_tree(node)
        assert cell_style == CellStyle(
            alignment=Alignment(horizontal="center"),
            applyAlignment=True,
            xfId=0,
        )


@pytest.fixture
def CellStyleList():
    from ..cell_style import CellStyleList
    return CellStyleList


class TestCellStyleList:

    def test_ctor(self, CellStyleList):
        cell_style = CellStyleList()
        xml = tostring(cell_style.to_tree())
        expected = """
        <cellXfs />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, CellStyleList):
        src = """
        <cellXfs count="0" />
        """
        node = fromstring(src)
        cell_style = CellStyleList.from_tree(node)
        assert cell_style == CellStyleList()


    def test_to_array(self, CellStyleList):
        src = """
        <cellXfs count="29">
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
            <xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="7" fillId="0" borderId="0" xfId="0" applyFont="1"/>
            <xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/>
            <xf numFmtId="0" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="left"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="right"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment vertical="top"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"/>
            <xf numFmtId="2" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="2" xfId="0" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment wrapText="1"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment shrinkToFit="1"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1" applyBorder="1"/>
            <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">
              <alignment horizontal="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="3" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="4" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="5" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="0" fillId="4" borderId="6" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
              <alignment horizontal="center" vertical="center"/>
            </xf>
            <xf numFmtId="0" fontId="6" fillId="5" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
        </cellXfs>
        """
        node = fromstring(src)
        xfs = CellStyleList.from_tree(node)
        xfs._to_array()
        assert len(xfs.styles) == 29
        assert len(xfs.alignments) == 8
        assert len(xfs.prots) == 0
