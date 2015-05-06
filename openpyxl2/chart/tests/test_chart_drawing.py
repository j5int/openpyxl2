from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def TwoCellAnchor():
    from ..chart_drawing import TwoCellAnchor
    return TwoCellAnchor


class TestTwoCellAnchor:

    def test_ctor(self, TwoCellAnchor):
        chart_drawing = TwoCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData />
        </twoCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TwoCellAnchor):
        src = """
        <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData></clientData>
         </twoCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = TwoCellAnchor.from_tree(node)
        assert chart_drawing == TwoCellAnchor()


@pytest.fixture
def OneCellAnchor():
    from ..chart_drawing import OneCellAnchor
    return OneCellAnchor


class TestOneCellAnchor:

    def test_ctor(self, OneCellAnchor):
        chart_drawing = OneCellAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="0" cy="0" />
          <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, OneCellAnchor):
        src = """
        <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="0" cy="0" />
          <clientData></clientData>
        </oneCellAnchor>
        """
        node = fromstring(src)
        chart_drawing = OneCellAnchor.from_tree(node)
        assert chart_drawing == OneCellAnchor()


@pytest.fixture
def AbsoluteAnchor():
    from ..chart_drawing import AbsoluteAnchor
    return AbsoluteAnchor


class TestAbsoluteAnchor:

    def test_ctor(self, AbsoluteAnchor):
        chart_drawing = AbsoluteAnchor()
        xml = tostring(chart_drawing.to_tree())
        expected = """
         <absoluteAnchor>
           <pos x="0" y="0" />
           <ext cx="0" cy="0" />
           <clientData></clientData>
         </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, AbsoluteAnchor):
        src = """
         <absoluteAnchor>
           <pos x="0" y="0" />
           <ext cx="0" cy="0" />
           <clientData></clientData>
         </absoluteAnchor>
         """
        node = fromstring(src)
        chart_drawing = AbsoluteAnchor.from_tree(node)
        assert chart_drawing == AbsoluteAnchor()


@pytest.fixture
def SpreadsheetDrawing():
    from ..chart_drawing import SpreadsheetDrawing
    return SpreadsheetDrawing


class TestSpreadsheetDrawing:

    def test_ctor(self, SpreadsheetDrawing):
        from ..chart_drawing import (
            OneCellAnchor,
            TwoCellAnchor,
            AbsoluteAnchor
        )
        a = [AbsoluteAnchor(), AbsoluteAnchor()]
        o = [OneCellAnchor()]
        t = [TwoCellAnchor(), TwoCellAnchor()]
        chart_drawing = SpreadsheetDrawing(absoluteAnchor=a, oneCellAnchor=o,
                                           twoCellAnchor=t)
        xml = tostring(chart_drawing.to_tree())
        expected = """
        <wsDr>
          <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
          <clientData></clientData>
          </twoCellAnchor>
          <twoCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
          <to>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </to>
            <clientData></clientData>
          </twoCellAnchor>
          <oneCellAnchor>
          <from>
            <col>0</col>
            <colOff>0</colOff>
            <row>0</row>
            <rowOff>0</rowOff>
          </from>
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </oneCellAnchor>
          <absoluteAnchor>
            <pos x="0" y="0"  />
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </absoluteAnchor>
          <absoluteAnchor>
            <pos x="0" y="0" />
            <ext cx="0" cy="0" />
            <clientData></clientData>
          </absoluteAnchor>
        </wsDr>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_write(self, SpreadsheetDrawing):

        class Chart:

            anchor = "E15"

        drawing = SpreadsheetDrawing()
        drawing.charts.append(Chart())
        xml = tostring(drawing._write())
        expected = """
        <wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
           xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
        <oneCellAnchor>
          <from>
            <col>4</col>
            <colOff>0</colOff>
            <row>14</row>
            <rowOff>0</rowOff>
          </from>
          <ext cx="5400000" cy="2700000"/>
          <graphicFrame>
            <nvGraphicFramePr>
              <cNvPr id="1" name="Chart 1"/>
              <cNvGraphicFramePr/>
            </nvGraphicFramePr>
            <xfrm/>
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
              </a:graphicData>
            </a:graphic>
          </graphicFrame>
          <clientData/>
        </oneCellAnchor>
        </wsDr>
        """
        diff = compare_xml (xml, expected)
        assert diff is None, diff
