from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl2.xml.constants import SHEET_DRAWING_NS
from openpyxl2.xml.functions import Element, fromstring, tostring

from openpyxl2.tests.helper import compare_xml


class TestShadow(object):

    def setup(self):
        from ..drawing import Shadow
        self.shadow = Shadow()

    def test_ctor(self):
        s = self.shadow
        assert s.visible == False
        assert s.blurRadius == 6
        assert s.distance == 2
        assert s.direction == 0
        assert s.alignment == "br"
        assert s.color.index == "00000000"
        assert s.alpha == 50


class TestDrawing(object):

    def setup(self):
        from ..drawing import Drawing
        self.drawing = Drawing()

    def test_ctor(self):
        d = self.drawing
        assert d.coordinates == ((1, 2), (16, 8))
        assert d.width == 21
        assert d.height == 192
        assert d.left == 0
        assert d.top == 0
        assert d.count == 0
        assert d.rotation == 0
        assert d.resize_proportional is False
        assert d.description == ""
        assert d.name == ""

    def test_width(self):
        d = self.drawing
        d.width = 100
        d.height = 50
        assert d.width == 100

    def test_proportional_width(self):
        d = self.drawing
        d.resize_proportional = True
        d.width = 100
        d.height = 50
        assert (d.width, d.height) == (5, 50)

    def test_height(self):
        d = self.drawing
        d.height = 50
        d.width = 100
        assert d.height == 50

    def test_proportional_height(self):
        d = self.drawing
        d.resize_proportional = True
        d.height = 50
        d.width = 100
        assert (d.width, d.height) == (100, 1000)

    def test_set_dimension(self):
        d = self.drawing
        d.resize_proportional = True
        d.set_dimension(100, 50)
        assert d.width == 6
        assert d.height == 50

        d.set_dimension(50, 500)
        assert d.width == 50
        assert d.height == 417

    def test_get_emu(self):
        d = self.drawing
        dims = d.get_emu_dimensions()
        assert dims == (0, 0, 200025, 1828800)


class DummyDrawing(object):

    """Shapes need charts which need drawings"""

    width = 10
    height = 20


class DummyChart(object):

    """Shapes need a chart to calculate their coordinates"""

    width = 100
    height = 100

    def __init__(self):
        self.drawing = DummyDrawing()

    def _get_margin_left(self):
        return 10

    def _get_margin_top(self):
        return 5

    def get_x_units(self):
        return 25

    def get_y_units(self):
        return 15


@pytest.fixture()
def ImageFile(datadir):
    from ..image import Image
    datadir.chdir()
    return Image("plain.png")


class DummySheet:

    pass


class TestDrawingWriter(object):

    def setup(self):
        from ..writer import DrawingWriter
        sheet = DummySheet()
        sheet._charts = []
        sheet._images = []
        self.dw = DrawingWriter(sheet=sheet)

    def test_write(self):
        xml = self.dw.write()
        expected = """<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" />"""
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_chart(self):
        from openpyxl2.drawing import Drawing
        root = Element("{%s}wsDr" % SHEET_DRAWING_NS)
        chart = DummyChart()
        drawing = Drawing()
        chart.drawing = drawing
        node = self.dw._write_chart(root, chart, 1)
        xml = tostring(node.to_tree())
        expected = """
        <absoluteAnchor xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
          <pos x="0" y="0"/>
          <ext cx="200025" cy="1828800"/>
          <graphicFrame>
            <nvGraphicFramePr>
              <cNvPr id="1" name="Chart 1"/>
              <cNvGraphicFramePr/>
            </nvGraphicFramePr>
           <xfrm />
            <a:graphic>
              <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1"/>
              </a:graphicData>
            </a:graphic>
          </graphicFrame>
          <clientData/>
        </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    @pytest.mark.pil_required
    def test_write_images(self, ImageFile):

        root = Element("wsDr")
        node = self.dw._write_image(root, ImageFile, 1)
        xml = tostring(node.to_tree())
        expected = """
        <absoluteAnchor xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
          <pos x="0" y="0"/>
          <ext cx="1123950" cy="1123950"/>
          <pic>
            <nvPicPr>
              <cNvPr descr="Name of file" id="1" name="Image 1"/>
              <cNvPicPr>
              </cNvPicPr>
            </nvPicPr>
            <blipFill>
              <a:blip cstate="print" r:embed="rId1"/>
            </blipFill>
            <spPr>
              <a:noFill/>
              <a:ln w="1">
                <a:noFill/>
              </a:ln>
            </spPr>
          </pic>
          <clientData/>
        </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.pil_required
    def test_write_anchor(self, ImageFile):
        drawing = ImageFile.drawing
        node = drawing.anchor
        xml = tostring(node.to_tree())
        expected = """
        <absoluteAnchor>
            <pos x="0" y="0"/>
            <ext cx="1123950" cy="1123950"/>
            <clientData></clientData>
        </absoluteAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    @pytest.mark.pil_required
    def test_write_anchor_onecell(self, ImageFile):
        drawing = ImageFile.drawing
        drawing.anchortype =  "oneCell"
        node = drawing.anchor
        xml = tostring(node.to_tree())
        expected = """
        <oneCellAnchor>
            <from>
                <col>0</col>
                <colOff>0</colOff>
                <row>0</row>
                <rowOff>0</rowOff>
            </from>
            <ext cx="1123950" cy="1123950"/>
            <clientData></clientData>
        </oneCellAnchor>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff

    def test_write_rels(self):
        self.dw._sheet._charts.append(None)
        self.dw._sheet._images.append(None)
        xml = self.dw.write_rels(1, 1)
        expected = """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
            <Relationship Id="rId1" Target="../charts/chart1.xml"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"/>
            <Relationship Id="rId1" Target="../media/image1.png"
            Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/>
        </Relationships>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff
