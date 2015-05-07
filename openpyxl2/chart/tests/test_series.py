from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import tostring, fromstring
from openpyxl2.tests.helper import compare_xml


class TestBarSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <val>
            <numRef>
                <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['bar']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestAreaSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['area']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestBubbleSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <xVal>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
             </numRef>
          </xVal>
          <yVal>
            <numRef>
              <f>Blatt1!$B$1:$B$12</f>
            </numRef>
          </yVal>
          <bubbleSize>
            <numLit>
              <formatCode>General</formatCode>
              <ptCount val="12"/>
              <pt idx="0">
                <v>1.1</v>
              </pt>
              <pt idx="1">
                <v>1.1</v>
              </pt>
              <pt idx="2">
                <v>1.1</v>
              </pt>
              <pt idx="3">
                <v>1.1</v>
              </pt>
              <pt idx="4">
                <v>1.1</v>
              </pt>
              <pt idx="5">
                <v>1.1</v>
              </pt>
              <pt idx="6">
                <v>1.1</v>
              </pt>
              <pt idx="7">
                <v>1.1</v>
              </pt>
              <pt idx="8">
                <v>1.1</v>
              </pt>
              <pt idx="9">
                <v>1.1</v>
              </pt>
              <pt idx="10">
                <v>1.1</v>
              </pt>
              <pt idx="11">
                <v>1.1</v>
              </pt>
            </numLit>
          </bubbleSize>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.xVal.numRef.ref == 'Blatt1!$A$1:$A$12'
        assert ser.yVal.numRef.ref == 'Blatt1!$B$1:$B$12'
        assert ser.bubbleSize.numLit.ptCount == 12
        assert ser.bubbleSize.numLit.pt[0].v == 1.1

        ser.__elements__ = attribute_mapping['bubble']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestPieSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <explosion val="25"/>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['pie']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff



class TestRadarSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <marker>
            <symbol val="none"/>
          </marker>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['radar']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestScatterSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <marker>
            <symbol val="none"/>
          </marker>
          <xVal>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </xVal>
          <yVal>
            <numRef>
              <f>Blatt1!$B$1:$B$12</f>
            </numRef>
          </yVal>
          <smooth val="0"/>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.xVal.numRef.ref == 'Blatt1!$A$1:$A$12'
        assert ser.yVal.numRef.ref == 'Blatt1!$B$1:$B$12'

        ser.__elements__ = attribute_mapping['scatter']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


class TestSurfaceSer:

    def test_from_tree(self):
        from ..series import Series, attribute_mapping

        src = """
        <ser>
          <idx val="0"/>
          <order val="0"/>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        """
        node = fromstring(src)
        ser = Series.from_tree(node)
        assert ser.idx == 0
        assert ser.order == 0
        assert ser.val.numRef.ref == 'Blatt1!$A$1:$A$12'

        ser.__elements__ = attribute_mapping['surface']
        xml = tostring(ser.to_tree())
        diff = compare_xml(xml, src)
        assert diff is None, diff


@pytest.fixture
def Series():
    from .. import Series
    return Series

class TestSeriesFactory:

    def test_ctor(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        series.__elements__ = ('idx', 'order', 'val')
        xml = tostring(series.to_tree())
        expected = """
        <ser>
          <idx val="0"></idx>
          <order val="0"></order>
          <val>
            <numRef>
              <f>Sheet1!$A$1:$A$10</f>
            </numRef>
          </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_manual_idx(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        series.__elements__ = ('idx', 'order', 'val')
        xml = tostring(series.to_tree(idx=5))
        expected = """
        <ser>
          <idx val="5"></idx>
          <order val="5"></order>
          <val>
            <numRef>
              <f>Sheet1!$A$1:$A$10</f>
            </numRef>
          </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_manual_order(self, Series):
        series = Series(values="Sheet1!$A$1:$A$10")
        series.order = 2
        series.__elements__ = ('idx', 'order', 'val')
        xml = tostring(series.to_tree(idx=5))
        expected = """
        <ser>
          <idx val="5"></idx>
          <order val="2"></order>
          <val>
            <numRef>
              <f>Sheet1!$A$1:$A$10</f>
            </numRef>
          </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_title(self, Series):
        series = Series("Sheet1!A1:A10", title="First Series")
        series.__elements__ = ('idx', 'order', 'tx')
        xml = tostring(series.to_tree(idx=0))
        expected = """
        <ser>
          <idx val="0"></idx>
          <order val="0"></order>
          <tx>
            <v>First Series</v>
          </tx>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_title_from_data(self, Series):
        series = Series("Sheet1!A1:A10", title_from_data=True)
        series.__elements__ = ('tx', 'val')
        xml = tostring(series.to_tree(idx=0))
        expected = """
        <ser>
        <tx>
          <strRef>
            <f>Sheet1!A1</f>
          </strRef>
         </tx>
        <val>
        <numRef>
           <f>Sheet1!A2:A10</f>
          </numRef>
        </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xy(self, Series):
        from ..series import XYSeries
        series = Series("A1:A10", xvalues="B1:B10")
        assert isinstance(series, XYSeries)


    def test_axis_labels(self, Series):
        series = Series("B1:B10", axis_labels="A1:A10")
        series.__elements__ = ('cat', 'val')
        xml = tostring(series.to_tree(idx=0))
        expected = """
            <ser>
            <cat>
            <numRef>
               <f>A1:A10</f>
              </numRef>
            </cat>
            <val>
            <numRef>
               <f>B1:B10</f>
              </numRef>
            </val>
            </ser>
            """
        diff = compare_xml(xml, expected)
        assert diff is None, diff



@pytest.fixture
def SeriesLabel():
    from ..series import SeriesLabel
    return SeriesLabel


class TestSeriesLabel:

    def test_ctor(self, SeriesLabel):
        label = SeriesLabel(v="Label")
        xml = tostring(label.to_tree())
        expected = """
        <tx>
          <v>Label</v>
        </tx>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, SeriesLabel):
        src = """
        <tx>
          <v>Label</v>
        </tx>
        """
        node = fromstring(src)
        label = SeriesLabel.from_tree(node)
        assert label == SeriesLabel(v="Label")
