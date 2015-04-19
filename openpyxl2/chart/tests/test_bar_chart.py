from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import tostring, fromstring
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def BarChart():
    from ..bar_chart import BarChart
    return BarChart


class TestBarChart:


    def test_ctor(self, BarChart):
        bc = BarChart()
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
          <barDir val="col" />
          <grouping val="clustered" />

          <axId val="60871424" />
          <axId val="60873344" />
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_tree(self, BarChart):
        src = """
            <c:barChart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
                <c:barDir val="col"/>
                <c:grouping val="clustered"/>
                <c:varyColors val="0"/>
                <c:ser>
                    <c:idx val="0"/>
                    <c:order val="0"/>
                    <c:invertIfNegative val="0"/>
                    <c:val>
                        <c:numRef>
                            <c:f>Blatt1!$A$1:$A$12</c:f>
                          </c:numRef>
                    </c:val>
                </c:ser>
                <c:dLbls>
                    <c:showLegendKey val="0"/>
                    <c:showVal val="0"/>
                    <c:showCatName val="0"/>
                    <c:showSerName val="0"/>
                    <c:showPercent val="0"/>
                    <c:showBubbleSize val="0"/>
                </c:dLbls>
                <c:gapWidth val="150"/>
                <c:axId val="2098063848"/>
                <c:axId val="2098059592"/>
            </c:barChart>
        """
        node = fromstring(src)
        bc = BarChart.from_tree(node)
        assert bc.grouping == "clustered"
        assert len(bc.ser) == 1
        assert bc.dLbls is not None

        # check roundtripping
        xml = tostring(bc.to_tree())
        expected = """
        <barChart>
        <barDir val="col"></barDir>
        <grouping val="clustered"></grouping>
        <ser>
          <val>
            <numRef>
              <f>Blatt1!$A$1:$A$12</f>
            </numRef>
          </val>
        </ser>
        <dLbls></dLbls>
        <gapWidth val="150"></gapWidth>
        <axId val="2098063848"></axId>
        <axId val="2098059592"></axId>
        </barChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_serialise(self, BarChart):

        bc = BarChart()
