
from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def StockChart():
    from ..stock_chart import StockChart
    return StockChart


class TestStockChart:

    def test_ctor(self, StockChart):
        from openpyxl2.chart.series import LineSer

        chart = StockChart(ser=[LineSer(), LineSer(), LineSer()])
        xml = tostring(chart.to_tree())
        expected = """
        <stockChart>
          <ser></ser>
          <ser></ser>
          <ser></ser>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </stockChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, StockChart):
        src = """
        <stockChart>
          <ser></ser>
          <ser></ser>
          <ser></ser>
          <dLbls></dLbls>
          <hiLowLines/>
          <axId val="2109232808"/>
          <axId val="2108950264"/>
        </stockChart>
        """
        node = fromstring(src)
        chart = StockChart.from_tree(node)
        assert dict(chart) == {}
