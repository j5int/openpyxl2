
from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def LineChart():
    from ..line_chart import LineChart
    return LineChart


class TestLineChart:

    def test_ctor(self, LineChart):
        chart = LineChart()
        xml = tostring(chart.to_tree())
        expected = """
        <lineChart>
          <grouping val="standard"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineChart):
        src = """
        <lineChart>
          <grouping val="stacked"></grouping>
          <axId val="10"></axId>
          <axId val="100"></axId>
        </lineChart>
        """
        node = fromstring(src)
        chart = LineChart.from_tree(node)
        assert dict(chart) == {}
        assert chart.grouping == "stacked"
