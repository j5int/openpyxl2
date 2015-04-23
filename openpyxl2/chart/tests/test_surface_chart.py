
from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def LineChart():
    from ..surface_chart import LineChart
    return LineChart


class TestLineChart:

    def test_ctor(self, LineChart):
        surface_chart = LineChart()
        xml = tostring(surface_chart.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, LineChart):
        src = """
        <root />
        """
        node = fromstring(src)
        surface_chart = LineChart.from_tree(node)
        assert dict(surface_chart) == {}

