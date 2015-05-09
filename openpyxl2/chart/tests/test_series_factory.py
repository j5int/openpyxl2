from __future__ import absolute_import

import pytest

from openpyxl2.xml.functions import tostring, fromstring
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def Series():
    from ..series_factory import SeriesFactory
    return SeriesFactory


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
           <f>Sheet1!$A$2:$A$10</f>
          </numRef>
        </val>
        </ser>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_xy(self, Series):
        from ..series import XYSeries
        series = Series("Sheet!A1:A10", xvalues="Sheet!B1:B10")
        assert isinstance(series, XYSeries)
