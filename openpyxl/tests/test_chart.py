# file openpyxl/tests/test_chart.py

# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

from datetime import date
import pytest


@pytest.fixture
def sheet(ten_row_sheet):
    ten_row_sheet.title = "reference"
    return ten_row_sheet


@pytest.fixture
def cell(sheet, Reference):
    return Reference(sheet, (0, 0))


@pytest.fixture
def cell_range(sheet, Reference):
    return Reference(sheet, (0, 0), (9, 0))


@pytest.fixture()
def empty_range(sheet, Reference):
    for i in range(10):
        sheet.cell(row=i, column=1).value = None
    return Reference(sheet, (0, 1), (9, 1))


@pytest.fixture()
def missing_values(sheet, Reference):
    vals = [None, None, 1, 2, 3, 4, 5, 6, 7, 8]
    for idx, val in enumerate(vals):
        sheet.cell(row=idx, column=2).value = val
    return Reference(sheet, (0, 2), (9, 2))


@pytest.fixture
def column_of_letters(sheet, Reference):
    for idx, l in enumerate("ABCDEFGHIJ"):
        sheet.cell(row=idx, column=1).value = l
    return Reference(sheet, (0, 1), (9, 1))


class TestErrorBar(object):

    def test_ctor(self, ErrorBar):
        with pytest.raises(TypeError):
            ErrorBar(None, range(10))


@pytest.fixture()
def series(cell_range, Series):
    return Series(values=cell_range)


class TestChart(object):

    def test_ctor(self, Chart):
        from openpyxl.charts import Legend
        from openpyxl.drawing import Drawing
        c = Chart()
        assert c.TYPE == None
        assert c.GROUPING == "standard"
        assert isinstance(c.legend, Legend)
        assert c.show_legend
        assert c.lang == 'en-GB'
        assert c.title == ''
        assert c.print_margins == {'b':0.75, 'l':0.7, 'r':0.7, 't':0.75,
                                   'header':0.3, 'footer':0.3}
        assert isinstance(c.drawing, Drawing)
        assert c.width == 0.6
        assert c.height == 0.6
        assert c.margin_top == 0.31
        assert c.series == []
        assert c.shapes == []
        with pytest.raises(ValueError):
            assert c.margin_left == 0

    def test_mymax(self, Chart):
        c = Chart()
        assert c.mymax(range(10)) == 9
        from string import ascii_letters as letters
        assert c.mymax(list(letters)) == "z"
        assert c.mymax(range(-10, 1)) == 0
        assert c.mymax([""]*10) == ""

    def test_mymin(self, Chart):
        c = Chart()
        assert c.mymin(range(10)) == 0
        from string import ascii_letters as letters
        assert c.mymin(list(letters)) == "A"
        assert c.mymin(range(-10, 1)) == -10
        assert c.mymin([""]*10) == ""

    def test_margin_top(self, Chart):
        c = Chart()
        assert c.margin_top == 0.31

    def test_margin_left(self, series, Chart):
        c = Chart()
        c.append(series)
        assert c.margin_left == 0.03375

    def test_set_margin_top(self, Chart):
        c = Chart()
        c.margin_top = 1
        assert c.margin_top == 0.31

    def test_set_margin_left(self, series, Chart):
        c = Chart()
        c.append(series)
        c.margin_left = 0
        assert c.margin_left  == 0.03375


class TestGraphChart(object):

    def test_ctor(self, GraphChart, Axis):
        c = GraphChart()
        assert isinstance(c.x_axis, Axis)
        assert isinstance(c.y_axis, Axis)

    def test_get_x_unit(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        assert c.get_x_units() == 10

    def test_get_y_unit(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        c.y_axis.max = 10
        assert c.get_y_units() == 190500

    def test_get_y_char(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        assert c.get_y_chars() == 1

    def test_compute_series_extremes(self, GraphChart, series):
        c = GraphChart()
        c.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 9

    def test_compute_series_max_dates(self, ws, Reference, Series, GraphChart):
        for i in range(1, 10):
            ws.append([date(2013, i, 1)])
        c = GraphChart()
        ref = Reference(ws, (0, 0), (9, 0))
        series = Series(ref)
        c.append(series)
        mini, maxi = c._get_extremes()
        assert mini == 0
        assert maxi == 41518.0

    def test_override_axis(self, GraphChart, series):
        c = GraphChart()
        c.add_serie(series)
        c.compute_axes()
        assert c.y_axis.min == 0
        assert c.y_axis.max == 10
        c.y_axis.min = -1
        c.y_axis.max = 5
        assert c.y_axis.min == -2
        assert c.y_axis.max == 6


class TestLineChart(object):

    def test_ctor(self, LineChart):
        c = LineChart()
        assert c.TYPE == "lineChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"


class TestPieChart(object):

    def test_ctor(self, PieChart):
        c = PieChart()
        assert c.TYPE, "pieChart"


class TestBarChart(object):

    def test_ctor(self, BarChart):
        c = BarChart()
        assert c.TYPE == "barChart"
        assert c.x_axis.type == "catAx"
        assert c.y_axis.type == "valAx"


class TestScatterChart(object):

    def test_ctor(self, ScatterChart):
        c = ScatterChart()
        assert c.TYPE == "scatterChart"
        assert c.x_axis.type == "valAx"
        assert c.y_axis.type == "valAx"
