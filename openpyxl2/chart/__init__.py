from __future__ import absolute_import

from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, DoughnutChart, ProjectedPieChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D


def Series(values=None, xvalues=None, name_ref=None, label=None,
           label_from_data=None):
    from .data_source import NumDataSource, NumRef, AxDataSource
    from .series import Series, XYSeries, SeriesLabel, StrRef
    from ..utils import SHEETRANGE_RE, cells_from_range, quote_sheetname
    from itertools import chain

    """
    High level function for creating series

    See http://exceluser.com/excel_help/functions/function-series.htm for a description
    """
    if label_from_data:
        m = SHEETRANGE_RE.match(values)
        sheetname = m.group('notquoted') or m.group('quotes')
        cells = m.group('cells')
        cells = cells_from_range(cells)
        cells = tuple(chain.from_iterable(cells))
        label = "{0}!{1}".format(quote_sheetname(sheetname), cells[0])
        values = "{0}!{1}:{2}".format(quote_sheetname(sheetname), cells[1], cells[-1])
        label = SeriesLabel(strRef=StrRef(label))
    elif label is not None:
        label = SeriesLabel(v=label)

    source = NumDataSource(numRef=NumRef(f=values))
    if xvalues is not None:
        series = XYSeries()
        series.yVal = source
        series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
    else:
        series = Series()
        series.val = source

    if label is not None:
        series.label = label
    return series
