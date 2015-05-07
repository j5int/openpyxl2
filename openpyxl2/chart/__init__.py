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
    from .series import Series, XYSeries, SeriesLabel

    """
    High level function for creating series

    See http://exceluser.com/excel_help/functions/function-series.htm for a description
    """
    source = NumDataSource(numRef=NumRef(f=values))
    if xvalues is not None:
        series = XYSeries()
        series.yVal = source
        series.xVal = AxDataSource(numRef=NumRef(f=xvalues))
    else:
        series = Series()
        series.val = source

    if label is not None:
        series.label = SeriesLabel(v=label)
    return series
