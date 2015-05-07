from __future__ import absolute_import

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Sequence,
    Alias
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedNoneSet,
    NestedBool,
)

from ._chart import ChartBase
from .axis import ValAx
from .series import XYSeries
from .label import DataLabels


class ScatterChart(ChartBase):

    tagname = "scatterChart"

    scatterStyle = NestedNoneSet(values=(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']))
    varyColors = NestedBool(allow_none=True)
    ser = Typed(expected_type=XYSeries, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    x_axis = Typed(expected_type=ValAx)
    y_axis = Typed(expected_type=ValAx)

    _series_type = "scatter"

    __elements__ = ('scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId',)

    def __init__(self,
                 scatterStyle=None,
                 varyColors=None,
                 ser=None,
                 dLbls=None,
                 axId=None,
                 extLst=None,
                ):
        self.scatterStyle = scatterStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.x_axis = ValAx(axId=10, crossAx=20)
        self.y_axis = ValAx(axId=20, crossAx=10)
        super(ScatterChart, self).__init__()
