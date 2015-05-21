from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Sequence,
    Typed,
    Alias,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedSet
)

from ._chart import ChartBase
from .axis import CatAx, ValAx
from .series import Series
from .label import DataLabels


class RadarChart(ChartBase):

    tagname = "radarChart"

    radarStyle = NestedSet(values=(['standard', 'marker', 'filled']))
    style = Alias("radarStyle")
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Typed(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    _series_type = "radar"

    x_axis = Typed(expected_type=CatAx)
    y_axis = Typed(expected_type=ValAx)

    __elements__ = ('radarStyle', 'varyColors', 'ser', 'dLbls', 'axId')

    def __init__(self,
                 radarStyle="standard",
                 varyColors=None,
                 ser=None,
                 dLbls=None,
                 axId=None,
                 extLst=None,
                ):
        self.radarStyle = radarStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.x_axis = CatAx()
        self.y_axis = ValAx()
        super(RadarChart, self).__init__()

