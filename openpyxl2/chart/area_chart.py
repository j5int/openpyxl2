from __future__ import absolute_import
#Autogenerated schema
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Set,
    Bool,
    Integer,
    Sequence,
    Alias,
)

from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedMinMax,
    NestedSet,
    NestedBool,
)

from ._chart import ChartBase
from .descriptors import NestedShapeProperties, NestedGapAmount
from .axis import CatAx, ValAx, SerAx
from .label import DataLabels
from .series import Series


class _AreaChartBase(ChartBase):

    grouping = NestedSet(values=(['percentStacked', 'standard', 'stacked']))
    varyColors = NestedBool(nested=True, allow_none=True)
    ser = Typed(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")
    dropLines = NestedShapeProperties()

    _series_type = "area"

    __elements__ = ('grouping', 'varyColors', 'ser', 'dLbls', 'dropLines')

    def __init__(self,
                 grouping="standard",
                 varyColors=None,
                 ser=None,
                 dLbls=None,
                 dropLines=None,
                ):
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        super(_AreaChartBase, self).__init__()


class AreaChart(_AreaChartBase):

    tagname = "areaChart"

    grouping = _AreaChartBase.grouping
    varyColors = _AreaChartBase.varyColors
    ser = _AreaChartBase.ser
    dLbls = _AreaChartBase.dLbls
    dropLines = _AreaChartBase.dropLines

    # chart properties actually used by containing classes
    x_axis = Typed(expected_type=CatAx)
    y_axis = Typed(expected_type=ValAx)

    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _AreaChartBase.__elements__ + ('axId',)

    def __init__(self,
                 axId=None,
                 extLst=None,
                 **kw
                ):
        self.x_axis = CatAx()
        self.y_axis = ValAx()
        super(AreaChart, self).__init__(**kw)


class AreaChart3D(AreaChart):

    tagname = "area3DChart"

    grouping = _AreaChartBase.grouping
    varyColors = _AreaChartBase.varyColors
    ser = _AreaChartBase.ser
    dLbls = _AreaChartBase.dLbls
    dropLines = _AreaChartBase.dropLines

    gapDepth = NestedGapAmount()

    x_axis = Typed(expected_type=CatAx)
    y_axis = Typed(expected_type=ValAx)
    z_axis = Typed(expected_type=SerAx, allow_none=True)

    __elements__ = AreaChart.__elements__ + ('gapDepth', )

    def __init__(self, gapDepth=None, **kw):
        self.gapDepth = gapDepth
        super(AreaChart3D, self).__init__(**kw)
        self.x_axis = CatAx()
        self.y_axis = ValAx()
        self.z_axis = SerAx()
