from __future__ import absolute_import
#Autogenerated schema
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Sequence,
    Alias,
    )
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedSet,
    NestedBool,
    NestedMinMax,
)

from ._chart import ChartBase
from .updown_bars import UpDownBars
from .descriptors import NestedGapAmount, NestedShapeProperties
from .axis import AxId
from .label import DataLabels
from .series import Series


class _LineChartBase(ChartBase):

    grouping = NestedSet(values=(['percentStacked', 'standard', 'stacked']))
    varyColors = NestedBool(allow_none=True)
    ser = Sequence(expected_type=Series, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")
    dropLines = NestedShapeProperties()

    _series_type = "line"

    __elements__ = ('grouping', 'varyColors', 'ser', 'dLbls', 'dropLines')

    def __init__(self,
                 grouping="standard",
                 varyColors=None,
                 ser=[],
                 dLbls=None,
                 dropLines=None,
                ):
        self.grouping = grouping
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines


class LineChart(_LineChartBase):

    tagname = "lineChart"

    grouping = _LineChartBase.grouping
    varyColors = _LineChartBase.varyColors
    ser = _LineChartBase.ser
    dLbls = _LineChartBase.dLbls
    dropLines =_LineChartBase.dropLines

    hiLowLines = NestedShapeProperties()
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    marker = NestedBool(allow_none=True)
    smooth = NestedBool(allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _LineChartBase.__elements__ + ('hiLowLines', 'upDownBars', 'marker', 'smooth', 'axId')

    def __init__(self,
                 hiLowLines=None,
                 upDownBars=None,
                 marker=None,
                 smooth=None,
                 axId=None,
                 extLst=None,
                 **kw
                ):
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.marker = marker
        self.smooth = smooth
        if axId is None:
            axId = (AxId(10), AxId(100))
        self.axId = axId
        super(LineChart, self).__init__(**kw)


class LineChart3D(_LineChartBase):

    tagname = "line3DChart"

    grouping = _LineChartBase.grouping
    varyColors = _LineChartBase.varyColors
    ser = _LineChartBase.ser
    dLbls = _LineChartBase.dLbls
    dropLines =_LineChartBase.dropLines

    gapDepth = NestedGapAmount()
    hiLowLines = NestedShapeProperties()
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    marker = NestedBool(allow_none=True)
    smooth = NestedBool(allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _LineChartBase.__elements__ + ('gapDepth', 'hiLowLines',
                                                  'upDownBars', 'marker', 'smooth', 'axId')

    def __init__(self,
                 gapDepth=None,
                 hiLowLines=None,
                 upDownBars=None,
                 marker=None,
                 smooth=None,
                 axId=None,
                 **kw
                ):
        self.gapDepth = gapDepth
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.marker = marker
        self.smooth = smooth
        if axId is None:
            axId = (AxId(10), AxId(100))
        self.axId = axId
        super(LineChart3D, self).__init__(**kw)
