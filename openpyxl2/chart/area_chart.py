
#Autogenerated schema
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Set,
    Bool,
    Integer,
)

from openpyxl2.descriptors.excel import ExtensionList
from .chartBase import ChartLines, GapAmount


class Grouping(Serialisable):

    val = Set(values=(['percentStacked', 'standard', 'stacked']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class _AreaChartBase(Serialisable):

    grouping = Typed(expected_type=Grouping, allow_none=True)
    varyColors = Bool(nested=True, allow_none=True)
    ser = Typed(expected_type=AreaSer, allow_none=True)
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    dropLines = Typed(expected_type=ChartLines, allow_none=True)

    __elements__ = ('grouping', 'varyColors', 'ser', 'dLbls', 'dropLines')

    def __init__(self,
                 grouping=None,
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


class AreaChart(_AreaChartBase):

    tagname = "areaChart"

    gapDepth = Typed(expected_type=GapAmount, allow_none=True, nested=True)
    axId = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _AreaChartBase.__elements__ + ('gapDepth', 'axId', 'extLst')

    def __init__(self,
                 axId=None,
                 extLst=None,
                ):
        self.axId = axId
        self.extLst = extLst


class AreaChart3D(AreaChart):

    tagname = "area3DChart"
