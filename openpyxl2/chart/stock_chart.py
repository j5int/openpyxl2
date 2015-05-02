from __future__ import absolute_import

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Sequence,
    Alias,
)
from openpyxl2.descriptors.excel import ExtensionList

from .axis import AxId
from .chartBase import UpDownBars, ChartLines
from .label import DataLabels
from .series import Series


class StockChart(Serialisable):

    tagname = "stockChart"

    ser = Sequence(expected_type=Series) #min 3, max4
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dropLines = Typed(expected_type=ChartLines, allow_none=True)
    hiLowLines = Typed(expected_type=ChartLines, allow_none=True)
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars',
                    'axId')

    def __init__(self,
                 ser=None,
                 dLbls=None,
                 dropLines=None,
                 hiLowLines=None,
                 upDownBars=None,
                 axId=None,
                 extLst=None,
                ):
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        if axId is None:
            axId = (AxId(10), AxId(100))
        self.axId = axId
