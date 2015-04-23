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

from .axis import AxId
from .series import ScatterSer
from .label import DataLabels


class ScatterChart(Serialisable):

    tagname = "scatterChart"

    scatterStyle = NestedNoneSet(values=(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']))
    varyColors = NestedBool(allow_none=True)
    ser = Typed(expected_type=ScatterSer, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    dataLabels = Alias("dLbls")
    axId = Sequence(expected_type=AxId)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

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
        if axId is None:
            axId = [AxId(10), AxId(100)]
        self.axId = axId
