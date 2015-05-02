"""
Collection of utility primitives for charts.
"""

from openpyxl2.compat import basestring
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax,
    Alias,
    String,
    Integer,
    Sequence,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedString,
    NestedText,
    NestedInteger,
)

from .shapes import ShapeProperties


class GapAmount(Serialisable):

    # needs to serialise to %
    val = MinMax(min=0, max=500)

    def __init__(self,
                 val=150,
                ):
        self.val = val


class Overlap(Serialisable):

    # needs to serialise to %

    val = MinMax(min=0, max=150)

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ChartLines(Serialisable):

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr


class UpDownBar(Serialisable):

    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('spPr',)

    def __init__(self,
                 spPr=None,
                ):
        self.spPr = spPr


class UpDownBars(Serialisable):

    gapWidth = Typed(expected_type=GapAmount, allow_none=True)
    upBars = Typed(expected_type=UpDownBar, allow_none=True)
    downBars = Typed(expected_type=UpDownBar, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars', 'extLst')

    def __init__(self,
                 gapWidth=None,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
        self.extLst = extLst
