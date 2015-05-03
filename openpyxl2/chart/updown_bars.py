"""
Collection of utility primitives for charts.
"""

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import Typed
from openpyxl2.descriptors.excel import ExtensionList

from .shapes import ShapeProperties
from .descriptors import (
    NestedGapAmount,
    NestedOverlap,
    NestedShapeProperties
)


class UpDownBars(Serialisable):

    tagname = "upbars"

    gapWidth = NestedGapAmount()
    upBars = NestedShapeProperties()
    downBars = NestedShapeProperties()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars')

    def __init__(self,
                 gapWidth=None,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
