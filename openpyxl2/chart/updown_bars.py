from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import Typed
from openpyxl2.descriptors.excel import ExtensionList

from .shapes import GraphicalProperties
from .axis import ChartLines
from .descriptors import NestedGapAmount


class UpDownBars(Serialisable):

    tagname = "upbars"

    gapWidth = NestedGapAmount()
    upBars = Typed(expected_type=ChartLines, allow_none=True)
    downBars = Typed(expected_type=ChartLines, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('gapWidth', 'upBars', 'downBars')

    def __init__(self,
                 gapWidth=150,
                 upBars=None,
                 downBars=None,
                 extLst=None,
                ):
        self.gapWidth = gapWidth
        self.upBars = upBars
        self.downBars = downBars
