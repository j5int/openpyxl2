from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    NoneSet,
    Float,
    Typed,
    Alias,
)

from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import (
    NestedNoneSet,
    NestedSet,
    NestedMinMax,
)

class ManualLayout(Serialisable):

    tagname = "manualLayout"

    layoutTarget = NestedNoneSet(values=(['inner', 'outer']))
    xMode = NestedSet(values=(['edge', 'factor']))
    yMode = NestedSet(values=(['edge', 'factor']))
    wMode = NestedSet(values=(['edge', 'factor']))
    hMode = NestedSet(values=(['edge', 'factor']))
    x = NestedMinMax(min=-1, max=1)
    y = NestedMinMax(min=-1, max=1)
    w = NestedMinMax(min=0, max=1)
    width = Alias('w')
    h = NestedMinMax(min=0, max=1)
    height = Alias('h')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layoutTarget', 'xMode', 'yMode', 'wMode', 'hMode', 'x',
                    'y', 'w', 'h')

    def __init__(self,
                 layoutTarget=None,
                 xMode=None,
                 yMode=None,
                 wMode=None,
                 hMode=None,
                 x=None,
                 y=None,
                 w=None,
                 h=None,
                 extLst=None,
                ):
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h


class Layout(Serialisable):

    tagname = "layout"

    manualLayout = Typed(expected_type=ManualLayout, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('manualLayout',)

    def __init__(self,
                 manualLayout=None,
                 extLst=None,
                ):
        self.manualLayout = manualLayout
