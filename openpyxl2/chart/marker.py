from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.compat import safe_string
from openpyxl2.xml.functions import Element

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Alias,
)

from openpyxl2.descriptors.excel import(
    ExtensionList,
)

from openpyxl2.descriptors.nested import (
    NestedBool,
    NestedInteger,
    NestedMinMax,
    NestedNoneSet,
)

from .layout import Layout
from .picture import PictureOptions
from .shapes import *
from .text import *
from .error_bar import *


def _marker_symbol(tagname, value, namespace=None):
    """
    Override serialisation because explicit none required
    """
    if namespace is not None:
        tagname = "{%s}%s" % (namespace, tagname)
    return Element(tagname, val=safe_string(value))


class Marker(Serialisable):

    tagname = "marker"

    symbol = NestedNoneSet(values=(['circle', 'dash', 'diamond', 'dot', 'picture',
                              'plus', 'square', 'star', 'triangle', 'x', 'auto']),
                           to_tree=_marker_symbol)
    size = NestedMinMax(min=2, max=72, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    ShapeProperties = Alias('sprPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('symbol', 'size', 'spPr')

    def __init__(self,
                 symbol=None,
                 size=None,
                 spPr=None,
                 extLst=None,
                ):
        self.symbol = symbol
        self.size = size
        self.spPr = spPr


class DataPoint(Serialisable):

    tagname = "dPt"

    idx = NestedInteger()
    invertIfNegative = NestedBool(allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    bubble3D = NestedBool(allow_none=True)
    explosion = Integer(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    shapeProperties = Alias('spPr')
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'invertIfNegative', 'marker', 'bubble3D',
                    'explosion', 'spPr', 'pictureOptions')

    def __init__(self,
                 idx=None,
                 invertIfNegative=None,
                 marker=None,
                 bubble3D=None,
                 explosion=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.idx = idx
        self.invertIfNegative = invertIfNegative
        self.marker = marker
        self.bubble3D = bubble3D
        self.explosion = explosion
        if spPr is None:
            spPr = ShapeProperties()
        self.spPr = spPr
        self.pictureOptions = pictureOptions
