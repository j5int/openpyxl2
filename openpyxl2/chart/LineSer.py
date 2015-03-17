#Autogenerated schema

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Float,
    MinMax,
    Bool,
    Set,
    NoneSet,
    String,
    Integer,)

from openpyxl2.descriptors.excel import(
    Coordinate,
    Percentage,
    HexBinary,
    TextPoint,
    ExtensionList,
)
from .layout import Layout
from .shapes import *
from .text import *
from .error_bar import *


class TrendlineLbl(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Tx, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 layout=None,
                 tx=None,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.tx = tx
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst



class Period(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Order(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class TrendlineType(Serialisable):

    val = Set(values=(['exp', 'linear', 'log', 'movingAvg', 'poly', 'power']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Trendline(Serialisable):

    name = String(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    trendlineType = Typed(expected_type=TrendlineType, )
    order = Typed(expected_type=Order, allow_none=True)
    period = Typed(expected_type=Period, allow_none=True)
    forward = Typed(expected_type=Double, allow_none=True)
    backward = Typed(expected_type=Double, allow_none=True)
    intercept = Typed(expected_type=Double, allow_none=True)
    dispRSqr = Typed(expected_type=Boolean, allow_none=True)
    dispEq = Typed(expected_type=Boolean, allow_none=True)
    trendlineLbl = Typed(expected_type=TrendlineLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 name=None,
                 spPr=None,
                 trendlineType=None,
                 order=None,
                 period=None,
                 forward=None,
                 backward=None,
                 intercept=None,
                 dispRSqr=None,
                 dispEq=None,
                 trendlineLbl=None,
                 extLst=None,
                ):
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl
        self.extLst = extLst


class DLbl(Serialisable):

    idx = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 idx=None,
                 extLst=None,
                ):
        self.idx = idx
        self.extLst = extLst


class DLbls(Serialisable):

    dLbl = Typed(expected_type=DLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 dLbl=None,
                 extLst=None,
                ):
        self.dLbl = dLbl
        self.extLst = extLst


class PictureStackUnit(Serialisable):

    val = Typed(expected_type=Float())

    def __init__(self,
                 val=None,
                ):
        self.val = val


class PictureFormat(Serialisable):

    val = Set(values=(['stretch', 'stack', 'stackScale']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class PictureOptions(Serialisable):

    applyToFront = Typed(expected_type=Boolean, allow_none=True)
    applyToSides = Typed(expected_type=Boolean, allow_none=True)
    applyToEnd = Typed(expected_type=Boolean, allow_none=True)
    pictureFormat = Typed(expected_type=PictureFormat, allow_none=True)
    pictureStackUnit = Typed(expected_type=PictureStackUnit, allow_none=True)

    def __init__(self,
                 applyToFront=None,
                 applyToSides=None,
                 applyToEnd=None,
                 pictureFormat=None,
                 pictureStackUnit=None,
                ):
        self.applyToFront = applyToFront
        self.applyToSides = applyToSides
        self.applyToEnd = applyToEnd
        self.pictureFormat = pictureFormat
        self.pictureStackUnit = pictureStackUnit


class UnsignedInt(Serialisable):

    val = Typed(expected_type=Integer, )

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Marker(Serialisable):

    symbol = Set(values=(['circle', 'dash', 'diamond', 'dot', 'none', 'picture', 'plus', 'square', 'star', 'triangle', 'x', 'auto']), nested=True)
    size = MinMax(min=2, max=72, nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 symbol=None,
                 size=None,
                 spPr=None,
                 extLst=None,
                ):
        self.symbol = symbol
        self.size = size
        self.spPr = spPr
        self.extLst = extLst


class DPt(Serialisable):

    idx = Typed(expected_type=UnsignedInt, )
    invertIfNegative = Typed(expected_type=Boolean, allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    bubble3D = Typed(expected_type=Boolean, allow_none=True)
    explosion = Typed(expected_type=UnsignedInt, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

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
        self.spPr = spPr
        self.pictureOptions = pictureOptions
        self.extLst = extLst


class LineSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DPt, allow_none=True)
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Typed(expected_type=Boolean, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 smooth=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.smooth = smooth
        self.extLst = extLst

