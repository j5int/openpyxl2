from __future__ import absolute_import

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import NestedInteger

from .shapes import ShapeProperties, Shape
from .chartBase import AxDataSource, NumDataSource, NumRef
from .error_bar import ErrorBars
from .label import DataLabels
from .marker import DataPoint, PictureOptions, Marker
from .trendline import Trendline


def Series(name_ref=None, cat_ref=None, values=None, order=None):
    """
    High level function for creating series

    See http://exceluser.com/excel_help/functions/function-series.htm for a description
    """
    series = BarSer(idx=0, order=0, val=NumDataSource(numRef=NumRef(f=values)))
    return series


class StrVal(Serialisable):

    idx = Integer()
    v = Typed(expected_type=String(), )

    def __init__(self,
                 idx=None,
                 v=None,
                ):
        self.idx = idx
        self.v = v


class StrData(Serialisable):

    ptCount = Integer(allow_none=True, nested=True)
    pt = Typed(expected_type=StrVal, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'pt', 'extLst')

    def __init__(self,
                 ptCount=None,
                 pt=None,
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.pt = pt
        self.extLst = extLst


class StrRef(Serialisable):

    f = Typed(expected_type=String, )
    strCache = Typed(expected_type=StrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'strCache', 'extLst')

    def __init__(self,
                 f=None,
                 strCache=None,
                 extLst=None,
                ):
        self.f = f
        self.strCache = strCache
        self.extLst = extLst


class SerTx(Serialisable):

    strRef = Typed(expected_type=StrRef)


class _SeriesBase(Serialisable):

    tagname = "ser"

    idx = NestedInteger()
    order = NestedInteger()
    tx = Typed(expected_type=SerTx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('idx', 'order', 'tx', 'spPr')

    def __init__(self,
                 idx=None,
                 order=None,
                 tx=None,
                 spPr=None,
                ):
        self.idx = idx
        self.order = order
        self.tx = tx
        self.spPr = spPr


class AreaSer(_SeriesBase):

    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _SeriesBase.__elements__ + ('pictureOptions', 'dPt',
                                               'dLbls', 'errBars', 'trendline', 'cat', 'val', 'extLst')

    def __init__(self,
                 pictureOptions=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.pictureOptions = pictureOptions
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.extLst = extLst


class BarSer(_SeriesBase):

    idx = NestedInteger()
    order = NestedInteger()
    tx = Typed(expected_type=SerTx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    invertIfNegative = Bool(nested=True, allow_none=True)
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    shape = Typed(expected_type=Shape, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _SeriesBase.__elements__ + ('invertIfNegative', 'pictureOptions', 'dPt',
                                'dLbls', 'trendline', 'errBars', 'cat', 'val', 'shape')

    def __init__(self,
                 idx=None,
                 order=None,
                 tx=None,
                 spPr=None,
                 invertIfNegative=None,
                 pictureOptions=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 cat=None,
                 val=None,
                 shape=None,
                 extLst=None,
                ):
        self.idx = idx
        self.order = order
        self.tx = tx
        self.spPr = spPr
        self.invertIfNegative = invertIfNegative
        self.pictureOptions = pictureOptions
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.cat = cat
        self.val = val
        self.shape = shape


class BubbleSer(_SeriesBase):

    invertIfNegative = Bool(nested=True, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    bubbleSize = Typed(expected_type=NumDataSource, allow_none=True)
    bubble3D = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _SeriesBase.__elements__ + ('invertIfNegative', 'dPt',
                                               'dLbls', 'trendline', 'errBars', 'xVal', 'yVal', 'bubbleSize',
                                               'bubble3D', 'extLst')

    def __init__(self,
                 invertIfNegative=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 xVal=None,
                 yVal=None,
                 bubbleSize=None,
                 bubble3D=None,
                 extLst=None,
                ):
        self.invertIfNegative = invertIfNegative
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.xVal = xVal
        self.yVal = yVal
        self.bubbleSize = bubbleSize
        self.bubble3D = bubble3D
        self.extLst = extLst


class PieSer(_SeriesBase):

    explosion = Integer(allow_none=True, nested=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _SeriesBase.__elements__ + ('explosion', 'dPt', 'dLbls',
                                               'cat', 'val', 'extLst')

    def __init__(self,
                 explosion=None,
                 dPt=None,
                 dLbls=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.explosion = explosion
        self.dPt = dPt
        self.dLbls = dLbls
        self.cat = cat
        self.val = val
        self.extLst = extLst


class RadarSer(_SeriesBase):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = _SeriesBase.__elements__ + ('marker', 'dPt', 'dLbls', 'cat', 'val', 'extLst')

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.cat = cat
        self.val = val
        self.extLst = extLst


class ScatterSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    xVal = Typed(expected_type=AxDataSource, allow_none=True)
    yVal = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'xVal', 'yVal', 'smooth', 'extLst')

    def __init__(self,
                 marker=None,
                 dPt=None,
                 dLbls=None,
                 trendline=None,
                 errBars=None,
                 xVal=None,
                 yVal=None,
                 smooth=None,
                 extLst=None,
                ):
        self.marker = marker
        self.dPt = dPt
        self.dLbls = dLbls
        self.trendline = trendline
        self.errBars = errBars
        self.xVal = xVal
        self.yVal = yVal
        self.smooth = smooth
        self.extLst = extLst


class SurfaceSer(Serialisable):

    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('cat', 'val', 'extLst')

    def __init__(self,
                 cat=None,
                 val=None,
                 extLst=None,
                ):
        self.cat = cat
        self.val = val
        self.extLst = extLst


class LineSer(Serialisable):

    marker = Typed(expected_type=Marker, allow_none=True)
    dPt = Typed(expected_type=DataPoint, allow_none=True)
    dLbls = Typed(expected_type=DataLabels, allow_none=True)
    trendline = Typed(expected_type=Trendline, allow_none=True)
    errBars = Typed(expected_type=ErrorBars, allow_none=True)
    cat = Typed(expected_type=AxDataSource, allow_none=True)
    val = Typed(expected_type=NumDataSource, allow_none=True)
    smooth = Bool(allow_none=True, nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('marker', 'dPt', 'dLbls', 'trendline', 'errBars', 'cat', 'val', 'smooth', 'extLst')

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
