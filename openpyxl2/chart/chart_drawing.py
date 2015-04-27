from __future__ import absolute_import

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
    Sequence,
)

from openpyxl2.descriptors.excel import Coordinate

from ._chart import RelId
from .shapes import Shape
from .graphic import (
    GroupShape,
    GraphicalObjectFrame,
    Connector,
    Picture,
    PositiveSize2D,
    Point2D,
    )


class AnchorClientData(Serialisable):

    fLocksWithSheet = Bool(allow_none=True)
    fPrintsWithSheet = Bool(allow_none=True)

    def __init__(self,
                 fLocksWithSheet=None,
                 fPrintsWithSheet=None,
                 ):
        self.fLocksWithSheet = fLocksWithSheet
        self.fPrintsWithSheet = fPrintsWithSheet


class AnchorMarker(Serialisable):

    tagname = "marker"

    col = Integer()
    colOff = Coordinate()
    row = Integer()
    rowOff = Coordinate()

    def __init__(self,
                 col=0,
                 colOff=0,
                 row=0,
                 rowOff=0,
                 ):
        self.col = col
        self.colOff = colOff
        self.row = row
        self.rowOff = rowOff


class _AnchorBase(Serialisable):

    #one of
    sp = Typed(expected_type=Shape, allow_none=True)
    grpSp = Typed(expected_type=GroupShape, allow_none=True)
    graphicFrame = Typed(expected_type=GraphicalObjectFrame, allow_none=True)
    cxnSp = Typed(expected_type=Connector, allow_none=True)
    pic = Typed(expected_type=Picture, allow_none=True)
    contentPart = Typed(expected_type=RelId, allow_none=True)

    clientData = Typed(expected_type=AnchorClientData)

    __elements__ = ('sp', 'grpSp', 'graphicFrame',
                    'cxnSp', 'pic', 'contentPart', 'clientData')

    def __init__(self,
                 clientData=None,
                 sp=None,
                 grpSp=None,
                 graphicFrame=None,
                 cxnSp=None,
                 pic=None,
                 contentPart=None
                 ):
        if clientData is None:
            clientData = AnchorClientData()
        self.clientData = clientData
        self.sp = sp
        self.grpSp = grpSp
        self.graphicFrame = graphicFrame
        self.cxnSp = cxnSp
        self.pic = pic
        self.contentPart = contentPart


class AbsoluteAnchor(_AnchorBase):

    tagname = "absoluteAnchor"

    pos = Typed(expected_type=Point2D)
    ext = Typed(expected_type=PositiveSize2D)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('pos', 'ext') + _AnchorBase.__elements__

    def __init__(self,
                 pos=None,
                 ext=None,
                 **kw
                ):
        if pos is None:
            pos = Point2D(0, 0)
        self.pos = pos
        if ext is None:
            ext = PositiveSize2D(0, 0)
        self.ext = ext
        super(AbsoluteAnchor, self).__init__(**kw)


class OneCellAnchor(_AnchorBase):

    tagname = "oneCellAnchor"

    frm = Typed(expected_type=AnchorMarker)
    ext = Typed(expected_type=PositiveSize2D)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('frm', 'ext') + _AnchorBase.__elements__


    def __init__(self,
                 frm=None,
                 ext=None,
                 **kw
                ):
        if frm is None:
            frm = AnchorMarker()
        self.frm = frm
        if ext is None:
            ext = PositiveSize2D(0, 0)
        self.ext = ext
        super(OneCellAnchor, self).__init__(**kw)


class TwoCellAnchor(_AnchorBase):

    tagname = "twoCellAnchor"

    editAs = NoneSet(values=(['twoCell', 'oneCell', 'absolute']))
    frm = Typed(expected_type=AnchorMarker)
    to = Typed(expected_type=AnchorMarker)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('frm', 'to') + _AnchorBase.__elements__

    def __init__(self,
                 editAs=None,
                 frm=None,
                 to=None,
                 **kw
                 ):
        self.editAs = editAs
        if frm is None:
            frm = AnchorMarker()
        self.frm = frm
        if to is None:
            to = AnchorMarker()
        self.to = to
        super(TwoCellAnchor, self).__init__(**kw)


class SpreadsheetDrawing(Serialisable):

    tagname = "wsDr"

    twoCellAnchor = Sequence(expected_type=TwoCellAnchor, allow_none=True)
    oneCellAnchor = Sequence(expected_type=OneCellAnchor, allow_none=True)
    absoluteAnchor = Sequence(expected_type=AbsoluteAnchor, allow_none=True)

    def __init__(self,
                 twoCellAnchor=None,
                 oneCellAnchor=None,
                 absoluteAnchor=None,
                 ):
        self.twoCellAnchor = twoCellAnchor
        self.oneCellAnchor = oneCellAnchor
        self.absoluteAnchor = absoluteAnchor

