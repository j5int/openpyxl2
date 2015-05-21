from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Bool,
    NoneSet,
    Integer,
    Sequence,
)
from openpyxl2.descriptors.nested import (
    NestedText,
    NestedNoneSet,
)
from openpyxl2.descriptors.excel import Relation

from openpyxl2.packaging.relationship import Relationship
from openpyxl2.utils import coordinate_to_tuple
from openpyxl2.utils.units import cm_to_EMU
from openpyxl2.drawing.image import Image

from openpyxl2.xml.constants import SHEET_DRAWING_NS, PKG_REL_NS
from openpyxl2.xml.functions import Element

from openpyxl2.chart._chart import ChartBase
from .shapes import (
    Point2D,
    PositiveSize2D,
)
from .fill import Blip
from .graphic import (
    GroupShape,
    GraphicFrame,
    Connector,
    PictureFrame,
    ChartRelation,
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

    col = NestedText(expected_type=int)
    colOff = NestedText(expected_type=int)
    row = NestedText(expected_type=int)
    rowOff = NestedText(expected_type=int)

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
    sp = NestedNoneSet(values=(['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax']))
    grpSp = Typed(expected_type=GroupShape, allow_none=True)
    graphicFrame = Typed(expected_type=GraphicFrame, allow_none=True)
    cxnSp = Typed(expected_type=Connector, allow_none=True)
    pic = Typed(expected_type=PictureFrame, allow_none=True)
    contentPart = Relation()

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

    _from = Typed(expected_type=AnchorMarker)
    ext = Typed(expected_type=PositiveSize2D)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('_from', 'ext') + _AnchorBase.__elements__


    def __init__(self,
                 _from=None,
                 ext=None,
                 **kw
                ):
        if _from is None:
            _from = AnchorMarker()
        self._from = _from
        if ext is None:
            ext = PositiveSize2D(0, 0)
        self.ext = ext
        super(OneCellAnchor, self).__init__(**kw)


class TwoCellAnchor(_AnchorBase):

    tagname = "twoCellAnchor"

    editAs = NoneSet(values=(['twoCell', 'oneCell', 'absolute']))
    _from = Typed(expected_type=AnchorMarker)
    to = Typed(expected_type=AnchorMarker)

    sp = _AnchorBase.sp
    grpSp = _AnchorBase.grpSp
    graphicFrame = _AnchorBase.graphicFrame
    cxnSp = _AnchorBase.cxnSp
    pic = _AnchorBase.pic
    contentPart = _AnchorBase.contentPart
    clientData = _AnchorBase.clientData

    __elements__ = ('_from', 'to') + _AnchorBase.__elements__

    def __init__(self,
                 editAs=None,
                 _from=None,
                 to=None,
                 **kw
                 ):
        self.editAs = editAs
        if _from is None:
            _from = AnchorMarker()
        self._from = _from
        if to is None:
            to = AnchorMarker()
        self.to = to
        super(TwoCellAnchor, self).__init__(**kw)


class SpreadsheetDrawing(Serialisable):

    tagname = "wsDr"

    twoCellAnchor = Sequence(expected_type=TwoCellAnchor, allow_none=True)
    oneCellAnchor = Sequence(expected_type=OneCellAnchor, allow_none=True)
    absoluteAnchor = Sequence(expected_type=AbsoluteAnchor, allow_none=True)

    __elements__ = ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor")

    def __init__(self,
                 twoCellAnchor=(),
                 oneCellAnchor=(),
                 absoluteAnchor=(),
                 ):
        self.twoCellAnchor = twoCellAnchor
        self.oneCellAnchor = oneCellAnchor
        self.absoluteAnchor = absoluteAnchor
        self.charts = []
        self.images = []
        self._rels = []


    def __hash__(self):
        """
        Just need to check for identity
        """
        return id(self)


    def _write(self):
        """
        create required structure and the serialise
        """
        anchors = []
        for idx, obj in enumerate(self.charts + self.images, 1):
            if isinstance(obj, ChartBase):
                rel = Relationship(type="chart", target='../charts/chart%s.xml' % obj._id)
                row, col = coordinate_to_tuple(obj.anchor)
                anchor = OneCellAnchor()
                anchor._from.row = row -1
                anchor._from.col = col -1
                anchor.ext.width = cm_to_EMU(obj.width)
                anchor.ext.height = cm_to_EMU(obj.height)
                anchor.graphicFrame = self._chart_frame(idx)
            elif isinstance(obj, Image):
                rel = Relationship(type="image", target='../media/image%s.png' % obj._id)
                anchor = obj.drawing.anchor
                anchor.pic = self._picture_frame(idx)

            anchors.append(anchor)
            self._rels.append(rel)

        self.oneCellAnchor = anchors

        tree = self.to_tree()
        tree.set('xmlns', SHEET_DRAWING_NS)
        return tree


    def _chart_frame(self, idx):
        chart_rel = ChartRelation("rId%s" % idx)
        frame = GraphicFrame()
        nv = frame.nvGraphicFramePr.cNvPr
        nv.id = idx
        nv.name = "Chart {0}".format(idx)
        frame.graphic.graphicData.chart = chart_rel
        return frame


    def _picture_frame(self, idx):
        pic = PictureFrame()
        pic.nvPicPr.cNvPr.desc = "Picture 1"
        pic.nvPicPr.cNvPr.id = idx
        pic.nvPicPr.cNvPicPr.picLocks
        pic.blipFill.blip = Blip()
        pic.blipFill.blip.embed = "rId{0}".format(idx)
        pic.blipFill.blip.cstate = "print"

        pic.spPr.noFill = True
        pic.spPr.ln.prstDash = None
        pic.spPr.ln.w = 1
        pic.spPr.ln.noFill = True
        return pic


    def _write_rels(self):
        root = Element("Relationships", xmlns=PKG_REL_NS)
        for idx, rel in enumerate(self._rels, 1):
            rel.id = "rId{0}".format(idx)
            root.append(rel.to_tree())
        return root
