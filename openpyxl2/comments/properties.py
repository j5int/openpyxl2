from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Float,
    Integer,
    Set,
    String,
    Bool,
)
from openpyxl2.descriptors.excel import Guid, ExtensionList
from openpyxl2.descriptors.sequence import NestedSequence

from .text import Text
from .author import AuthorList


class ObjectAnchor(Serialisable):

    moveWithCells = Bool(allow_none=True)
    sizeWithCells = Bool(allow_none=True)
    #z-order = Integer(allow_none=True) needs alias
    #from
    #to defs from xdr

    def __init__(self,
                 moveWithCells=None,
                 sizeWithCells=None,
                 #z-order=None,
                ):
        self.moveWithCells = moveWithCells
        self.sizeWithCells = sizeWithCells
        #self.z-order = z-order


class Properties(Serialisable):

    locked = Bool(allow_none=True)
    defaultSize = Bool(allow_none=True)
    _print = Bool(allow_none=True)
    disabled = Bool(allow_none=True)
    uiObject = Bool(allow_none=True)
    autoFill = Bool(allow_none=True)
    autoLine = Bool(allow_none=True)
    altText = String(allow_none=True)
    textHAlign = Set(values=(['left', 'center', 'right', 'justify', 'distributed']))
    textVAlign = Set(values=(['top', 'center', 'bottom', 'justify', 'distributed']))
    lockText = Bool(allow_none=True)
    justLastX = Bool(allow_none=True)
    autoScale = Bool(allow_none=True)
    rowHidden = Bool(allow_none=True)
    colHidden = Bool(allow_none=True)
    anchor = Typed(expected_type=ObjectAnchor, )

    __elements__ = ('anchor',)

    def __init__(self,
                 locked=None,
                 defaultSize=None,
                 _print=None,
                 disabled=None,
                 uiObject=None,
                 autoFill=None,
                 autoLine=None,
                 altText=None,
                 textHAlign=None,
                 textVAlign=None,
                 lockText=None,
                 justLastX=None,
                 autoScale=None,
                 rowHidden=None,
                 colHidden=None,
                 anchor=None,
                ):
        self.locked = locked
        self.defaultSize = defaultSize
        self._print = _print
        self.disabled = disabled
        self.uiObject = uiObject
        self.autoFill = autoFill
        self.autoLine = autoLine
        self.altText = altText
        self.textHAlign = textHAlign
        self.textVAlign = textVAlign
        self.lockText = lockText
        self.justLastX = justLastX
        self.autoScale = autoScale
        self.rowHidden = rowHidden
        self.colHidden = colHidden
        self.anchor = anchor



class Comment(Serialisable):

    tagname = "comment"

    ref = String()
    authorId = Integer(0)
    guid = Guid(allow_none=True)
    shapeId = Integer(allow_none=True)
    text = Typed(expected_type=Text)
    commentPr = Typed(expected_type=Properties, allow_none=True)

    __elements__ = ('text', 'commentPr')

    def __init__(self,
                 ref="",
                 authorId=0,
                 guid=None,
                 shapeId=None,
                 text=None,
                 commentPr=None,
                ):
        self.ref = ref
        self.authorId = authorId
        self.guid = guid
        self.shapeId = shapeId
        if text is None:
            text = Text()
        self.text = text
        self.commentPr = commentPr


class Comments(Serialisable):

    tagname = "comments"

    authors = Typed(expected_type=AuthorList)
    commentList = NestedSequence(expected_type=Comment)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('authors', 'commentList')

    def __init__(self,
                 authors=None,
                 commentList=None,
                 extLst=None,
                ):
        self.authors = authors
        self.commentList = commentList
