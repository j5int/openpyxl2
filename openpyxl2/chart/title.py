from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Alias,
)

from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.nested import NestedBool

from .text import Text, RichTextProperties
from .layout import Layout
from .shapes import ShapeProperties


class Title(Serialisable):
    tagname = "title"

    tx = Typed(expected_type=Text, allow_none=True)
    text = Alias('tx')
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = NestedBool(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    shapeProperties = Alias('spPr')
    txPr = Typed(expected_type=RichTextProperties, allow_none=True)
    body = Alias('txPr')
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tx', 'layout', 'overlay', 'spPr', 'txPr')

    def __init__(self,
                 tx=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        if tx is None:
            tx = Text()
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr
