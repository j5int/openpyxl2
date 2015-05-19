from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import math

from openpyxl2.compat import deprecated

from openpyxl2.styles.colors import Color, BLACK, WHITE
from openpyxl2.utils.units import (
    pixels_to_EMU,
    EMU_to_pixels,
    short_color,
)
from openpyxl2.chart.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor, AbsoluteAnchor


class Shadow(object):

    SHADOW_BOTTOM = 'b'
    SHADOW_BOTTOM_LEFT = 'bl'
    SHADOW_BOTTOM_RIGHT = 'br'
    SHADOW_CENTER = 'ctr'
    SHADOW_LEFT = 'l'
    SHADOW_TOP = 't'
    SHADOW_TOP_LEFT = 'tl'
    SHADOW_TOP_RIGHT = 'tr'

    def __init__(self):
        self.visible = False
        self.blurRadius = 6
        self.distance = 2
        self.direction = 0
        self.alignment = self.SHADOW_BOTTOM_RIGHT
        self.color = Color()
        self.alpha = 50


class Drawing(object):
    """ a drawing object - eg container for shapes or charts
        we assume user specifies dimensions in pixels; units are
        converted to EMU in the drawing part
    """

    count = 0

    def __init__(self):

        self.name = ''
        self.description = ''
        self.coordinates = ((1, 2), (16, 8))
        self.left = 0
        self.top = 0
        self._width = 21 # default in px
        self._height = 192 #default in px
        self.resize_proportional = False
        self.rotation = 0
        self.anchortype = "absolute"
        self.anchorcol = 0 # left cell
        self.anchorrow = 0 # top row


    @property
    def width(self):
        return self._width

    @width.setter
    def width(self, w):
        if self.resize_proportional and w:
            ratio = self._height / self._width
            self._height = round(ratio * w)
        self._width = w

    @property
    def height(self):
        return self._height

    @height.setter
    def height(self, h):
        if self.resize_proportional and h:
            ratio = self._width / self._height
            self._width = round(ratio * h)
        self._height = h

    def set_dimension(self, w=0, h=0):

        xratio = w / self._width
        yratio = h / self._height

        if self.resize_proportional and w and h:
            if (xratio * self._height) < h:
                self._height = math.ceil(xratio * self._height)
                self._width = w
            else:
                self._width = math.ceil(yratio * self._width)
                self._height = h

    @deprecated("Private method used when serialising")
    def get_emu_dimensions(self):
        """ return (x, y, w, h) in EMU """

        return (pixels_to_EMU(self.left), pixels_to_EMU(self.top),
            pixels_to_EMU(self._width), pixels_to_EMU(self._height))


    @property
    def anchor(self):
        if self.anchortype == "absolute":
            anchor = AbsoluteAnchor()
            anchor.pos.x = pixels_to_EMU(self.left)
            anchor.pos.y = pixels_to_EMU(self.top)

        elif self.anchortype == "oneCell":
            anchor = OneCellAnchor()
            anchor._from.col = self.anchorcol
            anchor._from.row = self.anchorrow

        anchor.ext.width = pixels_to_EMU(self._width)
        anchor.ext.height = pixels_to_EMU(self._height)

        return anchor
