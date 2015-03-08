from __future__ import absolute_import

from openpyxl2.descriptors import Integer, String, Typed
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.styles import (
    Font,
    Fill,
    GradientFill,
    PatternFill,
    Border,
    Alignment,
    Protection,
    )

from openpyxl2.xml.functions import localname, Element


class NumFmt(Serialisable):

    numFmtId = Integer()
    formatCode = String()

    def __init__(self,
                 numFmtId=None,
                 formatCode=None,
                ):
        self.numFmtId = numFmtId
        self.formatCode = formatCode


class DifferentialStyle(Serialisable):

    tagname = "dxf"

    __elements__ = ("font", "numFmt", "fill", "alignment", "border", "protection")

    font = Typed(expected_type=Font, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    fill = Typed(expected_type=Fill, allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    border = Typed(expected_type=Border, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)

    def __init__(self,
                 font=None,
                 numFmt=None,
                 fill=None,
                 alignment=None,
                 border=None,
                 protection=None,
                 extLst=None,
                ):
        self.font = font
        self.numFmt = numFmt
        self.fill = fill
        self.alignment = alignment
        self.border = border
        self.protection = protection
        self.extLst = extLst


    def __eq__(self, other):
        if not self.__class__ == other.__class__:
            return False
        for key in self.__elements__:
            if getattr(self, key) != getattr(other, key):
                return False
        return True


    def __ne__(self, other):
        return not self == other
