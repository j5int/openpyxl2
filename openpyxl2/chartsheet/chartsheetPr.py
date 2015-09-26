from __future__ import absolute_import

from openpyxl2.descriptors import (Bool, String, Typed)
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.styles import Color


class ChartsheetPr(Serialisable):
    tagname = "sheetPr"

    published = Bool(allow_none=True)
    codeName = String(allow_none=True)
    tabColor = Typed(expected_type=Color, allow_none=True)

    __elements__ = ('tabColor',)

    def __init__(self,
                 published=None,
                 codeName=None,
                 tabColor=None,
                 ):
        self.published = published
        self.codeName = codeName
        self.tabColor = tabColor
