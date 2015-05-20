from __future__ import absolute_import

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    String,
)
from openpyxl2.descriptors.excel import Relation


class Hyperlink(Serialisable):

    tagname = "hyperlink"

    ref = String()
    location = String(allow_none=True)
    tooltip = String(allow_none=True)
    display = String(allow_none=True)
    id = Relation()

    def __init__(self,
                 ref=None,
                 location=None,
                 tooltip=None,
                 display=None,
                 id=None,
                ):
        self.ref = ref
        self.location = location
        self.tooltip = tooltip
        self.display = display
        self.id = id
