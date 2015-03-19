from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Integer,
    Bool,
    Sequence,
)


class Break(Serialisable):

    id = Integer(allow_none=True)
    min = Integer(allow_none=True)
    max = Integer(allow_none=True)
    man = Bool(allow_none=True)
    pt = Bool(allow_none=True)

    def __init__(self,
                 id=None,
                 min=None,
                 max=None,
                 man=None,
                 pt=None,
                ):
        self.id = id
        self.min = min
        self.max = max
        self.man = man
        self.pt = pt


class PageBreak(Serialisable):

    count = Integer(allow_none=True)
    manualBreakCount = Integer(allow_none=True)
    brk = Sequence(expected_type=Break, allow_none=True)

    __elements__ = ('brk',)
    __attrs__ = ("count", "manualBreakCount",)

    def __init__(self,
                 count=None,
                 manualBreakCount=None,
                 brk=None,
                ):
        self.manualBreakCount = manualBreakCount
        self.brk = brk

    @property
    def count(self):
        return len(self.brk)

    @property
    def manualBreakCount(self):
        return len(self.brk)
