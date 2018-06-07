from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Integer,
    String,
    Sequence,
)

from .cell_range import CellRange


class MergeCell(CellRange):

    tagname = "mergeCell"
    ref = CellRange.coord

    __attrs__ = ("ref",)


    def __init__(self,
                 ref=None,
                ):
        super(MergeCell, self).__init__(ref)


class MergeCells(Serialisable):

    tagname = "mergeCells"

    count = Integer(allow_none=True)
    mergeCell = Sequence(expected_type=MergeCell, )

    __elements__ = ('mergeCell',)
    __attrs__ = ('count',)

    def __init__(self,
                 count=None,
                 mergeCell=(),
                ):
        self.mergeCell = mergeCell


    @property
    def count(self):
        return len(self.mergeCell)
