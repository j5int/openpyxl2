from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Integer,
    String,
    Sequence,
)

from openpyxl2.cell.cell import Cell, MergedCell
from openpyxl2.styles.borders import Border

from .cell_range import CellRange


class MergeCell(CellRange):

    tagname = "mergeCell"
    ref = CellRange.coord

    __attrs__ = ("ref",)


    def __init__(self,
                 ref=None,
                ):
        super(MergeCell, self).__init__(ref)


    def __copy__(self):
        return self.__class__(self.ref)


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


class MergedCellRange(CellRange):

    """
    MergedCellRange stores the border information of a merged cell in the top
    left cell of the merged cell.
    The remaining cells in the merged cell are stored as MergedCell objects and
    get their border information from the upper left cell.
    """

    def __init__(self, worksheet, coord):
        self.ws = worksheet
        super(MergedCellRange, self).__init__(range_string=coord)
        self.start_cell = None
        self._get_borders()


    def _get_borders(self):
        """
        If the upper left cell of the merged cell does not yet exist, it is
        created.
        The upper left cell gets the border information of the bottom and right
        border from the bottom right cell of the merged cell, if available.
        """

        # Top-left cell.
        if (self.min_row, self.min_col) in self.ws._cells:
            self.start_cell = self.ws._cells[(self.min_row, self.min_col)]
        else:
            self.start_cell = Cell(self.ws, row=self.min_row,
                    column=self.min_col)
            self.ws._cells[(self.start_cell.row,
                self.start_cell.column)] = self.start_cell

        if (self.max_row, self.max_col) in self.ws._cells:
            # Bottom-right cell
            end_cell = self.ws._cells[(self.max_row, self.max_col)]

            self.start_cell.border = self.start_cell.border + Border(
                right=end_cell.border.right, bottom=end_cell.border.bottom)


    def format(self):

        """
        Each cell of the merged cell is created as MergedCell if it does not
        already exist.

        The MergedCells at the edge of the merged cell gets its borders from
        the upper left cell.

         - The top MergedCells get the top border from the top left cell.
         - The bottom MergedCells get the bottom border from the top left cell.
         - The left MergedCells get the left border from the top left cell.
         - The right MergedCells get the right border from the top left cell.
        """


        edge_names = ['top', 'left', 'right', 'bottom']

        for border_name in edge_names:
            edge, border = self._side_for_edge(border_name)
            for coord in edge:
                cell = self.ws._cells.get(coord)
                if cell is None:
                    row, col = coord
                    cell = MergedCell(self.ws, row=row, column=col)
                    self.ws._cells[(cell.row, cell.column)] = cell
                cell.border += border


    def _side_for_edge(self, edge_name):
        """
        Returns the cells of an edge and its border.
        """

        edge = getattr(self, "_" + edge_name)
        side = getattr(self.start_cell.border, edge_name)
        border = Border()
        setattr(border, edge_name, side)
        return edge, border
