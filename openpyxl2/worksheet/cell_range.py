from __future__ import absolute_import, unicode_literals
# Copyright (c) 2010-2018 openpyxl

from copy import copy
from itertools import islice, product

from openpyxl2.compat.strings import safe_repr
from openpyxl2.descriptors import Strict
from openpyxl2.descriptors import MinMax, Sequence
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.styles.borders import Border
from openpyxl2.cell.cell import Cell, MergedCell

from openpyxl2.utils import (
    range_boundaries,
    range_to_tuple,
    get_column_letter,
    quote_sheetname,
)


class CellRange(Serialisable):
    """
    Represents a range in a sheet: title and coordinates.

    This object is used to perform operations on ranges, like:

    - shift, expand or shrink
    - union/intersection with another sheet range,

    We can check whether a range is:

    - equal or not equal to another,
    - disjoint of another,
    - contained in another.

    We can get:

    - the size of a range.
    - the range bounds (vertices)
    - the coordinates,
    - the string representation,

    """

    min_col = MinMax(min=1, max=18278, expected_type=int)
    min_row = MinMax(min=1, max=1048576, expected_type=int)
    max_col = MinMax(min=1, max=18278, expected_type=int)
    max_row = MinMax(min=1, max=1048576, expected_type=int)


    def __init__(self, range_string=None, min_col=None, min_row=None,
                 max_col=None, max_row=None, title=None):
        if range_string is not None:
            try:
                title, (min_col, min_row, max_col, max_row) = range_to_tuple(range_string)
            except ValueError:
                min_col, min_row, max_col, max_row = range_boundaries(range_string)

        self.min_col = min_col
        self.min_row = min_row
        self.max_col = max_col
        self.max_row = max_row
        self.title = title

        if min_col > max_col:
            fmt = "{max_col} must be greater than {min_col}"
            raise ValueError(fmt.format(min_col=min_col, max_col=max_col))
        if min_row > max_row:
            fmt = "{max_row} must be greater than {min_row}"
            raise ValueError(fmt.format(min_row=min_row, max_row=max_row))


    @property
    def bounds(self):
        """
        Vertices of the range as a tuple
        """
        return self.min_col, self.min_row, self.max_col, self.max_row


    @property
    def coord(self):
        """
        Excel style representation of the range
        """
        fmt = "{min_col}{min_row}:{max_col}{max_row}"
        if (self.min_col == self.max_col
            and self.min_row == self.max_row):
            fmt = "{min_col}{min_row}"

        return fmt.format(
            min_col=get_column_letter(self.min_col),
            min_row=self.min_row,
            max_col=get_column_letter(self.max_col),
            max_row=self.max_row
        )


    def _check_title(self, other):
        """
        Check whether comparisons between ranges are possible.
        Cannot compare ranges from different worksheets
        Skip if the range passed in has no title.
        """
        if not isinstance(other, CellRange):
            raise TypeError(repr(type(other)))

        if other.title and self.title != other.title:
            raise ValueError("Cannot work with ranges from different worksheets")


    def __repr__(self):
        fmt = u"<{cls} {coord}>"
        if self.title:
            fmt = u"<{cls} {title!r}!{coord}>"
        return safe_repr(fmt.format(cls=self.__class__.__name__, title=self.title, coord=self.coord))


    def _get_range_string(self):
        fmt = "{coord}"
        title = self.title
        if self.title:
            fmt = u"{title}!{coord}"
            title = quote_sheetname(self.title)
        return fmt.format(title=title, coord=self.coord)

    __unicode__ = _get_range_string


    def __str__(self):
        coord = self._get_range_string()
        return safe_repr(coord)


    def __copy__(self):
        return self.__class__(min_col=self.min_col, min_row=self.min_row,
                              max_col=self.max_col, max_row=self.max_row,
                              title=self.title)


    def shift(self, col_shift=0, row_shift=0):
        """
        Shift the range according to the shift values (*col_shift*, *row_shift*).

        :type col_shift: int
        :param col_shift: number of columns to be moved by, can be negative
        :type row_shift: int
        :param row_shift: number of rows to be moved by, can be negative
        :raise: :class:`ValueError` if any row or column index < 1
        """

        if (self.min_col + col_shift <= 0
            or self.min_row + row_shift <= 0):
            raise ValueError("Invalid shift value: col_shift={0}, row_shift={1}".format(col_shift, row_shift))
        self.min_col += col_shift
        self.min_row += row_shift
        self.max_col += col_shift
        self.max_row += row_shift


    def __ne__(self, other):
        """
        Test whether the ranges are not equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* != *other*.
        """
        try:
            self._check_title(other)
        except ValueError:
            return True

        return (
            other.min_row != self.min_row
            or self.max_row != other.max_row
            or other.min_col != self.min_col
            or self.max_col != other.max_col
                )


    def __eq__(self, other):
        """
        Test whether the ranges are equal.

        :type other: CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* == *other*.
        """
        return not self.__ne__(other)


    def issubset(self, other):
        """
        Test whether every element in the range is in *other*.

        :type other: SheetRange
        :param other: Other sheet range
        :return: ``True`` if *range* <= *other*.
        """
        self._check_title(other)

        return (
            (other.min_row <= self.min_row <= self.max_row <= other.max_row)
            and
            (other.min_col <= self.min_col <= self.max_col <= other.max_col)
        )

    __le__ = issubset


    def __lt__(self, other):
        """
        Test whether every element in the range is in *other*, but not all.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* < *other*.
        """
        return self.__le__(other) and self.__ne__(other)


    def issuperset(self, other):
        """
        Test whether every element in *other* is in the range.

        :type other: openpyxl.worksheet.cell_range.CellRange or tuple[int, int]
        :param other: Other sheet range or cell index (*row_idx*, *col_idx*).
        :return: ``True`` if *range* >= *other* (or *other* in *range*).
        """
        self._check_title(other)

        return (
            (self.min_row <= other.min_row <= other.max_row <= self.max_row)
            and
            (self.min_col <= other.min_col <= other.max_col <= self.max_col)
        )


    __ge__ = issuperset


    def __contains__(self, coord):
        """
        Check whether the range contains a particular cell coordinate
        """
        cr = self.__class__(coord)
        if cr.title is None:
            cr.title = self.title
        return self.issuperset(cr)


    def __gt__(self, other):
        """
        Test whether every element in *other* is in the range, but not all.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range
        :return: ``True`` if *range* > *other*.
        """
        return self.__ge__(other) and self.__ne__(other)


    def isdisjoint(self, other):
        """
        Return ``True`` if the range has no elements in common with other.
        Ranges are disjoint if and only if their intersection is the empty range.

        :type other: openpyxl.worksheet.cell_range.CellRange
        :param other: Other sheet range.
        :return: `True`` if the range has no elements in common with other.
        """
        self._check_title(other)

        # sort by top-left vertex
        if self.bounds > other.bounds:
            i = self
            self = other
            other = i

        return (self.max_col, self.max_row) < (other.min_col, other.max_row)


    def intersection(self, other):
        """
        Return a new range with elements common to the range and another

        :type others: tuple[openpyxl.worksheet.cell_range.CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        :raise: :class:`ValueError` if an *other* range don't intersect
            with the current range.
        """
        if self.isdisjoint(other):
            raise ValueError("Range {0} don't intersect {0}".format(self, other))

        min_row = max(self.min_row, other.min_row)
        max_row = min(self.max_row, other.max_row)
        min_col = max(self.min_col, other.min_col)
        max_col = min(self.max_col, other.max_col)

        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row)

    __and__ = intersection


    def union(self, other):
        """
        Return a new range with elements from the range and all *others*.

        :type others: tuple[openpyxl.worksheet.cell_range.CellRange]
        :param others: Other sheet ranges.
        :return: the current sheet range.
        """
        self._check_title(other)

        min_row = min(self.min_row, other.min_row)
        max_row = max(self.max_row, other.max_row)
        min_col = min(self.min_col, other.min_col)
        max_col = max(self.max_col, other.max_col)
        return CellRange(min_col=min_col, min_row=min_row, max_col=max_col,
                         max_row=max_row, title=self.title)


    __or__ = union


    def expand(self, right=0, down=0, left=0, up=0):
        """
        Expand the range by the dimensions provided.

        :type right: int
        :param right: expand range to the right by this number of cells
        :type down: int
        :param down: expand range down by this number of cells
        :type left: int
        :param left: expand range to the left by this number of cells
        :type up: int
        :param up: expand range up by this number of cells
        """
        self.min_col -= left
        self.min_row -= up
        self.max_col += right
        self.max_row += down


    def shrink(self, right=0, bottom=0, left=0, top=0):
        """
        Shrink the range by the dimensions provided.

        :type right: int
        :param right: shrink range from the right by this number of cells
        :type down: int
        :param down: shrink range from the top by this number of cells
        :type left: int
        :param left: shrink range from the left by this number of cells
        :type up: int
        :param up: shrink range from the bottown by this number of cells
        """
        self.min_col += left
        self.min_row += top
        self.max_col -= right
        self.max_row -= bottom


    @property
    def size(self):
        """ Return the size of the range as a dicitionary of rows and columns. """
        cols = self.max_col + 1 - self.min_col
        rows = self.max_row + 1 - self.min_row
        return {'columns':cols, 'rows':rows}

    @property
    def top(self):
        top = []
        for col in range(self.min_col, self.max_col+1):
            top.append((self.min_row, col))
        return top

    @property
    def bottom(self):
        bottom = []
        for col in range(self.min_col, self.max_col+1):
            bottom.append((self.max_row, col))
        return bottom

    @property
    def left(self):
        left = []
        for row in range(self.min_row, self.max_row+1):
            left.append((row, self.min_col))
        return left

    @property
    def right(self):
        right = []
        for row in range(self.min_row, self.max_row+1):
            right.append((row, self.max_col))
        return right

class MultiCellRange(Strict):


    ranges = Sequence(expected_type=CellRange)


    def __init__(self, ranges=()):
        if isinstance(ranges, str):
            ranges = [CellRange(r) for r in ranges.split()]
        self.ranges = ranges


    def __contains__(self, coord):
        for r in self.ranges:
            if coord in r:
                return True
        return False


    def __repr__(self):
        ranges = " ".join([str(r) for r in self.ranges])
        return "<{0} [{1}]>".format(self.__class__.__name__, ranges)


    def __str__(self):
        ranges = u" ".join([str(r) for r in self.ranges])
        return ranges

    __unicode__ = __str__


    def add(self, coord):
        """
        Add a cell coordinate or CellRange
        """
        cr = None
        if isinstance(coord, CellRange):
            cr = coord
            coord = cr.coord
        if coord not in self:
            if cr is None:
                cr = CellRange(coord)
            ranges = self.ranges
            ranges.append(cr)
            self.ranges = ranges


    def __iadd__(self, coord):
        self.add(coord)
        return self


    def __eq__(self, other):
        if  isinstance(other, str):
            other = self.__class__(other)
        return self.ranges == other.ranges


    def __ne__(self, other):
        return not self == other


    def __bool__(self):
        return bool(self.ranges)

    __nonzero__ = __bool__


    def remove(self, coord):
        if not isinstance(coord, CellRange):
            coord = CellRange(coord)
        self.ranges.remove(coord)


    def __iter__(self):
        for cr in self.ranges:
            yield cr


    def __copy__(self):
        n = MultiCellRange()

        for r in self.ranges:
            n.ranges.append(copy(r))
        return n


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
        self.get_borders()

    def __repr__(self):
        fmt = u"<{cls} {coord}>"
        if self.title:
            fmt = u"<{cls} {title!r}!{coord}>"
        return safe_repr(fmt.format(cls=self.__class__.__name__,
                                    title=self.title, coord=self.coord))


    def get_borders(self):
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
                    col_idx=self.min_col)
            self.ws._cells[(self.start_cell.row,
                self.start_cell.col_idx)] = self.start_cell

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

        border_options= (
                (self.top, Border(top=self.start_cell.border.top)),
                (self.bottom, Border(bottom=self.start_cell.border.bottom)),
                (self.left, Border(left=self.start_cell.border.left)),
                (self.right, Border(right=self.start_cell.border.right)),
                )

        for edge in border_options:
            for coord in edge[0]:
                cell = self.ws._cells.get(coord)
                if cell == None:
                    row, col = coord
                    cell = MergedCell(self.ws, row=row, col_idx=col)
                    self.ws._cells[(cell.row, cell.col_idx)] = cell
                cell.border += edge[1]
