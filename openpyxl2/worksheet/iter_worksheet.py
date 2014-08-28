from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

""" Iterators-based worksheet reader
*Still very raw*
"""
# stdlib
import operator
from itertools import groupby

# compatibility
from openpyxl2.compat import range
from openpyxl2.xml.functions import iterparse

# package
from openpyxl2.worksheet import Worksheet
from openpyxl2.cell import (
    ABSOLUTE_RE,
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
    Cell
)
from openpyxl2.cell.read_only import ReadOnlyCell, EMPTY_CELL
from openpyxl2.xml.functions import safe_iterator
from openpyxl2.xml.constants import SHEET_MAIN_NS


def read_dimension(source):
    min_row = min_col =  max_row = max_col = None
    DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS
    DATA_TAG = '{%s}sheetData' % SHEET_MAIN_NS
    it = iterparse(source, tag=[DIMENSION_TAG, DATA_TAG])
    for _event, element in it:
        if element.tag == DIMENSION_TAG:
            dim = element.get("ref")
            m = ABSOLUTE_RE.match(dim.upper())
            min_col, min_row, sep, max_col, max_row = m.groups()
            min_row = int(min_row)
            if max_col is None or max_row is None:
                max_col = min_col
                max_row = min_row
            else:
                max_row = int(max_row)
            return min_col, min_row, max_col, max_row

        elif element.tag == DATA_TAG:
            # Dimensions missing
            break
        element.clear()


ROW_TAG = '{%s}row' % SHEET_MAIN_NS
CELL_TAG = '{%s}c' % SHEET_MAIN_NS
VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
DIMENSION_TAG = '{%s}dimension' % SHEET_MAIN_NS


class IterableWorksheet(Worksheet):

    min_col = 'A'
    min_row = 1
    max_col = max_row = None

    def __init__(self, parent_workbook, title, worksheet_path,
                 xml_source, shared_strings, style_table):
        Worksheet.__init__(self, parent_workbook, title)
        self.worksheet_path = worksheet_path
        self.shared_strings = shared_strings
        self.base_date = parent_workbook.excel_base_date
        dimensions = read_dimension(self.xml_source)
        if dimensions is not None:
            self.min_col, self.min_row, self.max_col, self.max_row = dimensions

    @property
    def xml_source(self):
        return self.parent._archive.open(self.worksheet_path)

    @xml_source.setter
    def xml_source(self, value):
        """Base class is always supplied XML source, IteratableWorksheet obtains it on demand."""
        pass

    @property
    def dimensions(self):
        if not all([self.max_col, self.max_row]):
            raise ValueError("Worksheet is unsized, cannot calculate dimensions")
        return '%s%s:%s%s' % (self.min_col, self.min_row, self.max_col, self.max_row)

    def __getitem__(self, key):
        if isinstance(key, slice):
            key = "{0}:{1}".format(key.start, key.stop)
        if ":" in key:
            return self.iter_rows(key)
        return self.cell(key)

    def iter_rows(self, range_string=None, row_offset=0, column_offset=1):
        """ Returns a squared range based on the `range_string` parameter,
        using generators.

        :param range_string: range of cells (e.g. 'A1:C4')
        :type range_string: string

        :param row_offset: additional rows (e.g. 4)
        :type row: int

        :param column_offset: additonal columns (e.g. 3)
        :type column: int

        :rtype: generator

        """
        if range_string is not None:
            min_col, min_row, max_col, max_row = self._range_boundaries(range_string)
            max_col += column_offset
            max_row += row_offset
        else:
            min_col = column_index_from_string(self.min_col)
            max_col = self.max_col
            if max_col is not None:
                max_col = column_index_from_string(self.max_col) + 1
            min_row = self.min_row
            max_row = self.max_row

        return self.get_squared_range(min_col, min_row, max_col, max_row)

    def get_squared_range(self, min_col, min_row, max_col, max_row):
        """
        The source worksheet file may have columns or rows missing.
        Missing cells will be created.
        """
        if max_col is not None:
            expected_columns = [get_column_letter(ci) for ci in range(min_col, max_col)]
        else:
            expected_columns = []
        row_counter = min_row

        # get cells row by row
        for row, cells in groupby(self.get_cells(min_row, min_col,
                                                 max_row, max_col),
                                  operator.attrgetter('row')):
            full_row = []
            if row_counter < row:
                # Rows requested before those in the worksheet
                for gap_row in range(row_counter, row):
                    yield tuple(EMPTY_CELL for column in expected_columns)
                    row_counter = row

            if expected_columns:
                retrieved_columns = dict([(c.column, c) for c in cells])
                for column in expected_columns:
                    if column in retrieved_columns:
                        cell = retrieved_columns[column]
                        full_row.append(cell)
                    else:
                        # create missing cell
                        full_row.append(EMPTY_CELL)
            else:
                full_row = tuple(cells)
            row_counter = row + 1
            yield tuple(full_row)

    def get_cells(self, min_row, min_col, max_row, max_col):
        p = iterparse(self.xml_source, tag=[ROW_TAG], remove_blank_text=True)
        for _event, element in p:
            if element.tag == ROW_TAG:
                row = int(element.get("r"))
                if max_row is not None and row > max_row:
                    break
                if min_row <= row:
                    for cell in safe_iterator(element, CELL_TAG):
                        coord = cell.get('r')
                        column_str, row = coordinate_from_string(coord)
                        column = column_index_from_string(column_str)
                        if max_col is not None and column > max_col:
                            break
                        if min_col <= column:
                            data_type = cell.get('t', 'n')
                            style_id = cell.get('s')
                            formula = cell.findtext(FORMULA_TAG)
                            value = cell.findtext(VALUE_TAG)
                            if formula is not None and not self.parent.data_only:
                                data_type = Cell.TYPE_FORMULA
                                value = "=%s" % formula
                            yield ReadOnlyCell(self, row, column_str,
                                               value, data_type, style_id)
            if element.tag in (CELL_TAG, VALUE_TAG, FORMULA_TAG):
                # sub-elements of rows should be skipped
                continue
            element.clear()

    def _get_cell(self, coordinate):
        """.iter_rows always returns a generator of rows each of which
        contains a generator of cells. This can be empty in which case
        return None"""
        result = list(self.iter_rows(coordinate))
        if result:
            return result[0][0]

    def range(self, *args, **kwargs):
        # TODO return a range of cells, basically get_squared_range with same interface as Worksheet
        raise NotImplementedError("use 'iter_rows()' instead")

    @property
    def rows(self):
        return self.iter_rows()

    def calculate_dimension(self):
        return self.dimensions

    def get_highest_column(self):
        if self.max_col is not None:
            return column_index_from_string(self.max_col)

    def get_highest_row(self):
        return self.max_row

    def get_style(self, coordinate):
        raise NotImplementedError("use `cell.style` instead")
