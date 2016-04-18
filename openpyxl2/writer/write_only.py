from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


"""Write worksheets to xml representations in an optimized way"""

import atexit
from inspect import isgenerator
import os
from tempfile import NamedTemporaryFile

from openpyxl2.cell import Cell, WriteOnlyCell
from openpyxl2.worksheet import Worksheet
from openpyxl2.worksheet.related import Related
from openpyxl2.worksheet.dimensions import SheetFormatProperties

from openpyxl2.utils.exceptions import WorkbookAlreadySaved

from .etree_worksheet import write_cell
from .excel import ExcelWriter
from .relations import write_rels
from .worksheet import write_drawing
from openpyxl2.xml.constants import SHEET_MAIN_NS
from openpyxl2.xml.functions import xmlfile, Element

ALL_TEMP_FILES = []


@atexit.register
def _openpyxl_shutdown():
    global ALL_TEMP_FILES
    for path in ALL_TEMP_FILES:
        if os.path.exists(path):
            os.remove(path)


def create_temporary_file(suffix=''):
    fobj = NamedTemporaryFile(mode='w+', suffix=suffix,
                              prefix='openpyxl.', delete=False)
    filename = fobj.name
    ALL_TEMP_FILES.append(filename)
    return filename


class WriteOnlyWorksheet(object):
    """
    Streaming worksheet. Optimised to reduce memory by writing rows just in
    time.
    Cells can be styled and have comments Styles for rows and columns
    must be applied before writing cells
    """

    __saved = False
    writer = None
    _default_title = "Sheet"

    def __init__(self, parent, title):
        Worksheet.__init__(self, parent, title)

        self._max_col = 0
        self._max_row = 0

        self._fileobj_name = create_temporary_file()


    @property
    def filename(self):
        return self._fileobj_name


    def _write_header(self):
        """
        Generator that creates the XML file and the sheet header
        """

        with xmlfile(self.filename) as xf:
            with xf.element("worksheet", xmlns=SHEET_MAIN_NS):

                if self.sheet_properties:
                    pr = self.sheet_properties.to_tree()

                xf.write(pr)
                xf.write(self.views.to_tree())

                cols = self.column_dimensions.to_tree()

                self.sheet_format.outlineLevelCol = self.column_dimensions.max_outline
                xf.write(self.sheet_format.to_tree())

                if cols is not None:
                    xf.write(cols)

                with xf.element("sheetData"):
                    cell = WriteOnlyCell(self)
                    try:
                        while True:
                            row = (yield)
                            row_idx = self._max_row
                            with xf.element("row", r="%d" % row_idx):

                                for col_idx, value in enumerate(row, 1):
                                    if value is None:
                                        continue
                                    try:
                                        cell.value = value
                                    except ValueError:
                                        if isinstance(value, Cell):
                                            cell = value
                                        else:
                                            raise ValueError

                                    cell.col_idx = col_idx
                                    cell.row = row_idx

                                    styled = cell.has_style
                                    write_cell(xf, self, cell, styled)

                                    if styled: # styled cell or datetime
                                        cell = WriteOnlyCell(self)

                    except GeneratorExit:
                        pass

                if self.protection.sheet:
                    xf.write(worksheet.protection.to_tree())

                if self.auto_filter.ref:
                    xf.write(self.auto_filter.to_tree())

                if self.sort_state.ref:
                    xf.write(self.sort_state.to_tree())

                if self.data_validations.count:
                    xf.write(self.data_validations.to_tree())

                drawing = write_drawing(self)
                if drawing is not None:
                    xf.write(drawing)

                if self._comments:
                    legacyDrawing = Related(id="anysvml")
                    xml = legacyDrawing.to_tree("legacyDrawing")
                    xf.write(xml)

    def close(self):
        if self.__saved:
            self._already_saved()
        if self.writer is None:
            self.writer = self._write_header()
            next(self.writer)
        self.writer.close()
        self.__saved = True

    def _cleanup(self):
        os.remove(self.filename)

    def append(self, row):
        """
        :param row: iterable containing values to append
        :type row: iterable
        """

        if (not isgenerator(row) and
            not isinstance(row, (list, tuple, range))
            ):
            self._invalid_row(row)

        self._max_row += 1

        if self.writer is None:
            self.writer = self._write_header()
            next(self.writer)

        try:
            self.writer.send(row)
        except StopIteration:
            self._already_saved()


    def _already_saved(self):
        raise WorkbookAlreadySaved('Workbook has already been saved and cannot be modified or saved anymore.')


    def _invalid_row(self, iterable):
        raise TypeError('Value must be a list, tuple, range or a generator Supplied value is {0}'.format(
            type(iterable))
                        )

    def _write(self, shared_strings=None):
        self.close()
        with open(self.filename) as src:
            out = src.read()
        self._cleanup()
        return out

    # from Worksheet

    _add_row = Worksheet._add_row
    _add_column = Worksheet._add_column
    parent = Worksheet.parent
    _rel_type = Worksheet._rel_type
    _path = Worksheet._path
    sheet_view = Worksheet.sheet_view
    freeze_panes = Worksheet.freeze_panes


def save_dump(workbook, filename):
    if workbook.worksheets == []:
        workbook.create_sheet()
    writer = ExcelWriter(workbook)
    writer.save(filename)
    return True
