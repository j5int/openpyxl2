from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Read an xlsx file into Python"""

# Python stdlib imports
from zipfile import ZipFile, ZIP_DEFLATED, BadZipfile
from sys import exc_info
from io import BytesIO
import os.path
import warnings

# compatibility imports
from openpyxl2.compat import unicode, file

# Allow blanket setting of KEEP_VBA for testing
try:
    from ..tests import KEEP_VBA
except ImportError:
    KEEP_VBA = False


# package imports
from openpyxl2.utils.exceptions import InvalidFileException
from openpyxl2.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CORE,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_WORKBOOK_RELS,
    ARC_THEME,
    COMMENTS_NS,
    SHARED_STRINGS,
    EXTERNAL_LINK,
    XLTM,
    XLTX,
)

from openpyxl2.comments.properties import CommentSheet
from openpyxl2.workbook import Workbook
from openpyxl2.workbook.names.external import detect_external_links
from openpyxl2.workbook.names.named_range import read_named_ranges

from .strings import read_string_table
from openpyxl2.styles.stylesheet import apply_stylesheet

from openpyxl2.packaging.core import DocumentProperties
from openpyxl2.packaging.manifest import Manifest
from openpyxl2.packaging.workbook import WorkbookParser
from openpyxl2.packaging.relationship import get_dependents, get_rels_path

from openpyxl2.worksheet.read_only import ReadOnlyWorksheet
from openpyxl2.xml.functions import fromstring

from .worksheet import WorkSheetParser

# Use exc_info for Python 2 compatibility with "except Exception[,/ as] e"


CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'
SUPPORTED_FORMATS = ('.xlsx', '.xlsm', '.xltx', '.xltm')


def repair_central_directory(zipFile, is_file_instance):
    ''' trims trailing data from the central directory
    code taken from http://stackoverflow.com/a/7457686/570216, courtesy of Uri Cohen
    '''

    f = zipFile if is_file_instance else open(zipFile, 'rb+')
    data = f.read()
    pos = data.find(CENTRAL_DIRECTORY_SIGNATURE)  # End of central directory signature
    if (pos > 0):
        sio = BytesIO(data)
        sio.seek(pos + 22)  # size of 'ZIP end of central directory record'
        sio.truncate()
        sio.seek(0)
        return sio

    f.seek(0)
    return f



def _validate_archive(filename):
    """
    Check the file is a valid zipfile
    """
    is_file_like = hasattr(filename, 'read')

    if not is_file_like and os.path.isfile(filename):
        file_format = os.path.splitext(filename)[-1]
        if file_format not in SUPPORTED_FORMATS:
            if file_format == '.xls':
                msg = ('openpyxl does not support the old .xls file format, '
                       'please use xlrd to read this file, or convert it to '
                       'the more recent .xlsx file format.')
            elif file_format == '.xlsb':
                msg = ('openpyxl does not support binary format .xlsb, '
                       'please convert this file to .xlsx format if you want '
                       'to open it with openpyxl')
            else:
                msg = ('openpyxl does not support %s file format, '
                       'please check you can open '
                       'it with Excel first. '
                       'Supported formats are: %s') % (file_format,
                                                       ','.join(SUPPORTED_FORMATS))
            raise InvalidFileException(msg)


    if is_file_like:
        # fileobject must have been opened with 'rb' flag
        # it is required by zipfile
        if getattr(filename, 'encoding', None) is not None:
            raise IOError("File-object must be opened in binary mode")

    try:
        archive = ZipFile(filename, 'r', ZIP_DEFLATED)
    except BadZipfile:
        f = repair_central_directory(filename, is_file_like)
        archive = ZipFile(f, 'r', ZIP_DEFLATED)
    return archive


def load_workbook(filename, read_only=False, keep_vba=KEEP_VBA, data_only=False, guess_types=False):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`

    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool

    :param keep_vba: preseve vba content (this does NOT mean you can use it)
    :type keep_vba: bool

    :param guess_types: guess cell content type and do not read it from the file
    :type guess_types: bool

    :param data_only: controls whether cells with formulae have either the formula (default) or the value stored the last time Excel read the sheet
    :type data_only: bool

    :rtype: :class:`openpyxl2.workbook.Workbook`

    .. note::

        When using lazy load, all worksheets will be :class:`openpyxl.worksheet.iter_worksheet.IterableWorksheet`
        and the returned workbook will be read-only.

    """
    archive = _validate_archive(filename)
    read_only = read_only

    parser = WorkbookParser(archive)
    parser.parse()
    wb = parser.wb
    wb._data_only = data_only
    wb._read_only = read_only
    wb.guess_types = guess_types
    wb._sheets = []

    if read_only and guess_types:
        warnings.warn('Data types are not guessed when using iterator reader')

    valid_files = archive.namelist()

    # If are going to preserve the vba then attach a copy of the archive to the
    # workbook so that is available for the save.
    if keep_vba:
        wb.vba_archive = ZipFile(BytesIO(), 'a', ZIP_DEFLATED)
        for name in archive.namelist():
            wb.vba_archive.writestr(name, archive.read(name))


    if read_only:
        wb._archive = ZipFile(filename)

    # get workbook-level information
    if ARC_CORE in valid_files:
        src = fromstring(archive.read(ARC_CORE))
        wb.properties = DocumentProperties.from_tree(src)

    # is workbook a template or note
    src = archive.read(ARC_CONTENT_TYPES)
    root = fromstring(src)
    package = Manifest.from_tree(root)
    wb.is_template = XLTX in package or XLTM in package

    shared_strings = []
    ct = package.find(SHARED_STRINGS)
    if ct is not None:
        strings_path = ct.PartName[1:]
        shared_strings = read_string_table(archive.read(strings_path))


    if ARC_THEME in valid_files:
        wb.loaded_theme = archive.read(ARC_THEME)

    apply_stylesheet(archive, wb) # bind styles to workbook

    # get worksheets
    for sheet, rel in parser.find_sheets():
        sheet_name = sheet.name
        worksheet_path = rel.target
        rels_path = get_rels_path(worksheet_path)
        rels = []
        if rels_path in valid_files:
            rels = get_dependents(archive, rels_path)

        if not worksheet_path in valid_files:
            continue

        if read_only:
            ws = ReadOnlyWorksheet(wb, sheet_name, worksheet_path, None,
                                       shared_strings)

            wb._add_sheet(ws)
        else:
            fh = archive.open(worksheet_path)
            parser = WorkSheetParser(wb, sheet_name, fh, shared_strings)
            parser.parse()
            ws = wb[sheet_name]

            if rels:
                # assign any comments to cells
                for r in rels.find(COMMENTS_NS):
                    src = archive.read(r.target)
                    comment_sheet = CommentSheet.from_tree(fromstring(src))
                    for ref, comment in comment_sheet.comments:
                        ws.cell(coordinate=ref).comment = comment

                # preserve link to VML file if VBA
                if (
                    wb.vba_archive is not None
                    and ws.legacy_drawing is not None
                    ):
                    ws.legacy_drawing = rels[ws.legacy_drawing].target

        ws.sheet_state = sheet.state

    wb._differential_styles = [] # reset
    wb._named_ranges = list(read_named_ranges(archive.read(ARC_WORKBOOK), wb))

    if EXTERNAL_LINK in package:
        rels = get_dependents(archive, ARC_WORKBOOK_RELS)
        wb._external_links = list(detect_external_links(rels, archive))


    archive.close()
    return wb
