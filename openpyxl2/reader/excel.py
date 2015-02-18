from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

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
from openpyxl2.exceptions import OpenModeError, InvalidFileException
from openpyxl2.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CORE,
    ARC_WORKBOOK,
    ARC_STYLE,
    ARC_THEME,
    SHARED_STRINGS,
    EXTERNAL_LINK,
)

from openpyxl2.workbook import Workbook, DocumentProperties
from openpyxl2.workbook.names.external import detect_external_links
from openpyxl2.workbook.names.named_range import read_named_ranges
from openpyxl2.reader.strings import read_string_table
from openpyxl2.reader.style import read_style_table
from openpyxl2.reader.workbook import (
    read_content_types,
    read_properties_core,
    read_excel_base_date,
    detect_worksheets,
    read_rels,
    read_workbook_code_name,
)
from openpyxl2.reader.worksheet import read_worksheet
from openpyxl2.reader.comments import read_comments, get_comments_file
# Use exc_info for Python 2 compatibility with "except Exception[,/ as] e"


CENTRAL_DIRECTORY_SIGNATURE = b'\x50\x4b\x05\x06'
SUPPORTED_FORMATS = ('.xlsx', '.xlsm')


def repair_central_directory(zipFile, is_file_instance):
    ''' trims trailing data from the central directory
    code taken from http://stackoverflow.com/a/7457686/570216, courtesy of Uri Cohen
    '''

    f = zipFile if is_file_instance else open(zipFile, 'r+b')
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


def load_workbook(filename, read_only=False, use_iterators=False, keep_vba=KEEP_VBA, guess_types=False, data_only=False):
    """Open the given filename and return the workbook

    :param filename: the path to open or a file-like object
    :type filename: string or a file-like object open in binary mode c.f., :class:`zipfile.ZipFile`

    :param read_only: optimised for reading, content cannot be edited
    :type read_only: bool

    :param use_iterators: use lazy load for cells
    :type use_iterators: bool

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

    is_file_instance = isinstance(filename, file)

    read_only = read_only or use_iterators

    if is_file_instance:
        # fileobject must have been opened with 'rb' flag
        # it is required by zipfile
        if 'b' not in filename.mode:
            raise OpenModeError("File-object must be opened in binary mode")

    try:
        archive = ZipFile(filename, 'r', ZIP_DEFLATED)
    except BadZipfile:
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

        try:
            f = repair_central_directory(filename, is_file_instance)
            archive = ZipFile(f, 'r', ZIP_DEFLATED)
        except BadZipfile:
            e = exc_info()[1]
            raise InvalidFileException(unicode(e))
    except (BadZipfile, RuntimeError, IOError, ValueError):
        e = exc_info()[1]
        raise InvalidFileException(unicode(e))
    wb = Workbook(guess_types=guess_types, data_only=data_only, read_only=read_only)

    if read_only and guess_types:
        warnings.warn('Data types are not guessed when using iterator reader')

    try:
        _load_workbook(wb, archive, filename, read_only, keep_vba)
    except KeyError:
        e = exc_info()[1]
        raise InvalidFileException(unicode(e))

    archive.close()
    return wb


def _load_workbook(wb, archive, filename, read_only, keep_vba):

    valid_files = archive.namelist()

    # If are going to preserve the vba then attach a copy of the archive to the
    # workbook so that is available for the save.
    if keep_vba:
        try:
            f = open(filename, 'rb')
            s = f.read()
            f.close()
        except:
            pos = filename.tell()
            filename.seek(0)
            s = filename.read()
            filename.seek(pos)
        wb.vba_archive = ZipFile(BytesIO(s), 'r')

    if read_only:
        wb._archive = ZipFile(filename)

    # get workbook-level information
    try:
        wb.properties = read_properties_core(archive.read(ARC_CORE))
    except KeyError:
        wb.properties = DocumentProperties()
    wb._read_workbook_settings(archive.read(ARC_WORKBOOK))

    # what content types do we have?
    cts = dict(read_content_types(archive))
    rels = dict

    strings_path = cts.get(SHARED_STRINGS)
    if strings_path is not None:
        if strings_path.startswith("/"):
            strings_path = strings_path[1:]
        shared_strings = read_string_table(archive.read(strings_path))
    else:
        shared_strings = []

    try:
        wb.loaded_theme = archive.read(ARC_THEME)  # some writers don't output a theme, live with it (fixes #160)
    except KeyError:
        assert wb.loaded_theme == None, "even though the theme information is missing there is a theme object ?"

    style_table, color_index, cond_styles = read_style_table(archive.read(ARC_STYLE))
    wb.shared_styles = style_table
    wb.style_properties = {'dxf_list':cond_styles}
    wb.cond_styles = cond_styles

    wb.properties.excel_base_date = read_excel_base_date(xml_source=archive.read(ARC_WORKBOOK))

    # get worksheets
    wb.worksheets = []  # remove preset worksheet
    for sheet in detect_worksheets(archive):
        sheet_name = sheet['title']
        worksheet_path = sheet['path']
        if not worksheet_path in valid_files:
            continue

        if not read_only:
            new_ws = read_worksheet(archive.read(worksheet_path), wb,
                                    sheet_name, shared_strings, style_table,
                                    color_index=color_index,
                                    keep_vba=keep_vba)
        else:
            new_ws = read_worksheet(None, wb, sheet_name, shared_strings,
                                    style_table,
                                    color_index=color_index,
                                    worksheet_path=worksheet_path)

        new_ws.sheet_state = sheet.get('state') or 'visible'
        wb._add_sheet(new_ws)

        if not read_only:
        # load comments into the worksheet cells
            comments_file = get_comments_file(worksheet_path, archive, valid_files)
            if comments_file is not None:
                read_comments(new_ws, archive.read(comments_file))

    wb._named_ranges = list(read_named_ranges(archive.read(ARC_WORKBOOK), wb))

    wb.code_name = read_workbook_code_name(archive.read(ARC_WORKBOOK))

    if EXTERNAL_LINK in cts:
        rels = read_rels(archive)
        wb._external_links = list(detect_external_links(rels, archive))
