# coding=utf8

# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

# Python stdlib imports
from datetime import datetime
from io import BytesIO
from tempfile import NamedTemporaryFile
from zipfile import ZipFile

import pytest

# compatibility imports
from openpyxl2.compat import unicode

# package imports
from openpyxl2.collections import IndexedList
from openpyxl2.worksheet import Worksheet
from openpyxl2.workbook import Workbook
from openpyxl2.styles import numbers, Style
from openpyxl2.reader.worksheet import read_worksheet
from openpyxl2.reader.excel import load_workbook
from openpyxl2.reader.workbook import read_workbook_code_name
from openpyxl2.exceptions import InvalidFileException
from openpyxl2.date_time import CALENDAR_WINDOWS_1900, CALENDAR_MAC_1904
from openpyxl2.xml.constants import ARC_WORKBOOK


def test_read_standalone_worksheet(datadir):

    class DummyWb(object):

        encoding = 'utf-8'

        excel_base_date = CALENDAR_WINDOWS_1900
        _guess_types = True
        data_only = False

        def get_sheet_by_name(self, value):
            return None

        def get_sheet_names(self):
            return []

    datadir.join("reader").chdir()
    ws = None
    shared_strings = IndexedList(['hello'])

    with open('sheet2.xml') as src:
        ws = read_worksheet(src.read(), DummyWb(), 'Sheet 2', shared_strings,
                            {1: Style()})
        assert isinstance(ws, Worksheet)
        assert ws.cell('G5').value == 'hello'
        assert ws.cell('D30').value == 30
        assert ws.cell('K9').value == 0.09


@pytest.fixture
def standard_workbook(datadir):
    datadir.join("genuine").chdir()
    return load_workbook("empty.xlsx")


def test_read_standard_workbook(standard_workbook):
    wb = standard_workbook
    assert isinstance(wb, Workbook)


def test_read_standard_workbook_from_fileobj(datadir):
    datadir.join("genuine").chdir()
    fo = open('empty.xlsx', mode='rb')
    wb = load_workbook(fo)
    assert isinstance(wb, Workbook)


def test_read_worksheet(standard_workbook):
    wb = standard_workbook
    sheet2 = wb['Sheet2 - Numbers']
    assert isinstance(sheet2, Worksheet)
    assert 'This is cell G5' == sheet2['G5'].value
    assert 18 == sheet2['D18'].value
    assert sheet2['G9'].value is True
    assert sheet2['G10'].value is False


def test_read_nostring_workbook(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook('empty-no-string.xlsx')
    assert isinstance(wb, Workbook)


def test_read_empty_file(datadir):
    datadir.join("reader").chdir()
    with pytest.raises(InvalidFileException):
        load_workbook('null_file.xlsx')


@pytest.mark.parametrize("cell, number_format",
                    [
                        ('A1', numbers.FORMAT_GENERAL),
                        ('A2', numbers.FORMAT_DATE_XLSX14),
                        ('A3', numbers.FORMAT_NUMBER_00),
                        ('A4', numbers.FORMAT_DATE_TIME3),
                        ('A5', numbers.FORMAT_PERCENTAGE_00),
                    ]
                    )
def test_read_general_style(datadir, cell, number_format):
    datadir.join("genuine").chdir()
    wb = load_workbook('empty-with-styles.xlsx')
    return wb["Sheet1"]
    ws = workbook_with_styles
    assert ws[cell].number_format == number_format



@pytest.mark.parametrize("filename, epoch",
                         [
                             ("date_1900.xlsx", CALENDAR_WINDOWS_1900),
                             ("date_1904.xlsx",  CALENDAR_MAC_1904),
                         ]
                         )
def test_read_win_base_date(datadir, filename, epoch):
    datadir.join("reader").chdir()
    wb = load_workbook(filename)
    assert wb.properties.excel_base_date == epoch
    ws = wb["Sheet1"]
    assert ws['A1'].value == datetime(2011, 10, 31)


def test_repair_central_directory():
    from openpyxl2.reader.excel import repair_central_directory, CENTRAL_DIRECTORY_SIGNATURE

    data_a = b"foobarbaz" + CENTRAL_DIRECTORY_SIGNATURE
    data_b = b"bazbarfoo1234567890123456890"

    # The repair_central_directory looks for a magic set of bytes
    # (CENTRAL_DIRECTORY_SIGNATURE) and strips off everything 18 bytes past the sequence
    f = repair_central_directory(BytesIO(data_a + data_b), True)
    assert f.read() == data_a + data_b[:18]

    f = repair_central_directory(BytesIO(data_b), True)
    assert f.read() == data_b


def test_read_no_theme(datadir):
    datadir.join("genuine").chdir()
    wb = load_workbook('libreoffice_nrt.xlsx')
    assert wb


def test_read_cell_formulae(datadir):
    from openpyxl2.reader.worksheet import fast_parse
    datadir.join("reader").chdir()
    wb = Workbook()
    ws = wb.active
    fast_parse(ws, open( "worksheet_formula.xml"), ['', ''], {}, None)
    b1 = ws['B1']
    assert b1.data_type == 'f'
    assert b1.value == '=CONCATENATE(A1,A2)'
    a6 = ws['A6']
    assert a6.data_type == 'f'
    assert a6.value == '=SUM(A4:A5)'


def test_read_complex_formulae(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('formulae.xlsx')
    ws = wb.get_active_sheet()

    # Test normal forumlae
    assert ws.cell('A1').data_type != 'f'
    assert ws.cell('A2').data_type != 'f'
    assert ws.cell('A3').data_type == 'f'
    assert 'A3' not in ws.formula_attributes
    assert ws.cell('A3').value == '=12345'
    assert ws.cell('A4').data_type == 'f'
    assert 'A4' not in ws.formula_attributes
    assert ws.cell('A4').value == '=A2+A3'
    assert ws.cell('A5').data_type == 'f'
    assert 'A5' not in ws.formula_attributes
    assert ws.cell('A5').value == '=SUM(A2:A4)'

    # Test unicode
    expected = '=IF(ISBLANK(B16), "Düsseldorf", B16)'
    # Hack to prevent pytest doing it's own unicode conversion
    try:
        expected = unicode(expected, "UTF8")
    except TypeError:
        pass
    assert ws['A16'].value == expected

    # Test shared forumlae
    assert ws.cell('B7').data_type == 'f'
    assert ws.formula_attributes['B7']['t'] == 'shared'
    assert ws.formula_attributes['B7']['si'] == '0'
    assert ws.formula_attributes['B7']['ref'] == 'B7:E7'
    assert ws.cell('B7').value == '=B4*2'
    assert ws.cell('C7').data_type == 'f'
    assert ws.formula_attributes['C7']['t'] == 'shared'
    assert ws.formula_attributes['C7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['C7']
    assert ws.cell('C7').value == '='
    assert ws.cell('D7').data_type == 'f'
    assert ws.formula_attributes['D7']['t'] == 'shared'
    assert ws.formula_attributes['D7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['D7']
    assert ws.cell('D7').value == '='
    assert ws.cell('E7').data_type == 'f'
    assert ws.formula_attributes['E7']['t'] == 'shared'
    assert ws.formula_attributes['E7']['si'] == '0'
    assert 'ref' not in ws.formula_attributes['E7']
    assert ws.cell('E7').value == '='

    # Test array forumlae
    assert ws.cell('C10').data_type == 'f'
    assert 'ref' not in ws.formula_attributes['C10']['ref']
    assert ws.formula_attributes['C10']['t'] == 'array'
    assert 'si' not in ws.formula_attributes['C10']
    assert ws.formula_attributes['C10']['ref'] == 'C10:C14'
    assert ws.cell('C10').value == '=SUM(A10:A14*B10:B14)'
    assert ws.cell('C11').data_type != 'f'


def test_data_only(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook('formulae.xlsx', data_only=True)
    ws = wb.active
    ws.parent.data_only = True
    # Test cells returning values only, not formulae
    assert ws.formula_attributes == {}
    assert ws['A2'].data_type == 'n' and ws.cell('A2').value == 12345
    assert ws['A3'].data_type == 'n' and ws.cell('A3').value == 12345
    assert ws['A4'].data_type == 'n' and ws.cell('A4').value == 24690
    assert ws['A5'].data_type == 'n' and ws.cell('A5').value == 49380


def test_guess_types(datadir):
    datadir.join("genuine").chdir()
    for guess, dtype in ((True, float), (False, unicode)):
        wb = load_workbook('guess_types.xlsx', guess_types=guess)
        ws = wb.get_active_sheet()
        assert isinstance(ws.cell('D2').value, dtype), 'wrong dtype (%s) when guess type is: %s (%s instead)' % (dtype, guess, type(ws.cell('A1').value))


def test_get_xml_iter():
    #1 file object
    #2 stream (file-like)
    #3 string
    #4 zipfile
    from openpyxl2.reader.worksheet import _get_xml_iter
    from tempfile import TemporaryFile

    FUT = _get_xml_iter
    s = b""
    stream = FUT(s)
    assert isinstance(stream, BytesIO), type(stream)

    u = unicode(s)
    stream = FUT(u)
    assert isinstance(stream, BytesIO), type(stream)

    f = TemporaryFile(mode='rb+', prefix='openpyxl.', suffix='.unpack.temp')
    stream = FUT(f)
    assert stream == f
    f.close()

    t = TemporaryFile()
    z = ZipFile(t, mode="w")
    z.writestr("test", "whatever")
    stream = FUT(z.open("test"))
    assert hasattr(stream, "read")
    # z.close()
    try:
        z.close()
    except IOError:
        # you can't just close zipfiles in Windows
        if z.fp is not None:
            z.fp.close() # python 2.6
        else:
            z.close() # python 2.7


def test_read_autofilter(datadir):
    datadir.join("reader").chdir()
    wb = load_workbook("bug275.xlsx")
    ws = wb.active
    assert ws.auto_filter.ref == 'A1:B6'


class TestBadFormats:

    def test_xlsb(self):
        tmp = NamedTemporaryFile(suffix='.xlsb')
        with pytest.raises(InvalidFileException):
            load_workbook(filename=tmp.name)

    def test_xls(self):
        tmp = NamedTemporaryFile(suffix='.xls')
        with pytest.raises(InvalidFileException):
            load_workbook(filename=tmp.name)

    def test_no(self):
        tmp = NamedTemporaryFile(suffix='.no-format')
        with pytest.raises(InvalidFileException):
            load_workbook(filename=tmp.name)
