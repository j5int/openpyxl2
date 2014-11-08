# Copyright (c) 2010-2014 openpyxl
#

# Python stdlib imports
from datetime import time, datetime, timedelta, date

# 3rd party imports
import pytest

# package imports
from openpyxl2.compat import safe_string
from openpyxl2.worksheet import Worksheet
from openpyxl2.workbook import Workbook
from openpyxl2.exceptions import (
    CellCoordinatesException,
    )
from openpyxl2.date_time import CALENDAR_WINDOWS_1900
from openpyxl2.cell import (
    column_index_from_string,
    coordinate_from_string,
    get_column_letter,
    Cell,
    absolute_coordinate
    )
from openpyxl2.comments import Comment
from openpyxl2.styles import numbers

import decimal

@pytest.fixture
def build_dummy_worksheet():

    class Ws(object):
        class Wb(object):
            excel_base_date = CALENDAR_WINDOWS_1900
        encoding = 'utf-8'
        parent = Wb()
        title = "Dummy Worksheet"

    return Ws()


def test_coordinates():
    column, row = coordinate_from_string('ZF46')
    assert "ZF" == column
    assert 46 == row


def test_invalid_coordinate():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AAA')

def test_zero_row():
    with pytest.raises(CellCoordinatesException):
        coordinate_from_string('AQ0')

def test_absolute():
    assert '$ZF$51' == absolute_coordinate('ZF51')

def test_absolute_multiple():

    assert '$ZF$51:$ZF$53' == absolute_coordinate('ZF51:ZF$53')

@pytest.mark.parametrize("column, idx",
                         [
                         ('j', 10),
                         ('Jj', 270),
                         ('JJj', 7030),
                         ('A', 1),
                         ('Z', 26),
                         ('AA', 27),
                         ('AZ', 52),
                         ('BA', 53),
                         ('BZ',  78),
                         ('ZA',  677),
                         ('ZZ',  702),
                         ('AAA',  703),
                         ('AAZ',  728),
                         ('ABC',  731),
                         ('AZA', 1353),
                         ('ZZA', 18253),
                         ('ZZZ', 18278),
                         ]
                         )
def test_column_index(column, idx):
    assert column_index_from_string(column) == idx


@pytest.mark.parametrize("column",
                         ('JJJJ', '', '$', '1',)
                         )
def test_bad_column_index(column):
    with pytest.raises(ValueError):
        column_index_from_string(column)


@pytest.mark.parametrize("value", (0, 18729))
def test_column_letter_boundries(value):
    with pytest.raises(ValueError):
        get_column_letter(value)

@pytest.mark.parametrize("value, expected",
                         [
                        (18278, "ZZZ"),
                        (7030, "JJJ"),
                        (28, "AB"),
                        (27, "AA"),
                        (26, "Z")
                         ]
                         )
def test_column_letter(value, expected):
    assert get_column_letter(value) == expected


def test_initial_value():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1, value='17.5')
    assert cell.TYPE_STRING == cell.data_type


class TestCellValueTypes(object):

    @classmethod
    def setup_class(cls):
        wb = Workbook()
        ws = Worksheet(wb)
        cls.cell = Cell(ws, 'A', 1)

    def test_ctor(self):
        cell = self.cell
        assert cell.data_type == 'n'
        assert cell.column == 'A'
        assert cell.row == 1
        assert cell.coordinate == "A1"
        assert cell.value is None
        assert isinstance(cell.parent, Worksheet)
        assert cell.xf_index == 0
        assert cell.comment is None


    @pytest.mark.parametrize("datatype", ['n', 'd', 's', 'b', 'f', 'e'])
    def test_null(self, datatype):
        self.cell.data_type = datatype
        assert self.cell.data_type == datatype
        self.cell.value = None
        assert self.cell.data_type == 'n'


    @pytest.mark.parametrize("value, expected",
                             [
                                 (42, 42),
                                 ('4.2', 4.2),
                                 ('-42.000', -42),
                                 ( '0', 0),
                                 (0, 0),
                                 ( 0.0001, 0.0001),
                                 ('0.9999', 0.9999),
                                 ('99E-02', 0.99),
                                 ( 1e1, 10),
                                 ('4', 4),
                                 ('-1E3', -1000),
                                 ('2e+2', 200),
                                 (4, 4),
                                 (decimal.Decimal('3.14'), decimal.Decimal('3.14')),
                                 ('3.1%', 0.031),
                             ]
                            )
    def test_numeric(self, value, expected):
        self.cell.parent.parent._guess_types = True
        self.cell.value = value
        assert self.cell.internal_value == expected
        assert self.cell.data_type == 'n'


    @pytest.mark.parametrize("value", ['hello', ".", '0800'])
    def test_string(self, value):
        self.cell.value = 'hello'
        assert self.cell.data_type == 's'


    @pytest.mark.parametrize("value", ['=42', '=if(A1<4;-1;1)'])
    def test_formula(self, value):
        self.cell.value = value
        assert self.cell.data_type == 'f'


    @pytest.mark.parametrize("value", [True, False])
    def test_boolean(self, value):
        self.cell.value = True
        assert self.cell.data_type == 'b'


    @pytest.mark.parametrize("error_string", Cell.ERROR_CODES)
    def test_error_codes(self, error_string):
        self.cell.value = error_string
        assert self.cell.data_type == 'e'


    @pytest.mark.parametrize("value, internal, number_format",
                             [
                                 (
                                     datetime(2010, 7, 13, 6, 37, 41),
                                     40372.27616898148,
                                     "yyyy-mm-dd h:mm:ss"
                                 ),
                                 (
                                     date(2010, 7, 13),
                                     40372,
                                     "yyyy-mm-dd"
                                 ),
                             ]
                             )
    def test_insert_date(self, value, internal, number_format):
        self.cell.value = value
        assert self.cell.data_type == 'n'
        assert self.cell.internal_value == internal
        assert self.cell.is_date
        assert self.cell.number_format == number_format


    def test_empty_cell_formatted_as_date(self):
        self.cell.value = datetime.today()
        self.cell.value = None
        assert self.cell.is_date
        assert self.cell.value is None


def test_set_bad_type():
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)
    with pytest.raises(ValueError):
        cell.set_explicit_value(1, 'q')


def test_illegal_chacters():
    from openpyxl2.exceptions import IllegalCharacterError
    from openpyxl2.compat import range
    from itertools import chain
    ws = build_dummy_worksheet()
    cell = Cell(ws, 'A', 1)

    # The bytes 0x00 through 0x1F inclusive must be manually escaped in values.

    illegal_chrs = chain(range(9), range(11, 13), range(14, 32))
    for i in illegal_chrs:
        with pytest.raises(IllegalCharacterError):
            cell.value = chr(i)

        with pytest.raises(IllegalCharacterError):
            cell.value = "A {0} B".format(chr(i))

    cell.value = chr(33)
    cell.value = chr(9)  # Tab
    cell.value = chr(10)  # Newline
    cell.value = chr(13)  # Carriage return
    cell.value = " Leading and trailing spaces are legal "


values = (
    ('30:33.865633336', [('', '', '', '30', '33', '865633')]),
    ('03:40:16', [('03', '40', '16', '', '', '')]),
    ('03:40', [('03', '40', '',  '', '', '')]),
    ('55:72:12', []),
    )
@pytest.mark.parametrize("value, expected",
                             values)
def test_time_regex(value, expected):
    from openpyxl2.cell.cell import TIME_REGEX
    m = TIME_REGEX.findall(value)
    assert m == expected


values = (
    ('03:40:16', time(3, 40, 16)),
    ('03:40', time(3, 40)),
    ('30:33.865633336', time(0, 30, 33, 865633))
)
@pytest.mark.parametrize("value, expected",
                         values)
def test_time(value, expected):
    wb = Workbook(guess_types=True)
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = value
    assert cell.value == expected
    assert cell.TYPE_NUMERIC == cell.data_type


def test_timedelta():

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = timedelta(days=1, hours=3)
    assert cell.value == 1.125
    assert cell.TYPE_NUMERIC == cell.data_type


def test_date_format_on_non_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    cell.value = 'testme'
    assert 'testme' == cell.value

def test_set_get_date():
    today = datetime(2010, 1, 18, 14, 15, 20, 1600)
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = today
    assert today == cell.value

def test_repr():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    assert repr(cell), '<Cell Sheet1.A1>' == 'Got bad repr: %s' % repr(cell)

def test_is_date():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)
    cell.value = datetime.now()
    assert cell.is_date == True
    cell.value = 'testme'
    assert 'testme' == cell.value
    assert cell.is_date is False

def test_is_not_date_color_format():

    wb = Workbook()
    ws = Worksheet(wb)
    cell = Cell(ws, 'A', 1)

    cell.value = -13.5
    cell.style = cell.style.copy(number_format='0.00_);[Red]\(0.00\)')

    assert cell.is_date is False

def test_comment_count():
    wb = Workbook()
    ws = Worksheet(wb)
    cell = ws.cell(coordinate="A1")
    assert ws._comment_count == 0
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = Comment("text", "author")
    assert ws._comment_count == 1
    cell.comment = None
    assert ws._comment_count == 0
    cell.comment = None
    assert ws._comment_count == 0

def test_comment_assignment():
    wb = Workbook()
    ws = Worksheet(wb)
    c = Comment("text", "author")
    ws.cell(coordinate="A1").comment = c
    with pytest.raises(AttributeError):
        ws.cell(coordinate="A2").commment = c
    ws.cell(coordinate="A2").comment = Comment("text2", "author2")
    with pytest.raises(AttributeError):
        ws.cell(coordinate="A1").comment = ws.cell(coordinate="A2").comment
    # this should orphan c, so that assigning it to A2 does not raise AttributeError
    ws.cell(coordinate="A1").comment = None
    ws.cell(coordinate="A2").comment = c

def test_cell_offset():
    wb = Workbook()
    ws = Worksheet(wb)
    assert ws['B15'].offset(2, 1).coordinate == 'C17'


class TestEncoding:

    try:
        # Python 2
        pound = unichr(163)
    except NameError:
        # Python 3
        pound = chr(163)
    test_string = ('Compound Value (' + pound + ')').encode('latin1')

    def test_bad_encoding(self):
        wb = Workbook()
        ws = wb.active
        cell = ws['A1']
        with pytest.raises(UnicodeDecodeError):
            cell.check_string(self.test_string)
        with pytest.raises(UnicodeDecodeError):
            cell.value = self.test_string

    def test_good_encoding(self):
        wb = Workbook(encoding='latin1')
        ws = wb.active
        cell = ws['A1']
        cell.value = self.test_string


def test_style():
    from openpyxl2.styles import Font, Style
    wb = Workbook()
    ws = wb.active
    cell = ws['A1']
    new_style = Style(font=Font(bold=True))
    cell.style = new_style
    assert new_style in wb.shared_styles
