# Copyright (c) 2010-2014 openpyxl

import py.test

from openpyxl2.xml.constants import SHEET_MAIN_NS
from openpyxl2.xml.functions import safe_iterator, fromstring, iterparse

class DummyWorkbook:

    _guess_types = False
    data_only = False


class DummyWorksheet:

    def __init__(self):
        self.parent = DummyWorkbook()
        self.column_dimensions = {}
        self._styles = {}


def test_parse_col_dimensions(datadir):
    from openpyxl2.reader.worksheet import WorkSheetParser

    datadir.chdir()
    ws = DummyWorksheet()

    with open("complex-styles-worksheet.xml") as src:
        parser = WorkSheetParser(ws, src, {}, {})
        tree = iterparse(parser.source)
        for _, tag in tree:
            cols = safe_iterator(tag, '{%s}col' % SHEET_MAIN_NS)
            for col in cols:
                parser.parse_column_dimensions(col)
    assert set(ws.column_dimensions.keys()) == set(['A', 'C', 'E', 'I', 'G'])
    assert dict(ws.column_dimensions['A']) == {'max': '1', 'min': '1',
                                               'customWidth': '1',
                                               'width': '31.1640625'}


def test_sheet_protection(datadir):
    from .. worksheet import WorkSheetParser

    datadir.chdir()
    ws = DummyWorksheet()

    with open("protected_sheet.xml") as src:
        parser = WorkSheetParser(ws, src, {}, {})
        tree = iterparse(parser.source)
        for _, tag in tree:
            prot = safe_iterator(tag, '{%s}sheetProtection' % SHEET_MAIN_NS)
            for el in prot:
                parser.parse_sheet_protection(el)
    assert dict(ws.protection) == {
        'autoFilter': '1', 'deleteColumns': '1',
        'deleteRows': '1', 'formatCells': '1', 'formatColumns': '1', 'formatRows':
        '1', 'insertColumns': '1', 'insertHyperlinks': '1', 'insertRows': '1',
        'objects': '0', 'password': 'DAA7', 'pivotTables': '1', 'scenarios': '0',
        'selectLockedCells': '0', 'selectUnlockedCells': '0', 'sheet': '1', 'sort':
        '1'
    }

