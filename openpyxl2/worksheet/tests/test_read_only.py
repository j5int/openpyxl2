from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml


class DummyWorkbook:

    excel_base_date = None
    sheetnames = []


@pytest.fixture
def ReadOnlyWorksheet():
    from ..read_only import ReadOnlyWorksheet
    return ReadOnlyWorksheet


class TestReadOnlyWorksheet:

    def test_from_xml(self, datadir, ReadOnlyWorksheet):

        datadir.chdir()

        ws = ReadOnlyWorksheet(DummyWorkbook(), "Sheet", "", "sheet_inline_strings.xml", [])
        cells = tuple(ws.iter_rows(min_row=1, min_col=1, max_row=1, max_col=1))
        assert len(cells) == 1
        assert cells[0][0].value == "col1"
