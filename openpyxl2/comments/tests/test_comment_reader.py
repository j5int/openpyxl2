from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl


from zipfile import ZipFile

from openpyxl2.workbook import Workbook
from openpyxl2.worksheet import Worksheet
from .. import reader
from openpyxl2.reader.excel import load_workbook
from openpyxl2.xml.functions import fromstring

import pytest


@pytest.mark.parametrize("cell, author, text",
                         [
                             ['A1', 'Cuke', 'Cuke:\nFirst Comment'],
                             ['D1', 'Cuke', 'Cuke:\nSecond Comment'],
                             ['A2', 'Not Cuke', 'Not Cuke:\nThird Comment']
                         ]
                         )
def test_read_comments(datadir, cell, author, text):
    datadir.chdir()
    with open("comments2.xml") as src:
        xml = src.read()

    wb = Workbook()
    ws = Worksheet(wb)
    reader.read_comments(ws, xml)
    comment = ws[cell].comment
    assert comment.author == author
    assert comment.text == text


def test_get_comments_file(datadir):
    datadir.chdir()
    archive = ZipFile('comments.xlsx')
    assert reader.get_comments_file('xl/worksheets/sheet1.xml', archive) == 'xl/comments1.xml'
    assert reader.get_comments_file('xl/worksheets/sheet3.xml', archive) == 'xl/comments2.xml'
    assert reader.get_comments_file('xl/worksheets/sheet2.xml', archive) is None


def test_comments_cell_association(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx')
    assert wb['Sheet1'].cell(coordinate="A1").comment.author == "Cuke"
    assert wb['Sheet1'].cell(coordinate="A1").comment.text == "Cuke:\nFirst Comment"
    assert wb['Sheet2'].cell(coordinate="A1").comment is None
    assert wb['Sheet1'].cell(coordinate="D1").comment.text == "Cuke:\nSecond Comment"


def test_comments_with_iterators(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx', read_only=True)
    ws = wb['Sheet1']
    with pytest.raises(AttributeError):
        assert ws.cell(coordinate="A1").comment.author == "Cuke"
