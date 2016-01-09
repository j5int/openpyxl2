from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl2.comments import Comment
from openpyxl2.workbook import Workbook
from openpyxl2.worksheet import Worksheet

def test_init():
    wb = Workbook()
    ws = Worksheet(wb)
    c = Comment("text", "author")
    ws.cell(coordinate="A1").comment = c
    assert c._parent == ws.cell(coordinate="A1")
    assert c.text == "text"
    assert c.author == "author"
