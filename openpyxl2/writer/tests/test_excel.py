from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl2 import Workbook
from openpyxl2.worksheet.table import Table


def test_tables():
    wb = Workbook()
    ws = wb.active
    ws._rels = []
    t = Table(displayName="Table1", ref="A1:D10")
    ws.add_table(t)

    out = BytesIO()
    archive = ZipFile(out, "w")

    from ..excel import ExcelWriter
    writer = ExcelWriter(wb)
    writer._write_worksheets(archive)

    assert t.path[1:] in archive.namelist()
