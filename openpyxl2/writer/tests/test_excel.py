from io import BytesIO
from string import ascii_letters
from zipfile import ZipFile

import pytest

from openpyxl2.chart import BarChart
from openpyxl2.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl2 import Workbook
from openpyxl2.worksheet.table import Table


@pytest.fixture
def ExcelWriter():
    from ..excel import ExcelWriter
    return ExcelWriter


@pytest.fixture
def archive():
    out = BytesIO()
    return ZipFile(out, "w")


def test_tables(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active
    ws.append(list(ascii_letters))
    ws._rels = []
    t = Table(displayName="Table1", ref="A1:D10")
    ws.add_table(t)


    writer = ExcelWriter(wb, archive)
    writer._write_worksheets()

    assert t.path[1:] in archive.namelist()
    assert t.path in writer.manifest.filenames


def test_drawing(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    drawing = SpreadsheetDrawing()

    writer = ExcelWriter(wb, archive)
    writer._write_drawing(drawing)
    assert drawing.path == '/xl/drawings/drawing1.xml'
    assert drawing.path[1:] in archive.namelist()
    assert drawing.path in writer.manifest.filenames


def test_write_chart(ExcelWriter, archive):
    wb = Workbook()
    ws = wb.active

    chart = BarChart()
    ws.add_chart(chart)

    writer = ExcelWriter(wb, archive)
    writer._write_worksheets()
    assert 'xl/worksheets/sheet1.xml' in archive.namelist()
    assert ws.path in writer.manifest.filenames

    rel = ws._rels["rId1"]
    assert dict(rel) == {'Id': 'rId1', 'Target': '/xl/drawings/drawing1.xml',
                         'Type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}


@pytest.mark.pil_required
def test_write_images(datadir, ExcelWriter, archive):
    from openpyxl2.drawing.image import Image
    datadir.chdir()

    writer = ExcelWriter(None, archive)

    img = Image("plain.png")
    writer._images.append(img)

    writer._write_images()
    archive.close()

    zipinfo = archive.infolist()
    assert len(zipinfo) == 1
    assert zipinfo[0].filename == 'xl/media/image1.png'
    assert 'xl/media/image1.png' in archive.namelist()
