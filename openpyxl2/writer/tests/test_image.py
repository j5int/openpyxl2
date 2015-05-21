from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


# stdlib imports
from io import BytesIO
import zipfile

import pytest

# package imports
from openpyxl2.workbook import Workbook
from openpyxl2.writer.excel import ExcelWriter


@pytest.mark.pil_required
def test_write_images(datadir):
    datadir.chdir()
    wb = Workbook()
    ew = ExcelWriter(workbook=wb)
    from openpyxl2.drawing.image import Image
    img = Image("plain.png")
    wb._images.append(img)

    buf = BytesIO()

    archive = zipfile.ZipFile(buf, 'w')
    ew._write_images(archive)
    archive.close()

    buf.seek(0)
    archive = zipfile.ZipFile(buf, 'r')
    zipinfo = archive.infolist()
    assert len(zipinfo) == 1
    assert zipinfo[0].filename == 'xl/media/image1.png'
