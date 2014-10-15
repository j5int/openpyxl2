from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from io import BytesIO

from openpyxl2.exceptions import InvalidFileException
from .. excel import load_workbook

import pytest


def test_read_empty_file(datadir):
    datadir.chdir()
    with pytest.raises(InvalidFileException):
        load_workbook('null_file.xlsx')


def test_repair_central_directory():
    from ..excel import repair_central_directory, CENTRAL_DIRECTORY_SIGNATURE

    data_a = b"foobarbaz" + CENTRAL_DIRECTORY_SIGNATURE
    data_b = b"bazbarfoo1234567890123456890"

    # The repair_central_directory looks for a magic set of bytes
    # (CENTRAL_DIRECTORY_SIGNATURE) and strips off everything 18 bytes past the sequence
    f = repair_central_directory(BytesIO(data_a + data_b), True)
    assert f.read() == data_a + data_b[:18]

    f = repair_central_directory(BytesIO(data_b), True)
    assert f.read() == data_b
