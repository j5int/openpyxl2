from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl


# package imports
from openpyxl2.tests.helper import compare_xml
from openpyxl2.writer.theme import write_theme


def test_write_theme(datadir):
    datadir.chdir()
    content = write_theme()
    with open( 'theme1.xml') as expected:
        diff = compare_xml(content, expected.read())
        assert diff is None, diff
