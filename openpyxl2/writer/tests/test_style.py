# Copyright (c) 2010-2014 openpyxl

import pytest

from openpyxl2.styles.borders import Border, Side
from openpyxl2.styles.fills import GradientFill
from openpyxl2.styles.colors import Color
from openpyxl2.writer.styles import StyleWriter
from openpyxl2.tests.helper import get_xml, compare_xml


class DummyWorkbook:

    style_properties = []


def test_write_gradient_fill():
    fill = GradientFill(degree=90, stop=[Color(theme=0), Color(theme=4)])
    writer = StyleWriter(DummyWorkbook())
    writer._write_gradient_fill(writer._root, fill)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0" ?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <gradientFill degree="90" type="linear">
    <stop position="0">
      <color theme="0"/>
    </stop>
    <stop position="1">
      <color theme="4"/>
    </stop>
  </gradientFill>
</styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_borders():
    borders = Border()
    writer = StyleWriter(DummyWorkbook())
    writer._write_border(writer._root, borders)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <border>
    <left/>
    <right/>
    <top/>
    <bottom/>
    <diagonal/>
  </border>
</styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_font():
    wb = DummyWorkbook()
    from openpyxl2.styles import Font
    ft = Font(name='Calibri', charset=204, vertAlign='superscript')
    writer = StyleWriter(wb)
    writer._write_font(writer._root, ft)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sz val="11.0"></sz>
      <color rgb="00000000"></color>
      <name val="Calibri"></name>
      <family val="2"></family>
      <vertAlign val="superscript"></vertAlign>
      <charset val="204"></charset>
</styleSheet>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff
