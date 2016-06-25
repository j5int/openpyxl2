from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

import pytest

from openpyxl2.utils.indexed_list import IndexedList
from ..import (
    Font,
    Border,
    PatternFill,
    Alignment,
    Protection
)
from ..named_styles import (
    NamedStyles,
    NamedStyle,
)


def test_descriptor():
    from ..styleable import StyleDescriptor
    from ..cell_style import StyleArray
    from ..fonts import Font

    class Styled(object):

        font = StyleDescriptor('_fonts', "fontId")

        def __init__(self):
            self._style = StyleArray()
            self.parent = DummyWorksheet()

    styled = Styled()
    styled.font = Font()
    assert styled.font == Font()


class DummyWorkbook:

    _fonts = IndexedList()
    _fills = IndexedList()
    _borders = IndexedList()
    _protections = IndexedList()
    _alignments = IndexedList()
    _number_formats = IndexedList()
    _named_styles = NamedStyles()


class DummyWorksheet:

    parent = DummyWorkbook()


@pytest.fixture
def StyleableObject():
    from .. styleable import StyleableObject
    return StyleableObject


def test_has_style(StyleableObject):
    so = StyleableObject(sheet=DummyWorksheet())
    assert not so.has_style
    so.number_format= 'dd'
    assert so.has_style


class TestNamedStyle:

    def test_assign(self, StyleableObject):
        ws = DummyWorksheet()
        wb = ws.parent
        style = NamedStyle(name='Standard')
        wb._named_styles.append(style)

        so = StyleableObject(sheet=ws)
        so.style = 'Standard'


    def test_unknown_style(self, StyleableObject):
        so = StyleableObject(sheet=DummyWorksheet())
        with pytest.raises(ValueError):
            so.style = "Financial"


    def test_read(self, StyleableObject):
        ws = DummyWorksheet()
        wb = ws.parent

        style = NamedStyle(name='Red')
        wb._named_styles.append(style)

        so = StyleableObject(sheet=ws, style_array=list(range(9)))
        so._style.xfId = 1
        assert so.style == "Red"
