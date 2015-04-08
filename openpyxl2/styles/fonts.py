from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl2.descriptors import Float, Integer, Set, Bool, String, Alias, MinMax, NoneSet
from openpyxl2.descriptors.nested import Value, BoolValue
from .hashable import HashableObject
from .colors import ColorDescriptor, BLACK

from openpyxl2.compat import safe_string, basestring
from openpyxl2.xml.functions import Element, SubElement


class Font(HashableObject):
    """Font options used in styles."""

    spec = """18.8.22, p.3930"""

    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'


    name = Value(expected_type=basestring)
    charset = Value(allow_none=True, expected_type=int)
    family = MinMax(min=0, max=14, nested=True)
    sz = Value(expected_type=float)
    size = Alias("sz")
    b = BoolValue()
    bold = Alias("b")
    i = BoolValue()
    italic = Alias("i")
    strike = BoolValue()
    strikethrough = Alias("strike")
    outline = BoolValue()
    shadow = BoolValue()
    condense = BoolValue()
    extend = BoolValue()
    u = NoneSet(values=('single', 'double', 'singleAccounting',
                        'doubleAccounting'), nested=True
                )
    underline = Alias("u")
    vertAlign = NoneSet(values=('superscript', 'subscript', 'baseline'), nested=True)
    color = ColorDescriptor()
    scheme = NoneSet(values=("major", "minor"), nested=True)

    tagname = "font"

    __nested__ = ('name', 'charset', 'family', 'b', 'i', 'strike', 'outline',
                  'shadow', 'condense', 'extend', 'sz', 'u', 'vertAlign',
                  'scheme')

    __fields__ = ('name', 'charset', 'family', 'b', 'i', 'strike', 'outline',
                  'shadow', 'condense', 'extend', 'sz', 'u', 'vertAlign',
                  'scheme', 'color')

    @classmethod
    def _create_nested(cls, el, tag):
        if tag == "u":
            return el.get("val", "single")
        return super(Font, cls)._create_nested(el, tag)

    def to_tree(self, tagname=None):
        el = Element(self.tagname)
        attrs = list(self.__nested__)
        attrs.insert(10, 'color')
        for attr in attrs:
            value = getattr(self, attr)
            if value:
                if attr == 'color':
                    color = value.to_tree()
                    el.append(color)
                else:
                    SubElement(el, attr, val=safe_string(value))
        return el

    def __init__(self, name='Calibri', sz=11, b=False, i=False, charset=None,
                 u=None, strike=False, color=BLACK, scheme=None, family=2, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=False, shadow=False, condense=False,
                 extend=False):
        self.name = name
        self.family = family
        if size is not None:
            sz = size
        self.sz = sz
        if bold is not None:
            b = bold
        self.b = b
        if italic is not None:
            i = italic
        self.i = i
        if underline is not None:
            u = underline
        self.u = u
        if strikethrough is not None:
            strike = strikethrough
        self.strike = strike
        self.color = color
        self.vertAlign = vertAlign
        self.charset = charset
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.scheme = scheme


from . colors import Color

DEFAULT_FONT = Font(color=Color(theme=1), scheme="minor")
