from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.compat import safe_string

from openpyxl2.descriptors import (
    Strict,
    Typed,
    Integer,
    Bool,
    String,
    Sequence,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.serialisable import Serialisable

from .fills import PatternFill, Fill
from . fonts import Font, DEFAULT_FONT
from . borders import Border
from . alignment import Alignment
from . numbers import NumberFormatDescriptor
from . protection import Protection


class NamedStyle(Strict):

    """
    Named and editable styles
    """

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill)
    border = Typed(expected_type=Border)
    alignment = Typed(expected_type=Alignment)
    number_format = NumberFormatDescriptor()
    protection = Typed(expected_type=Protection)
    builtinId = Integer(allow_none=True)
    hidden = Bool(allow_none=True)

    __fields__ = ("name", "font", "fill", "border", "number_format", "alignment", "protection")

    def __init__(self,
                 name="Normal",
                 font=Font(),
                 fill=PatternFill(),
                 border=Border(),
                 alignment=Alignment(),
                 number_format=None,
                 protection=Protection(),
                 builtinId=0,
                 hidden=False,
                 ):
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
        self.builtinId = builtinId
        self.hidden = hidden


    def _make_key(self):
        """Use a tuple of fields as the basis for a key"""
        self._key = hash(tuple(getattr(self, x) for x in self.__fields__))

    def __hash__(self):
        if not hasattr(self, '_key'):
            self._make_key()
        return self._key


    def __eq__(self, other):
        if isinstance(other, self.__class__):
            if not hasattr(self, '_key'):
                self._make_key()
            if not hasattr(other, '_key'):
                other._make_key()
            return self._key == other._key


    def __ne__(self, other):
        return not self == other

    def __repr__(self):
        pieces = []
        for k in self.__fields__:
            value = getattr(self, k)
            pieces.append('%s=%s' % (k, repr(value)))
        return '%s(%s)' % (self.__class__.__name__, ', '.join(pieces))


    def __iter__(self):
        for key in ('name', 'builtinId', 'hidden', 'xfId'):
            value = getattr(self, key, None)
            if value is not None:
                yield key, safe_string(value)


class NamedCellStyle(Serialisable):

    """
    Pointer-based representation of named styles in XML
    xfId refers to the corresponding CellStyleXf
    """

    tagname = "cellStyle"

    name = String()
    xfId = Integer()
    builtinId = Integer(allow_none=True)
    iLevel = Integer(allow_none=True)
    hidden = Bool(allow_none=True)
    customBuiltin = Bool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()

    def __init__(self,
                 name=None,
                 xfId=None,
                 builtinId=None,
                 iLevel=None,
                 hidden=None,
                 customBuiltin=None,
                 extLst=None,
                ):
        self.name = name
        self.xfId = xfId
        self.builtinId = builtinId
        self.iLevel = iLevel
        self.hidden = hidden
        self.customBuiltin = customBuiltin


class NamedCellStyleList(Serialisable):

    tagname = "cellStyles"

    count = Integer(allow_none=True)
    cellStyle = Sequence(expected_type=NamedCellStyle)


    def __init__(self,
                 count=None,
                 cellStyle=(),
                ):
        self.cellStyle = cellStyle


    @property
    def count(self):
        return len(self.cellStyle)


    @property
    def styles(self):
        """
        Convert to NamedStyle objects and remove duplicates
        """
        styles = {}
        for ns in self.cellStyle:
            style = NamedStyle(name=ns.name,
                                hidden=ns.hidden
                                )
            style.xfId = ns.xfId
            styles[ns.name] = style
        return styles
