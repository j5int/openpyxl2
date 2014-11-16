from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from abc import abstractmethod, abstractproperty
from openpyxl2.compat.abc import ABC
from openpyxl2.utils.indexed_list import IndexedList

from .numbers import BUILTIN_FORMATS, BUILTIN_FORMATS_REVERSE


class StyleProxy(object):
    """
    Proxy formatting objects so that they cannot be altered
    """

    __slots__ = ('__target')

    def __init__(self, target):
        if not hasattr(target, 'copy'):
            raise TypeError("Proxied objects must have a copy method.")
        self.__target = target


    def __repr__(self):
        return repr(self.__target)


    def __getattr__(self, attr):
        return getattr(self.__target, attr)


    def __setattr__(self, attr, value):
        if attr != "_StyleProxy__target":
            raise AttributeError("Style objects are immutable and cannot be changed."
                                 "Reassign the style with a copy")
        super(StyleProxy, self).__setattr__(attr, value)


    def copy(self, **kw):
        """Return a copy of the proxied object. Keyword args will be passed through"""
        return self.__target.copy(**kw)


    def __eq__(self, other):
        return self.__target == other


    def __ne__(self, other):
        return not self == other


class StyledObject(ABC):
    """
    Mixin Class for stylable objects implementing proxy and lookup functions
    """

    @abstractmethod
    def __init__(self):
        self._font_id = None
        self._fill_id = None
        self._border_id = None
        self._alignment_id = None
        self._protection_id = None
        self._number_format_id = 0
        self._style_id = None

    @abstractproperty
    def _fonts(self):
        return IndexedList()

    @property
    def font(self):
        fo = self._fonts.get(self._font_id)
        if fo is not None:
            return StyleProxy(fo)

    @font.setter
    def font(self, value):
        self._font_id = self._fonts.add(value)


    @abstractproperty
    def _fills(self):
        return IndexedList()

    @property
    def fill(self):
        fo = self._fonts.get(self._fill_id)
        if fo is not None:
            return StyleProxy(fo)

    @fill.setter
    def font(self, value):
        self._fill_id = self._fills.add(value)


    @abstractproperty
    def _borders(self):
        return IndexedList()

    @property
    def border(self):
        fo = self._fonts.get(self._border_id)
        if fo is not None:
            return StyleProxy(fo)

    @border.setter
    def border(self, value):
        self._border_id = self._borders.add(value)


    @abstractproperty
    def _alignments(self):
        return IndexedList()

    @property
    def alignment(self):
        fo = self._fonts.get(self._alignment_id)
        if fo is not None:
            return StyleProxy(fo)

    @alignment.setter
    def alignment(self, value):
        self._alignment_id = self._alignments.add(value)


    @abstractproperty
    def _protections(self):
        return IndexedList()

    @property
    def protection(self):
        fo = self._fonts.get(self._protection_id)
        if fo is not None:
            return StyleProxy(fo)

    @protection.setter
    def protection(self, value):
        self._protection_id = self._protections.add(value)


    @abstractproperty
    def _styles(self):
        return IndexedList()

    @property
    def style(self):
        fo = self._styles.get(self._style_id)
        if fo is not None:
            return StyleProxy(fo)

    @style.setter
    def style(self, value):
        self._style_id = self._styles.add(value)


    @abstractproperty
    def _number_formats(self):
        return IndexedList()

    @property
    def number_format(self):
        if self._number_format_id < 164:
            return BUILTIN_FORMATS.get(self._number_format_id, "General")
        return self._number_formats[self._number_format_id - 164]

    @number_format.setter
    def number_format(self, value):
        _id = BUILTIN_FORMATS_REVERSE.get(value)
        if _id is None:
            _id = self._number_formats.add(value) + 164
        self._number_format_id = _id
