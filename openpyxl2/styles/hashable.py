from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from copy import copy

from openpyxl2.compat import deprecated
from openpyxl2.descriptors import Descriptor
from openpyxl2.descriptors.serialisable import Serialisable


class HashableObject(Serialisable):
    """Define how to hash property classes."""

    __fields__ = ()
    _key = None


    @deprecated("Use copy()")
    def copy(self, **kwargs):
        return copy(self)


    @property
    def key(self):
        """Use a tuple of fields as the basis for a key"""
        if self._key is None:
            fields = []
            for attr in self.__fields__:
                val = getattr(self, attr)
                if isinstance(val, list):
                    val = tuple(val)
                fields.append(val)
            self._key = hash(tuple(fields))
        return self._key

    def __hash__(self):
        return self.key

    def __eq__(self, other):
        if other.__class__ == self.__class__:
            return self.key == other.key
        return self.key == other

    def __ne__(self, other):
        return not self == other

    def __add__(self, other):
        vals = {}
        for attr in self.__fields__:
            vals[attr] = getattr(self, attr) or getattr(other, attr)
        return self.__class__(**vals)

    def __sub__(self, other):
        vals = {}
        if (self is other) or (self == other):
            return
        for attr in self.__fields__:
            if not getattr(other, attr) and getattr(self, attr):
                vals[attr] = getattr(self, attr)
        return self.__class__(**vals)
