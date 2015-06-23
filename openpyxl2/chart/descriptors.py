from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.compat import basestring

from openpyxl2.descriptors.nested import (
    NestedMinMax
    )

from openpyxl2.descriptors import Typed

from .data_source import NumFmt

"""
Utility descriptors for the chart module.
For convenience but also clarity.
"""

class NestedGapAmount(NestedMinMax):

    allow_none = True
    min = 0
    max = 500


class NestedOverlap(NestedMinMax):

    allow_none = True
    min = -100
    max = 100


class NumberFormatDescriptor(Typed):
    """
    Allow direct assignment of format code
    """

    expected_type = NumFmt
    allow_none = True

    def __set__(self, instance, value):
        if isinstance(value, basestring):
            value = NumFmt(value)
        super(NumberFormatDescriptor, self).__set__(instance, value)