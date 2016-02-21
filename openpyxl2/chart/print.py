from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Float,
    Typed,
    Alias,
)

from openpyxl2.worksheet.page import PrintPageSetup
from openpyxl2.worksheet.header_footer import HeaderFooter


class PageMargins(Serialisable):
    """
    Identical to openpyxl.worksheet.page.Pagemargins but element names are different :-/
    """
    tagname = "pageMargins"

    l = Float()
    left = Alias('l')
    r = Float()
    right = Alias('r')
    t = Float()
    top = Alias('t')
    b = Float()
    bottom = Alias('b')
    header = Float()
    footer = Float()

    def __init__(self, l=0.75, r=0.75, t=1, b=1, header=0.5, footer=0.5):
        self.l = l
        self.r = r
        self.t = t
        self.b = b
        self.header = header
        self.footer = footer


class PrintSettings(Serialisable):

    tagname = "printSettings"

    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)

    __elements__ = ("headerFooter", "pageMargins", "pageMargins")

    def __init__(self,
                 headerFooter=None,
                 pageMargins=None,
                 pageSetup=None,
                ):
        self.headerFooter = headerFooter
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
