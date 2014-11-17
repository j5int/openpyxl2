from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file


from openpyxl2.compat import safe_string
from openpyxl2.descriptors import Strict, Float, Typed, Bool, Integer, String, Set, MatchPattern
from openpyxl2.xml.functions import Element
from openpyxl2.xml.constants import SHEET_MAIN_NS, REL_NS

def untuple(value):
    if isinstance(value, tuple):
        return value[0]
    else:
        return value
    
class CTPageSetup(Strict):
    """ Worksheet page setup """

    tag = "{%s}pageSetup" % SHEET_MAIN_NS

    orientation = Set(values=(None, "default", "portrait", "landscape"))
    paperSize = Integer(allow_none=True)
    scale = Integer(allow_none=True)
    fitToHeight = Integer(allow_none=True)
    fitToWidth = Integer(allow_none=True)
    firstPageNumber = Integer(allow_none=True)
    useFirstPageNumber = Bool(allow_none=True)
    paperHeight = MatchPattern(pattern="[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)", allow_none=True)  # ST_PositiveUniversalMeasure
    paperWidth = MatchPattern(pattern="[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)", allow_none=True)  # ST_PositiveUniversalMeasure
    pageOrder = Set(values=(None, "downThenOver", "overThenDown"))
    usePrinterDefaults = Bool(allow_none=True)
    blackAndWhite = Bool(allow_none=True)
    draft = Bool(allow_none=True)
    cellComments = Set(values=(None, "none", "asDisplayed", "atEnd"))
    errors = Set(values=(None, "displayed", "blank", "dash", "NA"))
    horizontalDpi = Integer(allow_none=True)
    verticalDpi = Integer(allow_none=True)
    copies = Integer(allow_none=True)
    id = String(allow_none=True)

    def __init__(self, orientation=None,
                 paperSize=None,
                 scale=None,
                 fitToHeight=None,
                 fitToWidth=None,
                 firstPageNumber=None,
                 useFirstPageNumber=None,
                 paperHeight=None,
                 paperWidth=None,
                 pageOrder=None,
                 usePrinterDefaults=None,
                 blackAndWhite=None,
                 draft=None,
                 cellComments=None,
                 errors=None,
                 horizontalDpi=None,
                 verticalDpi=None,
                 copies=None,
                 id=None):
        self.orientation = untuple(orientation)
        self.paperSize = untuple(paperSize)
        self.scale = untuple(scale)
        self.fitToHeight = untuple(fitToHeight)
        self.fitToWidth = untuple(fitToWidth)
        self.firstPageNumber = untuple(firstPageNumber)
        self.useFirstPageNumber = untuple(useFirstPageNumber)
        self.paperHeight = untuple(paperHeight)
        self.paperWidth = untuple(paperWidth)
        self.pageOrder = untuple(pageOrder)
        self.usePrinterDefaults = untuple(usePrinterDefaults)
        self.blackAndWhite = untuple(blackAndWhite)
        self.draft = untuple(draft)
        self.cellComments = untuple(cellComments)
        self.errors = untuple(errors)
        self.horizontalDpi = untuple(horizontalDpi)
        self.verticalDpi = untuple(verticalDpi)
        self.copies = untuple(copies)
        self.id = untuple(id)

    def __iter__(self):
        for attr in ("orientation", "paperSize", "scale", "fitToHeight", "fitToWidth", "firstPageNumber", "useFirstPageNumber"
                     , "paperHeight", "paperWidth", "pageOrder", "usePrinterDefaults", "blackAndWhite", "draft", "cellComments", "errors"
                     , "horizontalDpi", "verticalDpi", "copies", "id"):
            value = getattr(self, attr)
            if value is not None:
                if attr == "id":
                    key = '{%s}id' % REL_NS   
                    yield key, safe_string(value)
                else:
                    yield attr, safe_string(value)

    def write_xml_element(self):

        el = Element(self.tag, dict(self))

        return el


class PrintOptions(Strict):
    """ Worksheet print options """

    tag = "{%s}printOptions" % SHEET_MAIN_NS
    horizontalCentered = Bool(allow_none=True)
    verticalCentered = Bool(allow_none=True)
    headings = Bool(allow_none=True)
    gridLines = Bool(allow_none=True)
    gridLinesSet = Bool(allow_none=True)

    def __init__(self, horizontalCentered=None,
                 verticalCentered=None,
                 headings=None,
                 gridLines=None,
                 gridLinesSet=None,
                 ):
        self.horizontalCentered = horizontalCentered
        self.verticalCentered = verticalCentered
        self.headings = headings
        self.gridLines = gridLines
        self.gridLinesSet = gridLinesSet

    def __iter__(self):
        for attr in ("horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet"):
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


    def write_xml_element(self):

        el = Element(self.tag, dict(self))

        return el


class PageSetup(object):
    """Information about page layout for this sheet"""

    setup = Typed(expected_type=CTPageSetup, allow_none=True)
    options = Typed(expected_type=PrintOptions, allow_none=True)
    orientation = None
    paperSize = None
    scale = None
    fitToHeight = None
    fitToWidth = None
    firstPageNumber = None
    useFirstPageNumber = None
    paperHeight = None
    paperWidth = None
    pageOrder = None
    usePrinterDefaults = None
    blackAndWhite = None
    draft = None
    cellComments = None
    errors = None
    horizontalDpi = None
    verticalDpi = None
    copies = None
    id = None
    horizontalCentered = None
    verticalCentered = None
    headings = None
    gridLines = None
    gridLinesSet = None

    @property
    def setup(self):
        setupGroup = CTPageSetup(orientation=self.orientation,
                                paperSize=self.paperSize,
                                scale=self.scale,
                                fitToHeight=self.fitToHeight,
                                fitToWidth=self.fitToWidth,
                                firstPageNumber=self.firstPageNumber,
                                useFirstPageNumber=self.useFirstPageNumber,
                                paperHeight=self.paperHeight,
                                paperWidth=self.paperWidth,
                                pageOrder=self.pageOrder,
                                usePrinterDefaults=self.usePrinterDefaults,
                                blackAndWhite=self.blackAndWhite,
                                draft=self.draft,
                                cellComments=self.cellComments,
                                errors=self.errors,
                                horizontalDpi=self.horizontalDpi,
                                verticalDpi=self.verticalDpi,
                                copies=self.copies,
                                id=self.id
                                )

        return setupGroup

    @setup.setter
    def setup(self, values):
        self.orientation = values.orientation,
        self.paperSize = values.paperSize,
        self.scale = values.scale,
        self.fitToHeight = values.fitToHeight,
        self.fitToWidth = values.fitToWidth,
        self.firstPageNumber = values.firstPageNumber,
        self.useFirstPageNumber = values.useFirstPageNumber,
        self.paperHeight = values.paperHeight,
        self.paperWidth = values.paperWidth,
        self.pageOrder = values.pageOrder,
        self.usePrinterDefaults = values.usePrinterDefaults,
        self.blackAndWhite = values.blackAndWhite,
        self.draft = values.draft,
        self.cellComments = values.cellComments,
        self.errors = values.errors,
        self.horizontalDpi = values.horizontalDpi,
        self.verticalDpi = values.verticalDpi,
        self.copies = values.copies
        self.id = values.id

    @property
    def options(self):
        optionsGroup = PrintOptions(horizontalCentered=self.horizontalCentered,
                 verticalCentered=self.verticalCentered,
                 headings=self.headings,
                 gridLines=self.gridLines,
                 gridLinesSet=self.gridLinesSet)

        return optionsGroup

    @options.setter
    def options(self, values):
        self.horizontalCentered = values.horizontalCentered,
        self.verticalCentered = values.verticalCentered,
        self.headings = values.headings,
        self.gridLines = values.gridLines,
        self.gridLinesSet = values.gridLinesSet


class PageMargins(Strict):
    """
    Information about page margins for view/print layouts.
    Standard values (in inches)
    left, right = 0.75
    top, bottom = 1
    header, footer = 0.5
    """

    left = Float()
    right = Float()
    top = Float()
    bottom = Float()
    header = Float()
    footer = Float()

    def __init__(self, left=0.75, right=0.75, top=1, bottom=1, header=0.5, footer=0.5):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.header = header
        self.footer = footer

    def __iter__(self):
        for key in ("left", "right", "top", "bottom", "header", "footer"):
            value = getattr(self, key)
            yield key, safe_string(value)
