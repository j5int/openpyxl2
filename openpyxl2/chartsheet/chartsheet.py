from __future__ import absolute_import

from openpyxl2.chartsheet.relation import Drawing, DrawingHF, SheetBackgroundPicture
from openpyxl2.chartsheet.chartsheetpr import ChartsheetPr
from openpyxl2.chartsheet.chartsheetprotection import ChartsheetProtection
from openpyxl2.chartsheet.chartsheetview import ChartsheetViews
from openpyxl2.chartsheet.customchartsheetview import CustomChartsheetViews
from openpyxl2.chartsheet.webpublishitem import WebPublishItems

from openpyxl2.descriptors import Typed
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.worksheet.page import (PageMargins, PrintPageSetup)
from openpyxl2.worksheet.header_footer import HeaderFooter


class Chartsheet(Serialisable):
    tagname = "chartsheet"

    sheetPr = Typed(expected_type=ChartsheetPr, allow_none=True)
    sheetViews = Typed(expected_type=ChartsheetViews, )
    sheetProtection = Typed(expected_type=ChartsheetProtection, allow_none=True)
    customSheetViews = Typed(expected_type=CustomChartsheetViews, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PrintPageSetup, allow_none=True)
    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    drawing = Typed(expected_type=Drawing, )
    drawingHF = Typed(expected_type=DrawingHF, allow_none=True)
    picture = Typed(expected_type=SheetBackgroundPicture, allow_none=True)
    webPublishItems = Typed(expected_type=WebPublishItems, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = (
        'sheetPr', 'sheetViews', 'sheetProtection', 'customSheetViews', 'pageMargins', 'pageSetup', 'headerFooter',
        'drawing', 'drawingHF', 'picture', 'webPublishItems')

    def __init__(self,
                 sheetPr=None,
                 sheetViews=None,
                 sheetProtection=None,
                 customSheetViews=None,
                 pageMargins=None,
                 pageSetup=None,
                 headerFooter=None,
                 drawing=None,
                 drawingHF=None,
                 picture=None,
                 webPublishItems=None,
                 extLst=None,
                 ):
        self.sheetPr = sheetPr
        self.sheetViews = sheetViews
        self.sheetProtection = sheetProtection
        self.customSheetViews = customSheetViews
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.headerFooter = headerFooter
        self.drawing = drawing
        self.drawingHF = drawingHF
        self.picture = picture
        self.webPublishItems = webPublishItems
