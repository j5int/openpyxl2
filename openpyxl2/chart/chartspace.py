from __future__ import absolute_import

"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax,
    Integer,
    NoneSet,
    String,
    Alias,
)
from openpyxl2.descriptors.excel import (
    Percentage,
    ExtensionList
    )

from openpyxl2.descriptors.nested import (
    NestedBool,
    NestedNoneSet,
    NestedInteger,
    NestedString,
)

from .colors import ColorMapping
from .text import Tx, TextBody
from .layout import Layout
from .shapes import ShapeProperties
from .legend import Legend
from .marker import PictureOptions, Marker
from .label import DataLabel

from .area_chart import AreaChart, AreaChart3D
from .bar_chart import BarChart, BarChart3D
from .bubble_chart import BubbleChart
from .line_chart import LineChart, LineChart3D
from .pie_chart import PieChart, PieChart3D, ProjectedPieChart, DoughnutChart
from .radar_chart import RadarChart
from .scatter_chart import ScatterChart
from .stock_chart import StockChart
from .surface_chart import SurfaceChart, SurfaceChart3D

from .axis import ValAx, CatAx, SerAx, DateAx
from .title import Title

from openpyxl2.worksheet.page import PageMargins, PageSetup
from openpyxl2.worksheet.header_footer import HeaderFooter


class Surface(Serialisable):

    tagname = "surface"

    thickness = NestedInteger(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    shapeProperties = Alias('spPr')
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('thickness', 'spPr', 'pictureOptions',)

    def __init__(self,
                 thickness=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.thickness = thickness
        self.spPr = spPr
        self.pictureOptions = pictureOptions


class View3D(Serialisable):

    rotX = Integer(allow_none=True, nested=True)
    hPercent = Percentage(allow_none=True, nested=True)
    rotY = Integer(allow_none=True, nested=True)
    depthPercent = Percentage(allow_none=True, nested=True)
    rAngAx = Bool(nested=True, allow_none=True)
    perspective = Integer(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('rotX', 'hPercent', 'rotY', 'depthPercent', 'rAngAx', 'perspective', 'extLst')

    def __init__(self,
                 rotX=None,
                 hPercent=None,
                 rotY=None,
                 depthPercent=None,
                 rAngAx=None,
                 perspective=None,
                 extLst=None,
                ):
        self.rotX = rotX
        self.hPercent = hPercent
        self.rotY = rotY
        self.depthPercent = depthPercent
        self.rAngAx = rAngAx
        self.perspective = perspective
        self.extLst = extLst


class PivotFmt(Serialisable):

    idx = Integer(nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    dLbl = Typed(expected_type=DataLabel, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'spPr', 'txPr', 'marker', 'dLbl', 'extLst')

    def __init__(self,
                 idx=None,
                 spPr=None,
                 txPr=None,
                 marker=None,
                 dLbl=None,
                 extLst=None,
                ):
        self.idx = idx
        self.spPr = spPr
        self.txPr = txPr
        self.marker = marker
        self.dLbl = dLbl


class PivotFmts(Serialisable):

    pivotFmt = Typed(expected_type=PivotFmt, allow_none=True)

    __elements__ = ('pivotFmt',)

    def __init__(self,
                 pivotFmt=None,
                ):
        self.pivotFmt = pivotFmt


class DataTable(Serialisable):

    tagname = "dTable"

    showHorzBorder = NestedBool(allow_none=True)
    showVertBorder = NestedBool(allow_none=True)
    showOutline = NestedBool(allow_none=True)
    showKeys = NestedBool(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    ShapeProperties = Alias('spPr')
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('showHorzBorder', 'showVertBorder', 'showOutline',
                    'showKeys', 'spPr', 'txPr')

    def __init__(self,
                 showHorzBorder=None,
                 showVertBorder=None,
                 showOutline=None,
                 showKeys=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.showHorzBorder = showHorzBorder
        self.showVertBorder = showVertBorder
        self.showOutline = showOutline
        self.showKeys = showKeys
        self.spPr = spPr
        self.txPr = txPr


class PlotArea(Serialisable):

    tagname = "plotArea"

    layout = Typed(expected_type=Layout, allow_none=True)
    dTable = Typed(expected_type=DataTable, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    # at least one chart
    areaChart = Typed(expected_type=AreaChart, allow_none=True)
    area3DChart = Typed(expected_type=AreaChart3D, allow_none=True)
    lineChart = Typed(expected_type=LineChart, allow_none=True)
    line3DChart = Typed(expected_type=LineChart3D, allow_none=True)
    stockChart = Typed(expected_type=StockChart, allow_none=True)
    radarChart = Typed(expected_type=RadarChart, allow_none=True)
    scatterChart = Typed(expected_type=ScatterChart, allow_none=True)
    pieChart = Typed(expected_type=PieChart, allow_none=True)
    pie3DChart = Typed(expected_type=PieChart3D, allow_none=True)
    doughnutChart = Typed(expected_type=DoughnutChart, allow_none=True)
    barChart = Typed(expected_type=BarChart, allow_none=True)
    bar3DChart = Typed(expected_type=BarChart3D, allow_none=True)
    ofPieChart = Typed(expected_type=ProjectedPieChart, allow_none=True)
    surfaceChart = Typed(expected_type=SurfaceChart, allow_none=True)
    surface3DChart = Typed(expected_type=SurfaceChart3D, allow_none=True)
    bubbleChart = Typed(expected_type=BubbleChart, allow_none=True)

    # maybe axes
    valAx = Typed(expected_type=ValAx, allow_none=True)
    catAx = Typed(expected_type=CatAx, allow_none=True)
    dateAx = Typed(expected_type=DateAx, allow_none=True)
    serAx = Typed(expected_type=SerAx, allow_none=True)

    __elements__ = ('layout', 'areaChart', 'area3DChart', 'lineChart',
                    'line3DChart', 'stockChart', 'radarChart', 'scatterChart', 'pieChart',
                    'pie3DChart', 'doughnutChart', 'barChart', 'bar3DChart', 'ofPieChart',
                    'surfaceChart', 'surface3DChart', 'bubbleChart', 'valAx', 'catAx', 'dateAx', 'serAx',
                    'dTable', 'spPr')

    def __init__(self,
                 layout=None,
                 dTable=None,
                 spPr=None,
                 areaChart=None,
                 area3DChart=None,
                 lineChart=None,
                 line3DChart=None,
                 stockChart=None,
                 radarChart=None,
                 scatterChart=None,
                 pieChart=None,
                 pie3DChart=None,
                 doughnutChart=None,
                 barChart=None,
                 bar3DChart=None,
                 ofPieChart=None,
                 surfaceChart=None,
                 surface3DChart=None,
                 bubbleChart=None,
                 valAx=None,
                 catAx=None,
                 serAx=None,
                 dateAx=None,
                 extLst=None,
                ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr
        self.areaChart = areaChart
        self.area3DChart = area3DChart
        self.lineChart = lineChart
        self.line3DChart = line3DChart
        self.stockChart = stockChart
        self.radarChart = radarChart
        self.scatterChart = scatterChart
        self.pieChart = pieChart
        self.pie3DChart = pie3DChart
        self.doughnutChart = doughnutChart
        self.barChart = barChart
        self.bar3DChart = bar3DChart
        self.ofPieChart = ofPieChart
        self.surfaceChart = surfaceChart
        self.surface3DChart = surface3DChart
        self.bubbleChart = bubbleChart
        self.valAx = valAx
        self.catAx = catAx
        self.dateAx = dateAx
        self.serAx = serAx


class ChartContainer(Serialisable):

    tagname = "chart"

    title = Typed(expected_type=Title, allow_none=True)
    autoTitleDeleted = NestedBool(allow_none=True)
    pivotFmts = Typed(expected_type=PivotFmts, allow_none=True)
    view3D = Typed(expected_type=View3D, allow_none=True)
    floor = Typed(expected_type=Surface, allow_none=True)
    sideWall = Typed(expected_type=Surface, allow_none=True)
    backWall = Typed(expected_type=Surface, allow_none=True)
    plotArea = Typed(expected_type=PlotArea, )
    legend = Typed(expected_type=Legend, allow_none=True)
    plotVisOnly = NestedBool(allow_none=True)
    dispBlanksAs = NestedNoneSet(values=(['span', 'gap', 'zero']))
    showDLblsOverMax = NestedBool(allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('title', 'autoTitleDeleted', 'pivotFmts', 'view3D',
                    'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly',
                    'dispBlanksAs', 'showDLblsOverMax')

    def __init__(self,
                 title=None,
                 autoTitleDeleted=None,
                 pivotFmts=None,
                 view3D=None,
                 floor=None,
                 sideWall=None,
                 backWall=None,
                 plotArea=None,
                 legend=None,
                 plotVisOnly=None,
                 dispBlanksAs=None,
                 showDLblsOverMax=None,
                 extLst=None,
                ):
        self.title = title
        self.autoTitleDeleted = autoTitleDeleted
        self.pivotFmts = pivotFmts
        self.view3D = view3D
        self.floor = floor
        self.sideWall = sideWall
        self.backWall = backWall
        if plotArea is None:
            plotArea = PlotArea()
        self.plotArea = plotArea
        self.legend = legend
        self.plotVisOnly = plotVisOnly
        self.dispBlanksAs = dispBlanksAs
        self.showDLblsOverMax = showDLblsOverMax


class Protection(Serialisable):

    chartObject = NestedBool(llow_none=True)
    data = NestedBool(allow_none=True)
    formatting = NestedBool(allow_none=True)
    selection = NestedBool(allow_none=True)
    userInterface = NestedBool(allow_none=True)

    def __init__(self,
                 chartObject=None,
                 data=None,
                 formatting=None,
                 selection=None,
                 userInterface=None,
                ):
        self.chartObject = chartObject
        self.data = data
        self.formatting = formatting
        self.selection = selection
        self.userInterface = userInterface


class PivotSource(Serialisable):

    name = String()
    fmtId = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 name=None,
                 fmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fmtId = fmtId


class ExternalData(Serialisable):

    autoUpdate = NestedBool(allow_none=True)

    def __init__(self,
                 autoUpdate=None,
                ):
        self.autoUpdate = autoUpdate


class RelId(Serialisable):

    pass # todo


class PrintSettings(Serialisable):

    headerFooter = Typed(expected_type=HeaderFooter, allow_none=True)
    pageMargins = Typed(expected_type=PageMargins, allow_none=True)
    pageSetup = Typed(expected_type=PageSetup, allow_none=True)
    legacyDrawingHF = Typed(expected_type=RelId, allow_none=True)

    def __init__(self,
                 headerFooter=None,
                 pageMargins=None,
                 pageSetup=None,
                 legacyDrawingHF=None,
                ):
        self.headerFooter = headerFooter
        self.pageMargins = pageMargins
        self.pageSetup = pageSetup
        self.legacyDrawingHF = legacyDrawingHF


class ChartSpace(Serialisable):

    tagname = "chartBase"

    date1904 = NestedBool(allow_none=True)
    lang = NestedString(allow_none=True)
    roundedCorners = NestedBool(allow_none=True)
    style = NestedInteger(allow_none=True)
    clrMapOvr = Typed(expected_type=ColorMapping, allow_none=True)
    pivotSource = Typed(expected_type=PivotSource, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)
    chart = Typed(expected_type=ChartContainer)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    externalData = Typed(expected_type=ExternalData, allow_none=True)
    printSettings = Typed(expected_type=PrintSettings, allow_none=True)
    userShapes = Typed(expected_type=RelId, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('date1904', 'lang', 'roundedCorners', 'style',
                    'clrMapOvr', 'pivotSource', 'protection', 'chart', 'spPr', 'txPr',
                    'externalData', 'printSettings', 'userShapes')

    def __init__(self,
                 date1904=None,
                 lang=None,
                 roundedCorners=None,
                 style=None,
                 clrMapOvr=None,
                 pivotSource=None,
                 protection=None,
                 chart=None,
                 spPr=None,
                 txPr=None,
                 externalData=None,
                 printSettings=None,
                 userShapes=None,
                 extLst=None,
                ):
        self.date1904 = date1904
        self.lang = lang
        self.roundedCorners = roundedCorners
        self.style = style
        self.clrMapOvr = clrMapOvr
        self.pivotSource = pivotSource
        self.protection = protection
        self.chart = chart
        self.spPr = spPr
        self.txPr = txPr
        self.externalData = externalData
        self.printSettings = printSettings
        self.userShapes = userShapes
        self.extLst = extLst
