#Autogenerated schema
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Integer,
    NoneSet,
    Set,
    Float,
    Bool,
    DateTime,
    String,
    Alias,
    Bool,
)

from openpyxl2.descriptors.excel import ExtensionList

from openpyxl2.worksheet.filters import (
    AutoFilter,
    CellRange,
    ColorFilter,
    CustomFilter,
    CustomFilters,
    DateGroupItem,
    DynamicFilter,
    FilterColumn,
    Filters,
    IconFilter,
    SortCondition,
    SortState,
    Top10,
)


class HierarchyUsage(Serialisable):

    hierarchyUsage = Integer()

    def __init__(self,
                 hierarchyUsage=None,
                ):
        self.hierarchyUsage = hierarchyUsage


class ColHierarchiesUsage(Serialisable):

    count = Integer()
    colHierarchyUsage = Typed(expected_type=HierarchyUsage, )

    __elements__ = ('colHierarchyUsage',)

    def __init__(self,
                 count=None,
                 colHierarchyUsage=None,
                ):
        self.count = count
        self.colHierarchyUsage = colHierarchyUsage


class RowHierarchiesUsage(Serialisable):

    count = Integer()
    rowHierarchyUsage = Typed(expected_type=HierarchyUsage, )

    __elements__ = ('rowHierarchyUsage',)

    def __init__(self,
                 count=None,
                 rowHierarchyUsage=None,
                ):
        self.count = count
        self.rowHierarchyUsage = rowHierarchyUsage


class PivotFilter(Serialisable):

    fld = Integer()
    mpFld = Integer(allow_none=True)
    type = Set(values=(['unknown', 'count', 'percent', 'sum', 'captionEqual', 'captionNotEqual', 'captionBeginsWith', 'captionNotBeginsWith', 'captionEndsWith', 'captionNotEndsWith', 'captionContains', 'captionNotContains', 'captionGreaterThan', 'captionGreaterThanOrEqual', 'captionLessThan', 'captionLessThanOrEqual', 'captionBetween', 'captionNotBetween', 'valueEqual', 'valueNotEqual', 'valueGreaterThan', 'valueGreaterThanOrEqual', 'valueLessThan', 'valueLessThanOrEqual', 'valueBetween', 'valueNotBetween', 'dateEqual', 'dateNotEqual', 'dateOlderThan', 'dateOlderThanOrEqual', 'dateNewerThan', 'dateNewerThanOrEqual', 'dateBetween', 'dateNotBetween', 'tomorrow', 'today', 'yesterday', 'nextWeek', 'thisWeek', 'lastWeek', 'nextMonth', 'thisMonth', 'lastMonth', 'nextQuarter', 'thisQuarter', 'lastQuarter', 'nextYear', 'thisYear', 'lastYear', 'yearToDate', 'Q1', 'Q2', 'Q3', 'Q4', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'M10', 'M11', 'M12']))
    evalOrder = Integer(allow_none=True)
    id = Integer()
    iMeasureHier = Integer(allow_none=True)
    iMeasureFld = Integer(allow_none=True)
    name = String(allow_none=True)
    description = String()
    stringValue1 = String()
    stringValue2 = String()
    autoFilter = Typed(expected_type=AutoFilter, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter', 'extLst')

    def __init__(self,
                 fld=None,
                 mpFld=None,
                 type=None,
                 evalOrder=None,
                 id=None,
                 iMeasureHier=None,
                 iMeasureFld=None,
                 name=None,
                 description=None,
                 stringValue1=None,
                 stringValue2=None,
                 autoFilter=None,
                 extLst=None,
                ):
        self.fld = fld
        self.mpFld = mpFld
        self.type = type
        self.evalOrder = evalOrder
        self.id = id
        self.iMeasureHier = iMeasureHier
        self.iMeasureFld = iMeasureFld
        self.name = name
        self.description = description
        self.stringValue1 = stringValue1
        self.stringValue2 = stringValue2
        self.autoFilter = autoFilter
        self.extLst = extLst


class PivotFilters(Serialisable):

    count = Integer()
    filter = Typed(expected_type=PivotFilter, allow_none=True)

    __elements__ = ('filter',)

    def __init__(self,
                 count=None,
                 filter=None,
                ):
        self.count = count
        self.filter = filter


class PivotTableStyle(Serialisable):

    name = String()
    showRowHeaders = Bool()
    showColHeaders = Bool()
    showRowStripes = Bool()
    showColStripes = Bool()
    showLastColumn = Bool(allow_none=True)

    def __init__(self,
                 name=None,
                 showRowHeaders=None,
                 showColHeaders=None,
                 showRowStripes=None,
                 showColStripes=None,
                 showLastColumn=None,
                ):
        self.name = name
        self.showRowHeaders = showRowHeaders
        self.showColHeaders = showColHeaders
        self.showRowStripes = showRowStripes
        self.showColStripes = showColStripes
        self.showLastColumn = showLastColumn


class Member(Serialisable):

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class Members(Serialisable):

    count = Integer()
    level = Integer(allow_none=True)
    member = Typed(expected_type=Member, )

    __elements__ = ('member',)

    def __init__(self,
                 count=None,
                 level=None,
                 member=None,
                ):
        self.count = count
        self.level = level
        self.member = member


class MemberProperty(Serialisable):

    name = String(allow_none=True)
    showCell = Bool(allow_none=True)
    showTip = Bool(allow_none=True)
    showAsCaption = Bool(allow_none=True)
    nameLen = Integer(allow_none=True)
    pPos = Integer(allow_none=True)
    pLen = Integer(allow_none=True)
    level = Integer(allow_none=True)
    field = Integer()

    def __init__(self,
                 name=None,
                 showCell=None,
                 showTip=None,
                 showAsCaption=None,
                 nameLen=None,
                 pPos=None,
                 pLen=None,
                 level=None,
                 field=None,
                ):
        self.name = name
        self.showCell = showCell
        self.showTip = showTip
        self.showAsCaption = showAsCaption
        self.nameLen = nameLen
        self.pPos = pPos
        self.pLen = pLen
        self.level = level
        self.field = field


class MemberProperties(Serialisable):

    count = Integer()
    mp = Typed(expected_type=MemberProperty, )

    __elements__ = ('mp',)

    def __init__(self,
                 count=None,
                 mp=None,
                ):
        self.count = count
        self.mp = mp


class PivotHierarchy(Serialisable):

    outline = Bool()
    multipleItemSelectionAllowed = Bool()
    subtotalTop = Bool()
    showInFieldList = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    includeNewItemsInFilter = Bool()
    caption = String(allow_none=True)
    mps = Typed(expected_type=MemberProperties, allow_none=True)
    members = Typed(expected_type=Members, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('mps', 'members', 'extLst')

    def __init__(self,
                 outline=None,
                 multipleItemSelectionAllowed=None,
                 subtotalTop=None,
                 showInFieldList=None,
                 dragToRow=None,
                 dragToCol=None,
                 dragToPage=None,
                 dragToData=None,
                 dragOff=None,
                 includeNewItemsInFilter=None,
                 caption=None,
                 mps=None,
                 members=None,
                 extLst=None,
                ):
        self.outline = outline
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.subtotalTop = subtotalTop
        self.showInFieldList = showInFieldList
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.caption = caption
        self.mps = mps
        self.members = members
        self.extLst = extLst


class PivotHierarchies(Serialisable):

    count = Integer()
    pivotHierarchy = Typed(expected_type=PivotHierarchy, )

    __elements__ = ('pivotHierarchy',)

    def __init__(self,
                 count=None,
                 pivotHierarchy=None,
                ):
        self.count = count
        self.pivotHierarchy = pivotHierarchy


class ChartFormat(Serialisable):

    chart = Integer()
    format = Integer()
    series = Bool()
    #pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 chart=None,
                 format=None,
                 series=None,
                 pivotArea=None,
                ):
        self.chart = chart
        self.format = format
        self.series = series
        self.pivotArea = pivotArea


class ChartFormats(Serialisable):

    count = Integer()
    chartFormat = Typed(expected_type=ChartFormat, )

    __elements__ = ('chartFormat',)

    def __init__(self,
                 count=None,
                 chartFormat=None,
                ):
        self.count = count
        self.chartFormat = chartFormat


class PivotAreas(Serialisable):

    count = Integer()
    #pivotArea = Typed(expected_type=PivotArea, allow_none=True)

    __elements__ = ('pivotArea',)

    def __init__(self,
                 count=None,
                 pivotArea=None,
                ):
        self.count = count
        self.pivotArea = pivotArea


class ConditionalFormat(Serialisable):

    scope = Set(values=(['selection', 'data', 'field']))
    type = NoneSet(values=(['all', 'row', 'column']))
    priority = Integer()
    pivotAreas = Typed(expected_type=PivotAreas, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotAreas', 'extLst')

    def __init__(self,
                 scope=None,
                 type=None,
                 priority=None,
                 pivotAreas=None,
                 extLst=None,
                ):
        self.scope = scope
        self.type = type
        self.priority = priority
        self.pivotAreas = pivotAreas
        self.extLst = extLst


class ConditionalFormats(Serialisable):

    count = Integer()
    conditionalFormat = Typed(expected_type=ConditionalFormat, )

    __elements__ = ('conditionalFormat',)

    def __init__(self,
                 count=None,
                 conditionalFormat=None,
                ):
        self.count = count
        self.conditionalFormat = conditionalFormat


class Format(Serialisable):

    action = Set(values=(['blank', 'formatting', 'drill', 'formula']))
    dxfId = Integer()
    #pivotArea = Typed(expected_type=PivotArea, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotArea', 'extLst')

    def __init__(self,
                 action=None,
                 dxfId=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.action = action
        self.dxfId = dxfId
        self.pivotArea = pivotArea
        self.extLst = extLst


class Formats(Serialisable):

    count = Integer()
    format = Typed(expected_type=Format, )

    __elements__ = ('format',)

    def __init__(self,
                 count=None,
                 format=None,
                ):
        self.count = count
        self.format = format


class DataField(Serialisable):

    name = String(allow_none=True)
    fld = Integer()
    subtotal = Set(values=(['average', 'count', 'countNums', 'max', 'min', 'product', 'stdDev', 'stdDevp', 'sum', 'var', 'varp']))
    showDataAs = Set(values=(['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index']))
    baseField = Integer()
    baseItem = Integer()
    numFmtId = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 name=None,
                 fld=None,
                 subtotal=None,
                 showDataAs=None,
                 baseField=None,
                 baseItem=None,
                 numFmtId=None,
                 extLst=None,
                ):
        self.name = name
        self.fld = fld
        self.subtotal = subtotal
        self.showDataAs = showDataAs
        self.baseField = baseField
        self.baseItem = baseItem
        self.numFmtId = numFmtId
        self.extLst = extLst


class DataFields(Serialisable):

    count = Integer()
    dataField = Typed(expected_type=DataField, )

    __elements__ = ('dataField',)

    def __init__(self,
                 count=None,
                 dataField=None,
                ):
        self.count = count
        self.dataField = dataField


class PageField(Serialisable):

    fld = Integer()
    item = Integer(allow_none=True)
    hier = Integer()
    name = String()
    cap = String()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 fld=None,
                 item=None,
                 hier=None,
                 name=None,
                 cap=None,
                 extLst=None,
                ):
        self.fld = fld
        self.item = item
        self.hier = hier
        self.name = name
        self.cap = cap
        self.extLst = extLst


class PageFields(Serialisable):

    count = Integer()
    pageField = Typed(expected_type=PageField, )

    __elements__ = ('pageField',)

    def __init__(self,
                 count=None,
                 pageField=None,
                ):
        self.count = count
        self.pageField = pageField


class X(Serialisable):

    v = Integer()

    def __init__(self,
                 v=None,
                ):
        self.v = v


class I(Serialisable):

    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank']))
    r = Integer()
    i = Integer()
    x = Typed(expected_type=X, allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 t=None,
                 r=None,
                 i=None,
                 x=None,
                ):
        self.t = t
        self.r = r
        self.i = i
        self.x = x


class colItems(Serialisable):

    count = Integer()
    i = Typed(expected_type=I, )

    __elements__ = ('i',)

    def __init__(self,
                 count=None,
                 i=None,
                ):
        self.count = count
        self.i = i


class Field(Serialisable):

    x = Integer()

    def __init__(self,
                 x=None,
                ):
        self.x = x



class ColFields(Serialisable):

    count = Integer()
    field = Typed(expected_type=Field, )

    __elements__ = ('field',)

    def __init__(self,
                 count=None,
                 field=None,
                ):
        self.count = count
        self.field = field



class rowItems(Serialisable):

    count = Integer()
    i = Typed(expected_type=I, )

    __elements__ = ('i',)

    def __init__(self,
                 count=None,
                 i=None,
                ):
        self.count = count
        self.i = i


class RowFields(Serialisable):

    count = Integer()
    field = Typed(expected_type=Field, )

    __elements__ = ('field',)

    def __init__(self,
                 count=None,
                 field=None,
                ):
        self.count = count
        self.field = field


class Extension(Serialisable):

    uri = String()

    def __init__(self,
                 uri=None,
                ):
        self.uri = uri


class ExtensionList(Serialisable):

    # uses element group EG_ExtensionList
    ext = Typed(expected_type=Extension, allow_none=True)

    __elements__ = ('ext',)

    def __init__(self,
                 ext=None,
                ):
        self.ext = ext


class Index(Serialisable):

    v = Integer()

    def __init__(self,
                 v=None,
                ):
        self.v = v


class PivotAreaReference(Serialisable):

    field = Integer(allow_none=True)
    count = Integer()
    selected = Bool()
    byPosition = Bool()
    relative = Bool()
    defaultSubtotal = Bool()
    sumSubtotal = Bool()
    countASubtotal = Bool()
    avgSubtotal = Bool()
    maxSubtotal = Bool()
    minSubtotal = Bool()
    productSubtotal = Bool()
    countSubtotal = Bool()
    stdDevSubtotal = Bool()
    stdDevPSubtotal = Bool()
    varSubtotal = Bool()
    varPSubtotal = Bool()
    x = Typed(expected_type=Index, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('x', 'extLst')

    def __init__(self,
                 field=None,
                 count=None,
                 selected=None,
                 byPosition=None,
                 relative=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 x=None,
                 extLst=None,
                ):
        self.field = field
        self.count = count
        self.selected = selected
        self.byPosition = byPosition
        self.relative = relative
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.x = x
        self.extLst = extLst


class PivotAreaReferences(Serialisable):

    count = Integer()
    reference = Typed(expected_type=PivotAreaReference, )

    __elements__ = ('reference',)

    def __init__(self,
                 count=None,
                 reference=None,
                ):
        self.count = count
        self.reference = reference


class PivotArea(Serialisable):

    field = Integer(allow_none=True)
    type = NoneSet(values=(['normal', 'data', 'all', 'origin', 'button', 'topEnd', 'topRight']))
    dataOnly = Bool()
    labelOnly = Bool()
    grandRow = Bool()
    grandCol = Bool()
    cacheIndex = Bool()
    outline = Bool()
    offset = String()
    collapsedLevelsAreSubtotals = Bool()
    axis = Set(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    fieldPosition = Integer(allow_none=True)
    references = Typed(expected_type=PivotAreaReferences, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('references', 'extLst')

    def __init__(self,
                 field=None,
                 type=None,
                 dataOnly=None,
                 labelOnly=None,
                 grandRow=None,
                 grandCol=None,
                 cacheIndex=None,
                 outline=None,
                 offset=None,
                 collapsedLevelsAreSubtotals=None,
                 axis=None,
                 fieldPosition=None,
                 references=None,
                 extLst=None,
                ):
        self.field = field
        self.type = type
        self.dataOnly = dataOnly
        self.labelOnly = labelOnly
        self.grandRow = grandRow
        self.grandCol = grandCol
        self.cacheIndex = cacheIndex
        self.outline = outline
        self.offset = offset
        self.collapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals
        self.axis = axis
        self.fieldPosition = fieldPosition
        self.references = references
        self.extLst = extLst


class AutoSortScope(Serialisable):

    pivotArea = Typed(expected_type=PivotArea, )

    __elements__ = ('pivotArea',)

    def __init__(self,
                 pivotArea=None,
                ):
        self.pivotArea = pivotArea


class Item(Serialisable):

    n = String()
    t = Set(values=(['data', 'default', 'sum', 'countA', 'avg', 'max', 'min', 'product', 'count', 'stdDev', 'stdDevP', 'var', 'varP', 'grand', 'blank']))
    h = Bool()
    s = Bool()
    sd = Bool()
    f = Bool()
    m = Bool()
    c = Bool()
    x = Integer(allow_none=True)
    d = Bool()
    e = Bool()

    def __init__(self,
                 n=None,
                 t=None,
                 h=None,
                 s=None,
                 sd=None,
                 f=None,
                 m=None,
                 c=None,
                 x=None,
                 d=None,
                 e=None,
                ):
        self.n = n
        self.t = t
        self.h = h
        self.s = s
        self.sd = sd
        self.f = f
        self.m = m
        self.c = c
        self.x = x
        self.d = d
        self.e = e


class Items(Serialisable):

    count = Integer()
    item = Typed(expected_type=Item, )

    __elements__ = ('item',)

    def __init__(self,
                 count=None,
                 item=None,
                ):
        self.count = count
        self.item = item


class PivotField(Serialisable):

    name = String(allow_none=True)
    axis = NoneSet(values=(['axisRow', 'axisCol', 'axisPage', 'axisValues']))
    dataField = Bool()
    subtotalCaption = String(allow_none=True)
    showDropDowns = Bool()
    hiddenLevel = Bool()
    uniqueMemberProperty = String(allow_none=True)
    compact = Bool()
    allDrilled = Bool()
    numFmtId = Integer(allow_none=True)
    outline = Bool()
    subtotalTop = Bool()
    dragToRow = Bool()
    dragToCol = Bool()
    multipleItemSelectionAllowed = Bool()
    dragToPage = Bool()
    dragToData = Bool()
    dragOff = Bool()
    showAll = Bool()
    insertBlankRow = Bool()
    serverField = Bool()
    insertPageBreak = Bool()
    autoShow = Bool()
    topAutoShow = Bool()
    hideNewItems = Bool()
    measureFilter = Bool()
    includeNewItemsInFilter = Bool()
    itemPageCount = Integer(allow_none=True)
    sortType = Set(values=(['manual', 'ascending', 'descending']))
    dataSourceSort = Bool(allow_none=True)
    nonAutoSortDefault = Bool()
    rankBy = Integer(allow_none=True)
    defaultSubtotal = Bool(allow_none=True)
    sumSubtotal = Bool(allow_none=True)
    countASubtotal = Bool(allow_none=True)
    avgSubtotal = Bool(allow_none=True)
    maxSubtotal = Bool(allow_none=True)
    minSubtotal = Bool(allow_none=True)
    productSubtotal = Bool(allow_none=True)
    countSubtotal = Bool(allow_none=True)
    stdDevSubtotal = Bool(allow_none=True)
    stdDevPSubtotal = Bool(allow_none=True)
    varSubtotal = Bool(allow_none=True)
    varPSubtotal = Bool(allow_none=True)
    showPropCell = Bool(allow_none=True)
    showPropTip = Bool(allow_none=True)
    showPropAsCaption = Bool(allow_none=True)
    defaultAttributeDrillState = Bool(allow_none=True)
    items = Typed(expected_type=Items, allow_none=True)
    autoSortScope = Typed(expected_type=AutoSortScope, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('items', 'autoSortScope', 'extLst')

    def __init__(self,
                 name=None,
                 axis=None,
                 dataField=None,
                 subtotalCaption=None,
                 showDropDowns=None,
                 hiddenLevel=None,
                 uniqueMemberProperty=None,
                 compact=None,
                 allDrilled=None,
                 numFmtId=None,
                 outline=None,
                 subtotalTop=None,
                 dragToRow=None,
                 dragToCol=None,
                 multipleItemSelectionAllowed=None,
                 dragToPage=None,
                 dragToData=None,
                 dragOff=None,
                 showAll=None,
                 insertBlankRow=None,
                 serverField=None,
                 insertPageBreak=None,
                 autoShow=None,
                 topAutoShow=None,
                 hideNewItems=None,
                 measureFilter=None,
                 includeNewItemsInFilter=None,
                 itemPageCount=None,
                 sortType=None,
                 dataSourceSort=None,
                 nonAutoSortDefault=None,
                 rankBy=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None,
                 showPropCell=None,
                 showPropTip=None,
                 showPropAsCaption=None,
                 defaultAttributeDrillState=None,
                 items=None,
                 autoSortScope=None,
                 extLst=None,
                ):
        self.name = name
        self.axis = axis
        self.dataField = dataField
        self.subtotalCaption = subtotalCaption
        self.showDropDowns = showDropDowns
        self.hiddenLevel = hiddenLevel
        self.uniqueMemberProperty = uniqueMemberProperty
        self.compact = compact
        self.allDrilled = allDrilled
        self.numFmtId = numFmtId
        self.outline = outline
        self.subtotalTop = subtotalTop
        self.dragToRow = dragToRow
        self.dragToCol = dragToCol
        self.multipleItemSelectionAllowed = multipleItemSelectionAllowed
        self.dragToPage = dragToPage
        self.dragToData = dragToData
        self.dragOff = dragOff
        self.showAll = showAll
        self.insertBlankRow = insertBlankRow
        self.serverField = serverField
        self.insertPageBreak = insertPageBreak
        self.autoShow = autoShow
        self.topAutoShow = topAutoShow
        self.hideNewItems = hideNewItems
        self.measureFilter = measureFilter
        self.includeNewItemsInFilter = includeNewItemsInFilter
        self.itemPageCount = itemPageCount
        self.sortType = sortType
        self.dataSourceSort = dataSourceSort
        self.nonAutoSortDefault = nonAutoSortDefault
        self.rankBy = rankBy
        self.defaultSubtotal = defaultSubtotal
        self.sumSubtotal = sumSubtotal
        self.countASubtotal = countASubtotal
        self.avgSubtotal = avgSubtotal
        self.maxSubtotal = maxSubtotal
        self.minSubtotal = minSubtotal
        self.productSubtotal = productSubtotal
        self.countSubtotal = countSubtotal
        self.stdDevSubtotal = stdDevSubtotal
        self.stdDevPSubtotal = stdDevPSubtotal
        self.varSubtotal = varSubtotal
        self.varPSubtotal = varPSubtotal
        self.showPropCell = showPropCell
        self.showPropTip = showPropTip
        self.showPropAsCaption = showPropAsCaption
        self.defaultAttributeDrillState = defaultAttributeDrillState
        self.items = items
        self.autoSortScope = autoSortScope
        self.extLst = extLst


class PivotFields(Serialisable):

    count = Integer()
    pivotField = Typed(expected_type=PivotField, )

    __elements__ = ('pivotField',)

    def __init__(self,
                 count=None,
                 pivotField=None,
                ):
        self.count = count
        self.pivotField = pivotField


class Location(Serialisable):

    ref = String()
    firstHeaderRow = Integer()
    firstDataRow = Integer()
    firstDataCol = Integer()
    rowPageCount = Integer(allow_none=True)
    colPageCount = Integer(allow_none=True)

    def __init__(self,
                 ref=None,
                 firstHeaderRow=None,
                 firstDataRow=None,
                 firstDataCol=None,
                 rowPageCount=None,
                 colPageCount=None,
                ):
        self.ref = ref
        self.firstHeaderRow = firstHeaderRow
        self.firstDataRow = firstDataRow
        self.firstDataCol = firstDataCol
        self.rowPageCount = rowPageCount
        self.colPageCount = colPageCount


class PivotTableDefinition(Serialisable):

    #Using attribute groupAG_AutoFormat
    name = String(allow_none=True)
    cacheId = Integer()
    dataOnRows = Bool()
    dataPosition = Integer(allow_none=True)
    dataCaption = String()
    grandTotalCaption = String()
    errorCaption = String()
    showError = Bool()
    missingCaption = String()
    showMissing = Bool()
    pageStyle = String()
    pivotTableStyle = String()
    vacatedStyle = String()
    tag = String()
    updatedVersion = Integer()
    minRefreshableVersion = Integer()
    asteriskTotals = Bool()
    showItems = Bool()
    editData = Bool()
    disableFieldList = Bool()
    showCalcMbrs = Bool()
    visualTotals = Bool()
    showMultipleLabel = Bool()
    showDataDropDown = Bool()
    showDrill = Bool()
    printDrill = Bool()
    showMemberPropertyTips = Bool()
    showDataTips = Bool()
    enableWizard = Bool()
    enableDrill = Bool()
    enableFieldProperties = Bool()
    preserveFormatting = Bool()
    useAutoFormatting = Bool()
    pageWrap = Integer()
    pageOverThenDown = Bool()
    subtotalHiddenItems = Bool()
    rowGrandTotals = Bool()
    colGrandTotals = Bool()
    fieldPrintTitles = Bool()
    itemPrintTitles = Bool()
    mergeItem = Bool()
    showDropZones = Bool()
    createdVersion = Integer()
    indent = Integer()
    showEmptyRow = Bool()
    showEmptyCol = Bool()
    showHeaders = Bool()
    compact = Bool()
    outline = Bool()
    outlineData = Bool()
    compactData = Bool()
    published = Bool()
    gridDropZones = Bool()
    immersive = Bool()
    multipleFieldFilters = Bool()
    chartFormat = Integer()
    rowHeaderCaption = String()
    colHeaderCaption = String()
    fieldListSortAscending = Bool()
    mdxSubqueries = Bool()
    customListSort = Bool(allow_none=True)
    autoFormatId = Integer()
    applyNumberFormats = Bool()
    applyBorderFormats = Bool()
    applyFontFormats = Bool()
    applyPatternFormats = Bool()
    applyAlignmentFormats = Bool()
    applyWidthHeightFormats = Bool()
    location = Typed(expected_type=Location, )
    pivotFields = Typed(expected_type=PivotFields, allow_none=True)
    rowFields = Typed(expected_type=RowFields, allow_none=True)
    rowItems = Typed(expected_type=rowItems, allow_none=True)
    colFields = Typed(expected_type=ColFields, allow_none=True)
    colItems = Typed(expected_type=colItems, allow_none=True)
    pageFields = Typed(expected_type=PageFields, allow_none=True)
    dataFields = Typed(expected_type=DataFields, allow_none=True)
    formats = Typed(expected_type=Formats, allow_none=True)
    conditionalFormats = Typed(expected_type=ConditionalFormats, allow_none=True)
    chartFormats = Typed(expected_type=ChartFormats, allow_none=True)
    pivotHierarchies = Typed(expected_type=PivotHierarchies, allow_none=True)
    pivotTableStyleInfo = Typed(expected_type=PivotTableStyle, allow_none=True)
    filters = Typed(expected_type=PivotFilters, allow_none=True)
    rowHierarchiesUsage = Typed(expected_type=RowHierarchiesUsage, allow_none=True)
    colHierarchiesUsage = Typed(expected_type=ColHierarchiesUsage, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('location', 'pivotFields', 'rowFields', 'rowItems', 'colFields', 'colItems', 'pageFields', 'dataFields', 'formats', 'conditionalFormats', 'chartFormats', 'pivotHierarchies', 'pivotTableStyleInfo', 'filters', 'rowHierarchiesUsage', 'colHierarchiesUsage', 'extLst')

    def __init__(self,
                 name=None,
                 cacheId=None,
                 dataOnRows=None,
                 dataPosition=None,
                 dataCaption=None,
                 grandTotalCaption=None,
                 errorCaption=None,
                 showError=None,
                 missingCaption=None,
                 showMissing=None,
                 pageStyle=None,
                 pivotTableStyle=None,
                 vacatedStyle=None,
                 tag=None,
                 updatedVersion=None,
                 minRefreshableVersion=None,
                 asteriskTotals=None,
                 showItems=None,
                 editData=None,
                 disableFieldList=None,
                 showCalcMbrs=None,
                 visualTotals=None,
                 showMultipleLabel=None,
                 showDataDropDown=None,
                 showDrill=None,
                 printDrill=None,
                 showMemberPropertyTips=None,
                 showDataTips=None,
                 enableWizard=None,
                 enableDrill=None,
                 enableFieldProperties=None,
                 preserveFormatting=None,
                 useAutoFormatting=None,
                 pageWrap=None,
                 pageOverThenDown=None,
                 subtotalHiddenItems=None,
                 rowGrandTotals=None,
                 colGrandTotals=None,
                 fieldPrintTitles=None,
                 itemPrintTitles=None,
                 mergeItem=None,
                 showDropZones=None,
                 createdVersion=None,
                 indent=None,
                 showEmptyRow=None,
                 showEmptyCol=None,
                 showHeaders=None,
                 compact=None,
                 outline=None,
                 outlineData=None,
                 compactData=None,
                 published=None,
                 gridDropZones=None,
                 immersive=None,
                 multipleFieldFilters=None,
                 chartFormat=None,
                 rowHeaderCaption=None,
                 colHeaderCaption=None,
                 fieldListSortAscending=None,
                 mdxSubqueries=None,
                 customListSort=None,
                 autoFormatId=None,
                 applyNumberFormats=None,
                 applyBorderFormats=None,
                 applyFontFormats=None,
                 applyPatternFormats=None,
                 applyAlignmentFormats=None,
                 applyWidthHeightFormats=None,
                 location=None,
                 pivotFields=None,
                 rowFields=None,
                 rowItems=None,
                 colFields=None,
                 colItems=None,
                 pageFields=None,
                 dataFields=None,
                 formats=None,
                 conditionalFormats=None,
                 chartFormats=None,
                 pivotHierarchies=None,
                 pivotTableStyleInfo=None,
                 filters=None,
                 rowHierarchiesUsage=None,
                 colHierarchiesUsage=None,
                 extLst=None,
                ):
        self.name = name
        self.cacheId = cacheId
        self.dataOnRows = dataOnRows
        self.dataPosition = dataPosition
        self.dataCaption = dataCaption
        self.grandTotalCaption = grandTotalCaption
        self.errorCaption = errorCaption
        self.showError = showError
        self.missingCaption = missingCaption
        self.showMissing = showMissing
        self.pageStyle = pageStyle
        self.pivotTableStyle = pivotTableStyle
        self.vacatedStyle = vacatedStyle
        self.tag = tag
        self.updatedVersion = updatedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.asteriskTotals = asteriskTotals
        self.showItems = showItems
        self.editData = editData
        self.disableFieldList = disableFieldList
        self.showCalcMbrs = showCalcMbrs
        self.visualTotals = visualTotals
        self.showMultipleLabel = showMultipleLabel
        self.showDataDropDown = showDataDropDown
        self.showDrill = showDrill
        self.printDrill = printDrill
        self.showMemberPropertyTips = showMemberPropertyTips
        self.showDataTips = showDataTips
        self.enableWizard = enableWizard
        self.enableDrill = enableDrill
        self.enableFieldProperties = enableFieldProperties
        self.preserveFormatting = preserveFormatting
        self.useAutoFormatting = useAutoFormatting
        self.pageWrap = pageWrap
        self.pageOverThenDown = pageOverThenDown
        self.subtotalHiddenItems = subtotalHiddenItems
        self.rowGrandTotals = rowGrandTotals
        self.colGrandTotals = colGrandTotals
        self.fieldPrintTitles = fieldPrintTitles
        self.itemPrintTitles = itemPrintTitles
        self.mergeItem = mergeItem
        self.showDropZones = showDropZones
        self.createdVersion = createdVersion
        self.indent = indent
        self.showEmptyRow = showEmptyRow
        self.showEmptyCol = showEmptyCol
        self.showHeaders = showHeaders
        self.compact = compact
        self.outline = outline
        self.outlineData = outlineData
        self.compactData = compactData
        self.published = published
        self.gridDropZones = gridDropZones
        self.immersive = immersive
        self.multipleFieldFilters = multipleFieldFilters
        self.chartFormat = chartFormat
        self.rowHeaderCaption = rowHeaderCaption
        self.colHeaderCaption = colHeaderCaption
        self.fieldListSortAscending = fieldListSortAscending
        self.mdxSubqueries = mdxSubqueries
        self.customListSort = customListSort
        self.autoFormatId = autoFormatId
        self.applyNumberFormats = applyNumberFormats
        self.applyBorderFormats = applyBorderFormats
        self.applyFontFormats = applyFontFormats
        self.applyPatternFormats = applyPatternFormats
        self.applyAlignmentFormats = applyAlignmentFormats
        self.applyWidthHeightFormats = applyWidthHeightFormats
        self.location = location
        self.pivotFields = pivotFields
        self.rowFields = rowFields
        self.rowItems = rowItems
        self.colFields = colFields
        self.colItems = colItems
        self.pageFields = pageFields
        self.dataFields = dataFields
        self.formats = formats
        self.conditionalFormats = conditionalFormats
        self.chartFormats = chartFormats
        self.pivotHierarchies = pivotHierarchies
        self.pivotTableStyleInfo = pivotTableStyleInfo
        self.filters = filters
        self.rowHierarchiesUsage = rowHierarchiesUsage
        self.colHierarchiesUsage = colHierarchiesUsage
        self.extLst = extLst
