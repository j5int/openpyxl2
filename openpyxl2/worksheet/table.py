from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Set,
    Float,
    DateTime,
    Bool,
    Integer,
    NoneSet,
    String,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.sequence import NestedSequence

from .filters import (
    AutoFilter,
    CustomFilter,
    CustomFilters,
    Filters,
    IconFilter,
    ColorFilter,
    DynamicFilter,
    FilterColumn,
    SortCondition,
    SortState
)


class TableStyleInfo(Serialisable):

    name = String(allow_none=True)
    showFirstColumn = Bool(allow_none=True)
    showLastColumn = Bool(allow_none=True)
    showRowStripes = Bool(allow_none=True)
    showColumnStripes = Bool(allow_none=True)

    def __init__(self,
                 name=None,
                 showFirstColumn=None,
                 showLastColumn=None,
                 showRowStripes=None,
                 showColumnStripes=None,
                ):
        self.name = name
        self.showFirstColumn = showFirstColumn
        self.showLastColumn = showLastColumn
        self.showRowStripes = showRowStripes
        self.showColumnStripes = showColumnStripes


class XmlColumnPr(Serialisable):

    mapId = Integer()
    xpath = String()
    denormalized = Bool(allow_none=True)
    xmlDataType = String()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 mapId=None,
                 xpath=None,
                 denormalized=None,
                 xmlDataType=None,
                 extLst=None,
                ):
        self.mapId = mapId
        self.xpath = xpath
        self.denormalized = denormalized
        self.xmlDataType = xmlDataType
        self.extLst = extLst


class TableFormula(Serialisable):

    array = Bool()

    def __init__(self,
                 array=None,
                ):
        self.array = array


class TableColumn(Serialisable):

    tagname = "tableColumn"

    id = Integer()
    uniqueName = String(allow_none=True)
    name = String()
    totalsRowFunction = NoneSet(values=(['sum', 'min', 'max', 'average',
                                         'count', 'countNums', 'stdDev', 'var', 'custom']))
    totalsRowLabel = String(allow_none=True)
    queryTableFieldId = Integer(allow_none=True)
    headerRowDxfId = Integer(allow_none=True)
    dataDxfId = Integer(allow_none=True)
    totalsRowDxfId = Integer(allow_none=True)
    headerRowCellStyle = String(allow_none=True)
    dataCellStyle = String(allow_none=True)
    totalsRowCellStyle = String(allow_none=True)
    calculatedColumnFormula = Typed(expected_type=TableFormula, allow_none=True)
    totalsRowFormula = Typed(expected_type=TableFormula, allow_none=True)
    xmlColumnPr = Typed(expected_type=XmlColumnPr, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('calculatedColumnFormula', 'totalsRowFormula',
                    'xmlColumnPr', 'extLst')

    def __init__(self,
                 id=None,
                 uniqueName=None,
                 name=None,
                 totalsRowFunction=None,
                 totalsRowLabel=None,
                 queryTableFieldId=None,
                 headerRowDxfId=None,
                 dataDxfId=None,
                 totalsRowDxfId=None,
                 headerRowCellStyle=None,
                 dataCellStyle=None,
                 totalsRowCellStyle=None,
                 calculatedColumnFormula=None,
                 totalsRowFormula=None,
                 xmlColumnPr=None,
                 extLst=None,
                ):
        self.id = id
        self.uniqueName = uniqueName
        self.name = name
        self.totalsRowFunction = totalsRowFunction
        self.totalsRowLabel = totalsRowLabel
        self.queryTableFieldId = queryTableFieldId
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.calculatedColumnFormula = calculatedColumnFormula
        self.totalsRowFormula = totalsRowFormula
        self.xmlColumnPr = xmlColumnPr
        self.extLst = extLst


class TableColumns(Serialisable):

    count = Integer(allow_none=True)
    tableColumn = Typed(expected_type=TableColumn, )

    __elements__ = ('tableColumn',)

    def __init__(self,
                 count=None,
                 tableColumn=None,
                ):
        self.count = count
        self.tableColumn = tableColumn


class Top10(Serialisable):

    top = Bool(allow_none=True)
    percent = Bool(allow_none=True)
    val = Float()
    filterVal = Float(allow_none=True)

    def __init__(self,
                 top=None,
                 percent=None,
                 val=None,
                 filterVal=None,
                ):
        self.top = top
        self.percent = percent
        self.val = val
        self.filterVal = filterVal


class DateGroupItem(Serialisable):

    year = Integer()
    month = Integer(allow_none=True)
    day = Integer(allow_none=True)
    hour = Integer(allow_none=True)
    minute = Integer(allow_none=True)
    second = Integer(allow_none=True)
    dateTimeGrouping = Set(values=(['year', 'month', 'day', 'hour', 'minute',
                                    'second']))

    def __init__(self,
                 year=None,
                 month=None,
                 day=None,
                 hour=None,
                 minute=None,
                 second=None,
                 dateTimeGrouping=None,
                ):
        self.year = year
        self.month = month
        self.day = day
        self.hour = hour
        self.minute = minute
        self.second = second
        self.dateTimeGrouping = dateTimeGrouping


class Table(Serialisable):

    tagname = "table"

    id = Integer()
    name = String(allow_none=True)
    displayName = String()
    comment = String(allow_none=True)
    ref = String()
    tableType = NoneSet(values=(['worksheet', 'xml', 'queryTable']))
    headerRowCount = Integer(allow_none=True)
    insertRow = Bool(allow_none=True)
    insertRowShift = Bool(allow_none=True)
    totalsRowCount = Integer(allow_none=True)
    totalsRowShown = Bool(allow_none=True)
    published = Bool(allow_none=True)
    headerRowDxfId = Integer(allow_none=True)
    dataDxfId = Integer(allow_none=True)
    totalsRowDxfId = Integer(allow_none=True)
    headerRowBorderDxfId = Integer(allow_none=True)
    tableBorderDxfId = Integer(allow_none=True)
    totalsRowBorderDxfId = Integer(allow_none=True)
    headerRowCellStyle = String(allow_none=True)
    dataCellStyle = String(allow_none=True)
    totalsRowCellStyle = String(allow_none=True)
    connectionId = Integer(allow_none=True)
    autoFilter = Typed(expected_type=AutoFilter, allow_none=True)
    sortState = Typed(expected_type=SortState, allow_none=True)
    tableColumns = NestedSequence(expected_type=TableColumn, count=True)
    tableStyleInfo = Typed(expected_type=TableStyleInfo, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('autoFilter', 'sortState', 'tableColumns',
                    'tableStyleInfo')

    def __init__(self,
                 id=None,
                 name=None,
                 displayName=None,
                 comment=None,
                 ref=None,
                 tableType=None,
                 headerRowCount=None,
                 insertRow=None,
                 insertRowShift=None,
                 totalsRowCount=None,
                 totalsRowShown=None,
                 published=None,
                 headerRowDxfId=None,
                 dataDxfId=None,
                 totalsRowDxfId=None,
                 headerRowBorderDxfId=None,
                 tableBorderDxfId=None,
                 totalsRowBorderDxfId=None,
                 headerRowCellStyle=None,
                 dataCellStyle=None,
                 totalsRowCellStyle=None,
                 connectionId=None,
                 autoFilter=None,
                 sortState=None,
                 tableColumns=(),
                 tableStyleInfo=None,
                 extLst=None,
                ):
        self.id = id
        self.name = name
        self.displayName = displayName
        self.comment = comment
        self.ref = ref
        self.tableType = tableType
        self.headerRowCount = headerRowCount
        self.insertRow = insertRow
        self.insertRowShift = insertRowShift
        self.totalsRowCount = totalsRowCount
        self.totalsRowShown = totalsRowShown
        self.published = published
        self.headerRowDxfId = headerRowDxfId
        self.dataDxfId = dataDxfId
        self.totalsRowDxfId = totalsRowDxfId
        self.headerRowBorderDxfId = headerRowBorderDxfId
        self.tableBorderDxfId = tableBorderDxfId
        self.totalsRowBorderDxfId = totalsRowBorderDxfId
        self.headerRowCellStyle = headerRowCellStyle
        self.dataCellStyle = dataCellStyle
        self.totalsRowCellStyle = totalsRowCellStyle
        self.connectionId = connectionId
        self.autoFilter = autoFilter
        self.sortState = sortState
        self.tableColumns = tableColumns
        self.tableStyleInfo = tableStyleInfo
