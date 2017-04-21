# Copyright (c) 2010-2017 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Bool,
    Float,
    Set,
    NoneSet,
    String,
    Integer,
    DateTime,
    Sequence,
)

from openpyxl2.descriptors.excel import HexBinary, ExtensionList
from openpyxl2.descriptors.nested import NestedInteger
from openpyxl2.descriptors.sequence import NestedSequence
from openpyxl2.xml.constants import SHEET_MAIN_NS

from .pivot import (
    PivotArea,
    PivotAreaReference,
    PivotAreaReferenceList,
)
from .record import (
    Boolean,
    Error,
    Missing,
    Number,
    Text,
    TupleList,
    PivotDateTime,
    SharedItem,
)

class MeasureDimensionMap(Serialisable):

    measureGroup = Integer(allow_none=True)
    dimension = Integer(allow_none=True)

    def __init__(self,
                 measureGroup=None,
                 dimension=None,
                ):
        self.measureGroup = measureGroup
        self.dimension = dimension


class MeasureDimensionMaps(Serialisable):

    count = Integer()
    map = Typed(expected_type=MeasureDimensionMap, allow_none=True)

    __elements__ = ('map',)

    def __init__(self,
                 count=None,
                 map=None,
                ):
        self.count = count
        self.map = map


class MeasureGroup(Serialisable):

    name = String()
    caption = String()

    def __init__(self,
                 name=None,
                 caption=None,
                ):
        self.name = name
        self.caption = caption


class MeasureGroups(Serialisable):

    count = Integer()
    measureGroup = Typed(expected_type=MeasureGroup, allow_none=True)

    __elements__ = ('measureGroup',)

    def __init__(self,
                 count=None,
                 measureGroup=None,
                ):
        self.count = count
        self.measureGroup = measureGroup


class PivotDimension(Serialisable):

    measure = Bool()
    name = String()
    uniqueName = String()
    caption = String()

    def __init__(self,
                 measure=None,
                 name=None,
                 uniqueName=None,
                 caption=None,
                ):
        self.measure = measure
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption


class Dimensions(Serialisable):

    count = Integer()
    dimension = Typed(expected_type=PivotDimension, allow_none=True)

    __elements__ = ('dimension',)

    def __init__(self,
                 count=None,
                 dimension=None,
                ):
        self.count = count
        self.dimension = dimension


class CalculatedMember(Serialisable):

    name = String()
    mdx = String()
    memberName = String()
    hierarchy = String()
    parent = String()
    solveOrder = Integer()
    set = Bool()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('extLst',)

    def __init__(self,
                 name=None,
                 mdx=None,
                 memberName=None,
                 hierarchy=None,
                 parent=None,
                 solveOrder=None,
                 set=None,
                 extLst=None,
                ):
        self.name = name
        self.mdx = mdx
        self.memberName = memberName
        self.hierarchy = hierarchy
        self.parent = parent
        self.solveOrder = solveOrder
        self.set = set
        self.extLst = extLst


class CalculatedMembers(Serialisable):

    count = Integer()
    calculatedMember = Typed(expected_type=CalculatedMember, )

    __elements__ = ('calculatedMember',)

    def __init__(self,
                 count=None,
                 calculatedMember=None,
                ):
        self.count = count
        self.calculatedMember = calculatedMember


class CalculatedItem(Serialisable):

    field = Integer(allow_none=True)
    formula = String()
    pivotArea = Typed(expected_type=PivotArea, )
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('pivotArea', 'extLst')

    def __init__(self,
                 field=None,
                 formula=None,
                 pivotArea=None,
                 extLst=None,
                ):
        self.field = field
        self.formula = formula
        self.pivotArea = pivotArea
        self.extLst = extLst


class CalculatedItems(Serialisable):

    count = Integer()
    calculatedItem = Typed(expected_type=CalculatedItem, )

    __elements__ = ('calculatedItem',)

    def __init__(self,
                 count=None,
                 calculatedItem=None,
                ):
        self.count = count
        self.calculatedItem = calculatedItem


class ServerFormat(Serialisable):

    culture = String(allow_none=True)
    format = String(allow_none=True)

    def __init__(self,
                 culture=None,
                 format=None,
                ):
        self.culture = culture
        self.format = format


class ServerFormats(Serialisable):

    count = Integer()
    serverFormat = Typed(expected_type=ServerFormat, allow_none=True)

    __elements__ = ('serverFormat',)

    def __init__(self,
                 count=None,
                 serverFormat=None,
                ):
        self.count = count
        self.serverFormat = serverFormat


class Query(Serialisable):

    mdx = String()
    tpls = Typed(expected_type=TupleList, allow_none=True)

    __elements__ = ('tpls',)

    def __init__(self,
                 mdx=None,
                 tpls=None,
                ):
        self.mdx = mdx
        self.tpls = tpls


class QueryCache(Serialisable):

    count = Integer()
    query = Typed(expected_type=Query, )

    __elements__ = ('query',)

    def __init__(self,
                 count=None,
                 query=None,
                ):
        self.count = count
        self.query = query


class OLAPSet(Serialisable):

    count = Integer()
    maxRank = Integer()
    setDefinition = String()
    sortType = NoneSet(values=(['ascending', 'descending', 'ascendingAlpha', 'descendingAlpha', 'ascendingNatural', 'descendingNatural']))
    queryFailed = Bool()
    tpls = Typed(expected_type=TupleList, allow_none=True)
    sortByTuple = Typed(expected_type=TupleList, allow_none=True)

    __elements__ = ('tpls', 'sortByTuple')

    def __init__(self,
                 count=None,
                 maxRank=None,
                 setDefinition=None,
                 sortType=None,
                 queryFailed=None,
                 tpls=None,
                 sortByTuple=None,
                ):
        self.count = count
        self.maxRank = maxRank
        self.setDefinition = setDefinition
        self.sortType = sortType
        self.queryFailed = queryFailed
        self.tpls = tpls
        self.sortByTuple = sortByTuple


class OLAPSets(Serialisable):

    count = Integer()
    set = Typed(expected_type=OLAPSet, )

    __elements__ = ('set',)

    def __init__(self,
                 count=None,
                 set=None,
                ):
        self.count = count
        self.set = set


class PCDSDTCEntries(Serialisable):

    count = Integer()
    # some elements are choice
    m = Typed(expected_type=Missing, )
    n = Typed(expected_type=Number, )
    e = Typed(expected_type=Error, )
    s = Typed(expected_type=Text)

    __elements__ = ('m', 'n', 'e', 's')

    def __init__(self,
                 count=None,
                 m=None,
                 n=None,
                 e=None,
                 s=None,
                ):
        self.count = count
        self.m = m
        self.n = n
        self.e = e
        self.s = s


class TupleCache(Serialisable):

    entries = Typed(expected_type=PCDSDTCEntries, allow_none=True)
    sets = Typed(expected_type=OLAPSets, allow_none=True)
    queryCache = Typed(expected_type=QueryCache, allow_none=True)
    serverFormats = Typed(expected_type=ServerFormats, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('entries', 'sets', 'queryCache', 'serverFormats', 'extLst')

    def __init__(self,
                 entries=None,
                 sets=None,
                 queryCache=None,
                 serverFormats=None,
                 extLst=None,
                ):
        self.entries = entries
        self.sets = sets
        self.queryCache = queryCache
        self.serverFormats = serverFormats
        self.extLst = extLst


class PCDKPI(Serialisable):

    uniqueName = String()
    caption = String(allow_none=True)
    displayFolder = String()
    measureGroup = String()
    parent = String()
    value = String()
    goal = String()
    status = String()
    trend = String()
    weight = String()
    time = String()

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 displayFolder=None,
                 measureGroup=None,
                 parent=None,
                 value=None,
                 goal=None,
                 status=None,
                 trend=None,
                 weight=None,
                 time=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.parent = parent
        self.value = value
        self.goal = goal
        self.status = status
        self.trend = trend
        self.weight = weight
        self.time = time


class PCDKPIs(Serialisable):

    count = Integer()
    kpi = Typed(expected_type=PCDKPI, allow_none=True)

    __elements__ = ('kpi',)

    def __init__(self,
                 count=None,
                 kpi=None,
                ):
        self.count = count
        self.kpi = kpi


class GroupMember(Serialisable):

    uniqueName = String()
    group = Bool()

    def __init__(self,
                 uniqueName=None,
                 group=None,
                ):
        self.uniqueName = uniqueName
        self.group = group


class GroupMembers(Serialisable):

    count = Integer()
    groupMember = Typed(expected_type=GroupMember, )

    __elements__ = ('groupMember',)

    def __init__(self,
                 count=None,
                 groupMember=None,
                ):
        self.count = count
        self.groupMember = groupMember


class LevelGroup(Serialisable):

    name = String()
    uniqueName = String()
    caption = String()
    uniqueParent = String()
    id = Integer()
    groupMembers = Typed(expected_type=GroupMembers, )

    __elements__ = ('groupMembers',)

    def __init__(self,
                 name=None,
                 uniqueName=None,
                 caption=None,
                 uniqueParent=None,
                 id=None,
                 groupMembers=None,
                ):
        self.name = name
        self.uniqueName = uniqueName
        self.caption = caption
        self.uniqueParent = uniqueParent
        self.id = id
        self.groupMembers = groupMembers


class Groups(Serialisable):

    count = Integer()
    group = Typed(expected_type=LevelGroup, )

    __elements__ = ('group',)

    def __init__(self,
                 count=None,
                 group=None,
                ):
        self.count = count
        self.group = group


class GroupLevel(Serialisable):

    uniqueName = String()
    caption = String()
    user = Bool()
    customRollUp = Bool()
    groups = Typed(expected_type=Groups, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('groups', 'extLst')

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 user=None,
                 customRollUp=None,
                 groups=None,
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.user = user
        self.customRollUp = customRollUp
        self.groups = groups
        self.extLst = extLst


class GroupLevels(Serialisable):

    count = Integer()
    groupLevel = Typed(expected_type=GroupLevel, )

    __elements__ = ('groupLevel',)

    def __init__(self,
                 count=None,
                 groupLevel=None,
                ):
        self.count = count
        self.groupLevel = groupLevel


class FieldUsage(Serialisable):

    x = Integer()

    def __init__(self,
                 x=None,
                ):
        self.x = x


class FieldsUsage(Serialisable):

    count = Integer()
    fieldUsage = Typed(expected_type=FieldUsage, allow_none=True)

    __elements__ = ('fieldUsage',)

    def __init__(self,
                 count=None,
                 fieldUsage=None,
                ):
        self.count = count
        self.fieldUsage = fieldUsage


class CacheHierarchy(Serialisable):

    uniqueName = String()
    caption = String(allow_none=True)
    measure = Bool()
    set = Bool()
    parentSet = Integer(allow_none=True)
    iconSet = Integer()
    attribute = Bool()
    time = Bool()
    keyAttribute = Bool()
    defaultMemberUniqueName = String()
    allUniqueName = String()
    allCaption = String()
    dimensionUniqueName = String()
    displayFolder = String()
    measureGroup = String()
    measures = Bool()
    count = Integer()
    oneField = Bool()
    memberValueDatatype = Integer(allow_none=True)
    unbalanced = Bool(allow_none=True)
    unbalancedGroup = Bool(allow_none=True)
    hidden = Bool()
    fieldsUsage = Typed(expected_type=FieldsUsage, allow_none=True)
    groupLevels = Typed(expected_type=GroupLevels, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('fieldsUsage', 'groupLevels', 'extLst')

    def __init__(self,
                 uniqueName=None,
                 caption=None,
                 measure=None,
                 set=None,
                 parentSet=None,
                 iconSet=None,
                 attribute=None,
                 time=None,
                 keyAttribute=None,
                 defaultMemberUniqueName=None,
                 allUniqueName=None,
                 allCaption=None,
                 dimensionUniqueName=None,
                 displayFolder=None,
                 measureGroup=None,
                 measures=None,
                 count=None,
                 oneField=None,
                 memberValueDatatype=None,
                 unbalanced=None,
                 unbalancedGroup=None,
                 hidden=None,
                 fieldsUsage=None,
                 groupLevels=None,
                 extLst=None,
                ):
        self.uniqueName = uniqueName
        self.caption = caption
        self.measure = measure
        self.set = set
        self.parentSet = parentSet
        self.iconSet = iconSet
        self.attribute = attribute
        self.time = time
        self.keyAttribute = keyAttribute
        self.defaultMemberUniqueName = defaultMemberUniqueName
        self.allUniqueName = allUniqueName
        self.allCaption = allCaption
        self.dimensionUniqueName = dimensionUniqueName
        self.displayFolder = displayFolder
        self.measureGroup = measureGroup
        self.measures = measures
        self.count = count
        self.oneField = oneField
        self.memberValueDatatype = memberValueDatatype
        self.unbalanced = unbalanced
        self.unbalancedGroup = unbalancedGroup
        self.hidden = hidden
        self.fieldsUsage = fieldsUsage
        self.groupLevels = groupLevels
        self.extLst = extLst


class CacheHierarchies(Serialisable):

    count = Integer()
    cacheHierarchy = Typed(expected_type=CacheHierarchy, allow_none=True)

    __elements__ = ('cacheHierarchy',)

    def __init__(self,
                 count=None,
                 cacheHierarchy=None,
                ):
        self.count = count
        self.cacheHierarchy = cacheHierarchy


class GroupItems(Serialisable):

    count = Integer()
    # some elements are choice
    m = Typed(expected_type=Missing, )
    n = Typed(expected_type=Number, )
    b = Bool(nested=True, )
    e = Typed(expected_type=Error, )
    s = Typed(expected_type=Text)
    d = Typed(xpected_type=PivotDateTime,)

    __elements__ = ('m', 'n', 'b', 'e', 's', 'd')

    def __init__(self,
                 count=None,
                 m=None,
                 n=None,
                 b=None,
                 e=None,
                 s=None,
                 d=None,
                ):
        self.count = count
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d


class DiscretePr(Serialisable):

    count = Integer()
    x = NestedInteger(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 count=None,
                 x=None,
                ):
        self.count = count
        self.x = x


class RangePr(Serialisable):

    autoStart = Bool()
    autoEnd = Bool()
    groupBy = Set(values=(['range', 'seconds', 'minutes', 'hours', 'days', 'months', 'quarters', 'years']))
    startNum = Float()
    endNum = Float()
    startDate = DateTime()
    endDate = DateTime()
    groupInterval = Float()

    def __init__(self,
                 autoStart=None,
                 autoEnd=None,
                 groupBy=None,
                 startNum=None,
                 endNum=None,
                 startDate=None,
                 endDate=None,
                 groupInterval=None,
                ):
        self.autoStart = autoStart
        self.autoEnd = autoEnd
        self.groupBy = groupBy
        self.startNum = startNum
        self.endNum = endNum
        self.startDate = startDate
        self.endDate = endDate
        self.groupInterval = groupInterval


class FieldGroup(Serialisable):

    par = Integer(allow_none=True)
    base = Integer(allow_none=True)
    rangePr = Typed(expected_type=RangePr, allow_none=True)
    discretePr = Typed(expected_type=DiscretePr, allow_none=True)
    groupItems = Typed(expected_type=GroupItems, allow_none=True)

    __elements__ = ('rangePr', 'discretePr', 'groupItems')

    def __init__(self,
                 par=None,
                 base=None,
                 rangePr=None,
                 discretePr=None,
                 groupItems=None,
                ):
        self.par = par
        self.base = base
        self.rangePr = rangePr
        self.discretePr = discretePr
        self.groupItems = groupItems


class SharedItems(Serialisable):

    tagname = "sharedItems"

    m = Sequence(expected_type=Missing)
    n = Sequence(expected_type=Number)
    b = Sequence(expected_type=Boolean)
    e = Sequence(expected_type=Error)
    s = Sequence(expected_type=Text)
    d = Sequence(expected_type=PivotDateTime)
    # attributes are optional and must be derived from associated cache records
    containsSemiMixedTypes = Bool(allow_none=True)
    containsNonDate = Bool(allow_none=True)
    containsDate = Bool(allow_none=True)
    containsString = Bool(allow_none=True)
    containsBlank = Bool(allow_none=True)
    containsMixedTypes = Bool(allow_none=True)
    containsNumber = Bool(allow_none=True)
    containsInteger = Bool(allow_none=True)
    minValue = Float(allow_none=True)
    maxValue = Float(allow_none=True)
    minDate = DateTime(allow_none=True)
    maxDate = DateTime(allow_none=True)
    longText = Bool(allow_none=True)

    __elements__ = ('m', 'n', 'b', 'e', 's', 'd')
    __attrs__ = ('count',)

    def __init__(self,
                 m=(),
                 n=(),
                 b=(),
                 e=(),
                 s=(),
                 d=(),
                 containsSemiMixedTypes=True,
                 containsNonDate=True,
                 containsDate=False,
                 containsString=True,
                 containsBlank=False,
                 containsMixedTypes=False,
                 containsNumber=False,
                 containsInteger=False,
                 minValue=None,
                 maxValue=None,
                 minDate=None,
                 maxDate=None,
                 count=None,
                 longText=False,
                ):
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d


    @property
    def count(self):
        return len(self.m + self.n + self.b + self.e + self.s + self.d)


class CacheField(Serialisable):

    tagname = "cacheField"

    sharedItems = Typed(expected_type=SharedItems, allow_none=True)
    fieldGroup = Typed(expected_type=FieldGroup, allow_none=True)
    mpMap = NestedInteger(allow_none=True, attribute="v")
    extLst = Typed(expected_type=ExtensionList, allow_none=True)
    name = String()
    caption = String(allow_none=True)
    propertyName = String(allow_none=True)
    serverField = Bool(allow_none=True)
    uniqueList = Bool(allow_none=True)
    numFmtId = Integer(allow_none=True)
    formula = String(allow_none=True)
    sqlType = Integer(allow_none=True)
    hierarchy = Integer(allow_none=True)
    level = Integer(allow_none=True)
    databaseField = Bool(allow_none=True)
    mappingCount = Integer(allow_none=True)
    memberPropertyField = Bool(allow_none=True)

    __elements__ = ('sharedItems', 'fieldGroup', 'mpMap')

    def __init__(self,
                 sharedItems=None,
                 fieldGroup=None,
                 mpMap=None,
                 extLst=None,
                 name=None,
                 caption=None,
                 propertyName=None,
                 serverField=False,
                 uniqueList=True,
                 numFmtId=None,
                 formula=None,
                 sqlType=0,
                 hierarchy=0,
                 level=0,
                 databaseField=True,
                 mappingCount=None,
                 memberPropertyField=False,
                ):
        self.sharedItems = sharedItems
        self.fieldGroup = fieldGroup
        self.mpMap = mpMap
        self.extLst = extLst
        self.name = name
        self.caption = caption
        self.propertyName = propertyName
        self.serverField = serverField
        self.uniqueList = uniqueList
        self.numFmtId = numFmtId
        self.formula = formula
        self.sqlType = sqlType
        self.hierarchy = hierarchy
        self.level = level
        self.databaseField = databaseField
        self.mappingCount = mappingCount
        self.memberPropertyField = memberPropertyField


class RangeSet(Serialisable):

    i1 = Integer(allow_none=True)
    i2 = Integer(allow_none=True)
    i3 = Integer(allow_none=True)
    i4 = Integer(allow_none=True)
    ref = String()
    name = String(allow_none=True)
    sheet = String(allow_none=True)

    def __init__(self,
                 i1=None,
                 i2=None,
                 i3=None,
                 i4=None,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.i1 = i1
        self.i2 = i2
        self.i3 = i3
        self.i4 = i4
        self.ref = ref
        self.name = name
        self.sheet = sheet


class RangeSets(Serialisable):

    count = Integer(allow_none=True)
    rangeSet = Typed(expected_type=RangeSet, )

    __elements__ = ('rangeSet',)

    def __init__(self,
                 count=None,
                 rangeSet=None,
                ):
        self.count = count
        self.rangeSet = rangeSet


class PageItem(Serialisable):

    name = String()

    def __init__(self,
                 name=None,
                ):
        self.name = name


class PCDSCPage(Serialisable):

    count = Integer(allow_none=True)
    pageItem = Typed(expected_type=PageItem, allow_none=True)

    __elements__ = ('pageItem',)

    def __init__(self,
                 count=None,
                 pageItem=None,
                ):
        self.count = count
        self.pageItem = pageItem


class Pages(Serialisable):

    count = Integer(allow_none=True)
    page = Typed(expected_type=PCDSCPage, )

    __elements__ = ('page',)

    def __init__(self,
                 count=None,
                 page=None,
                ):
        self.count = count
        self.page = page


class Consolidation(Serialisable):

    autoPage = Bool(allow_none=True)
    pages = Typed(expected_type=Pages, allow_none=True)
    rangeSets = Typed(expected_type=RangeSets, )

    __elements__ = ('pages', 'rangeSets')

    def __init__(self,
                 autoPage=None,
                 pages=None,
                 rangeSets=None,
                ):
        self.autoPage = autoPage
        self.pages = pages
        self.rangeSets = rangeSets


class WorksheetSource(Serialisable):

    tagname = "worksheetSource"

    ref = String(allow_none=True)
    name = String(allow_none=True)
    sheet = String(allow_none=True)

    def __init__(self,
                 ref=None,
                 name=None,
                 sheet=None,
                ):
        self.ref = ref
        self.name = name
        self.sheet = sheet


class CacheSource(Serialisable):

    tagname = "cacheSource"

    type = Set(values=(['worksheet', 'external', 'consolidation', 'scenario']))
    connectionId = Integer(allow_none=True)
    # some elements are choice
    worksheetSource = Typed(expected_type=WorksheetSource, )
    consolidation = Typed(expected_type=Consolidation, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('worksheetSource', 'consolidation',)

    def __init__(self,
                 type=None,
                 connectionId=None,
                 worksheetSource=None,
                 consolidation=None,
                 extLst=None,
                ):
        self.type = type
        self.connectionId = connectionId
        self.worksheetSource = worksheetSource
        self.consolidation = consolidation


class PivotCacheDefinition(Serialisable):

    tagname = "pivotCacheDefinition"

    invalid = Bool(allow_none=True)
    saveData = Bool(allow_none=True)
    refreshOnLoad = Bool(allow_none=True)
    optimizeMemory = Bool(allow_none=True)
    enableRefresh = Bool(allow_none=True)
    refreshedBy = String(allow_none=True)
    refreshedDate = Float(allow_none=True)
    refreshedDateIso = DateTime(allow_none=True)
    backgroundQuery = Bool()
    missingItemsLimit = Integer(allow_none=True)
    createdVersion = Integer(allow_none=True)
    refreshedVersion = Integer(allow_none=True)
    minRefreshableVersion = Integer(allow_none=True)
    recordCount = Integer(allow_none=True)
    upgradeOnRefresh = Bool(allow_none=True)
    tupleCache = Bool(allow_none=True)
    supportSubquery = Bool(allow_none=True)
    supportAdvancedDrill = Bool(allow_none=True)
    cacheSource = Typed(expected_type=CacheSource, )
    cacheFields = NestedSequence(expected_type=CacheField, count=True)
    cacheHierarchies = Typed(expected_type=CacheHierarchies, allow_none=True)
    kpis = Typed(expected_type=PCDKPIs, allow_none=True)
    tupleCache = Typed(expected_type=TupleCache, allow_none=True)
    calculatedItems = Typed(expected_type=CalculatedItems, allow_none=True)
    calculatedMembers = Typed(expected_type=CalculatedMembers, allow_none=True)
    dimensions = Typed(expected_type=Dimensions, allow_none=True)
    measureGroups = Typed(expected_type=MeasureGroups, allow_none=True)
    maps = Typed(expected_type=MeasureDimensionMaps, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('cacheSource', 'cacheFields', 'cacheHierarchies', 'kpis',
                    'tupleCache', 'calculatedItems', 'calculatedMembers', 'dimensions',
                    'measureGroups', 'maps',)

    def __init__(self,
                 invalid=None,
                 saveData=None,
                 refreshOnLoad=None,
                 optimizeMemory=None,
                 enableRefresh=None,
                 refreshedBy=None,
                 refreshedDate=None,
                 refreshedDateIso=None,
                 backgroundQuery=None,
                 missingItemsLimit=None,
                 createdVersion=None,
                 refreshedVersion=None,
                 minRefreshableVersion=None,
                 recordCount=None,
                 upgradeOnRefresh=None,
                 tupleCache=None,
                 supportSubquery=None,
                 supportAdvancedDrill=None,
                 cacheSource=None,
                 cacheFields=(),
                 cacheHierarchies=None,
                 kpis=None,
                 calculatedItems=None,
                 calculatedMembers=None,
                 dimensions=None,
                 measureGroups=None,
                 maps=None,
                 extLst=None,
                ):
        self.invalid = invalid
        self.saveData = saveData
        self.refreshOnLoad = refreshOnLoad
        self.optimizeMemory = optimizeMemory
        self.enableRefresh = enableRefresh
        self.refreshedBy = refreshedBy
        self.refreshedDate = refreshedDate
        self.refreshedDateIso = refreshedDateIso
        self.backgroundQuery = backgroundQuery
        self.missingItemsLimit = missingItemsLimit
        self.createdVersion = createdVersion
        self.refreshedVersion = refreshedVersion
        self.minRefreshableVersion = minRefreshableVersion
        self.recordCount = recordCount
        self.upgradeOnRefresh = upgradeOnRefresh
        self.tupleCache = tupleCache
        self.supportSubquery = supportSubquery
        self.supportAdvancedDrill = supportAdvancedDrill
        self.cacheSource = cacheSource
        self.cacheFields = cacheFields
        self.cacheHierarchies = cacheHierarchies
        self.kpis = kpis
        self.tupleCache = tupleCache
        self.calculatedItems = calculatedItems
        self.calculatedMembers = calculatedMembers
        self.dimensions = dimensions
        self.measureGroups = measureGroups
        self.maps = maps


    def to_tree(self):
        node = super(PivotCacheDefinition, self).to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node
