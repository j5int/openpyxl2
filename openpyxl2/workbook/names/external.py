from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import posixpath

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    String,
    Bool,
    Integer,
    NoneSet,
    Sequence,
)
from openpyxl2.descriptors.excel import Relation, ExtensionList
from openpyxl2.descriptors.nested import NestedText
from openpyxl2.descriptors.sequence import NestedSequence

from openpyxl2.packaging.relationship import Relationship, RelationshipList
from openpyxl2.xml.constants import (
    SHEET_MAIN_NS,
    REL_NS,
    EXTERNAL_LINK_NS,
)
from openpyxl2.xml.functions import (
    fromstring,
    safe_iterator,
    Element,
    SubElement,
)


"""Manage links to external Workbooks"""


class ExternalCell(Serialisable):

    r = String()
    t = NoneSet(values=(['b', 'd', 'n', 'e', 's', 'str', 'inlineStr']))
    vm = Integer(allow_none=True)
    v = NestedText(allow_none=True, expected_type=str)

    def __init__(self,
                 r=None,
                 t=None,
                 vm=None,
                 v=None,
                ):
        self.r = r
        self.t = t
        self.vm = vm
        self.v = v


class ExternalRow(Serialisable):

    r = Integer()
    cell = Typed(expected_type=ExternalCell, allow_none=True)

    __elements__ = ('cell',)

    def __init__(self,
                 r=None,
                 cell=None,
                ):
        self.r = r
        self.cell = cell


class ExternalSheetData(Serialisable):

    sheetId = Integer()
    refreshError = Bool(allow_none=True)
    row = Typed(expected_type=ExternalRow, allow_none=True)

    __elements__ = ('row',)

    def __init__(self,
                 sheetId=None,
                 refreshError=None,
                 row=None,
                ):
        self.sheetId = sheetId
        self.refreshError = refreshError
        self.row = row


class ExternalSheetDataSet(Serialisable):

    sheetData = Typed(expected_type=ExternalSheetData, )

    __elements__ = ('sheetData',)

    def __init__(self,
                 sheetData=None,
                ):
        self.sheetData = sheetData


class ExternalDefinedName(Serialisable):

    name = String()
    refersTo = String(allow_none=True)
    sheetId = Integer(allow_none=True)

    def __init__(self,
                 name=None,
                 refersTo=None,
                 sheetId=None,
                ):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


class ExternalDefinedNames(Serialisable):

    definedName = Sequence(expected_type=ExternalDefinedName, allow_none=True)

    __elements__ = ('definedName',)

    def __init__(self,
                 definedName=(),
                ):
        self.definedName = definedName


class ExternalSheetName(Serialisable):

    val = String()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ExternalSheetNames(Serialisable):

    sheetName = Typed(expected_type=ExternalSheetName, )

    __elements__ = ('sheetName',)

    def __init__(self,
                 sheetName=None,
                ):
        self.sheetName = sheetName


class ExternalBook(Serialisable):

    sheetNames = Typed(expected_type=ExternalSheetNames, allow_none=True)
    definedNames = Typed(expected_type=ExternalDefinedNames, allow_none=True)
    sheetDataSet = Typed(expected_type=ExternalSheetDataSet, allow_none=True)
    id = Relation()

    __elements__ = ('sheetNames', 'definedNames', 'sheetDataSet')

    def __init__(self,
                 sheetNames=None,
                 definedNames=None,
                 sheetDataSet=None,
                 id=None,
                ):
        self.sheetNames = sheetNames
        self.definedNames = definedNames
        self.sheetDataSet = sheetDataSet
        self.id = id


class ExternalLink(Serialisable):

    tagname = "externalLink"

    externalBook = Typed(expected_type=ExternalBook, allow_none=True)

    __elements__ = ('externalBook', )

    def __init__(self,
                 externalBook=None,
                 ddeLink=None,
                 oleLink=None,
                 extLst=None,
                ):
        self.externalBook = externalBook
        # ignore other items for the moment.


    def to_tree(self):
        node = super(ExternalLink, self).to_tree()
        node.set("xmlns", SHEET_MAIN_NS)
        return node


def parse_books(xml):
    tree = fromstring(xml)

    rels = RelationshipList.from_tree(tree)
    for r in rels.Relationship:
        return r


def parse_ranges(xml):
    tree = fromstring(xml)
    tree = ExternalLink.from_tree(tree)

    book = tree.externalBook
    if book is None:
        return

    return book.definedNames.definedName

from openpyxl2.packaging.relationship import get_rels_path, get_dependents


def detect_external_links(rels, archive):
    """
    Given a workbooks rels, return the externalLink and the path of the linked workbook
    """

    for r in rels:

        if r.Type == EXTERNAL_LINK_NS:
            book_path = get_rels_path(r.Target)
            book_xml = archive.read(book_path)
            Book = parse_books(book_xml)

            range_xml = archive.read(r.Target)
            Book.links = list(parse_ranges(range_xml))
            yield Book


def write_external_link(links):
    """Serialise links to ranges in a single external worbook"""
    book = ExternalBook(id="rId1", definedNames=ExternalDefinedNames())
    for l in links:
        book.definedNames.definedName.append(l)
    root = ExternalLink(externalBook=book)
    return root.to_tree()


def write_external_book_rel(book):
    """Serialise link to external file"""
    root = RelationshipList()
    root.append(book)
    return root.to_tree()
