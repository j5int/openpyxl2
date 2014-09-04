from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl2.descriptors import String, Strict
from openpyxl2.xml.constants import SHEET_MAIN_NS, REL_NS, PKG_REL_NS
from openpyxl2.xml.functions import fromstring, safe_iterator

"""Manage links to external Workbooks"""


class ExternalBook(Strict):

    """
    Map the relationship of one workbook to another
    """

    Id = String()
    Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath"
    TargetMode = "External"
    Target = String()

    def __init__(self, Id, Target, TargetMode=None, Type=None):
        self.Id = Id
        self.Target = Target
        links = []

    def __iter__(self):
        for attr in ('Id', 'Type', 'TargetMode', 'Target'):
            value = getattr(self, attr)
            yield attr, value


class ExternalRange(Strict):

    """
    Map external named ranges
    NB. the specification for these is different to named ranges within a workbook
    See 18.14.5
    """

    name = String()
    refersTo = String()
    sheetId = String(allow_none=True)

    def __init__(self, name, refersTo, sheetId=None):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


    def __iter__(self):
        for attr in ('name', 'refersTo', 'sheetId'):
            value = getattr(self, attr, None)
            if value is not None:
                yield attr, value


def parse_books(xml):
    tree = fromstring(xml)
    rels = tree.findall('{%s}Relationship' % PKG_REL_NS)
    for r in rels:
        yield ExternalBook(**r.attrib)


def parse_ranges(xml):
    tree = fromstring(xml)
    book = tree.find('{%s}externalBook' % SHEET_MAIN_NS)
    names = book.find('{%s}definedNames' % SHEET_MAIN_NS)
    for n in safe_iterator(names, '{%s}definedName' % SHEET_MAIN_NS):
        yield ExternalRange(**n.attrib)


def parse_external_links(rels, archive):
    for rId, d in rels:
        if d['type'] == EXTERNAL_LINK:
            pth = os.path.split()
            f_name = pth[-1]
            dir_name = pth[:-1]
            book_path = os.path.join(dir_name, "_rels", f_name + ".rels")
            xml = archive.read(book_path)
            Book = parse_books(xml)
            Book.links = list(parse_ranges(d['path']))
            yield Book
