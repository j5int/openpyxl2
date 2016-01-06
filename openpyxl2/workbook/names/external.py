from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import os

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import String
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


class ExternalRange(Serialisable):

    """
    Map external named ranges
    NB. the specification for these is different to named ranges within a workbook
    See 18.14.5
    """

    name = String()
    refersTo = String(allow_none=True)
    sheetId = String(allow_none=True)

    def __init__(self, name, refersTo=None, sheetId=None):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


def parse_books(xml):
    tree = fromstring(xml)

    rels = RelationshipList.from_tree(tree)
    for r in rels.Relationship:
        return r


def parse_ranges(xml):
    tree = fromstring(xml)
    book = tree.find('{%s}externalBook' % SHEET_MAIN_NS)
    if book is None:
        return
    names = book.find('{%s}definedNames' % SHEET_MAIN_NS)
    for n in safe_iterator(names, '{%s}definedName' % SHEET_MAIN_NS):
        yield ExternalRange(**n.attrib)


def detect_external_links(rels, archive):
    for rId, d in rels:
        if d['type'] == EXTERNAL_LINK_NS:
            pth = os.path.split(d['path'])
            f_name = pth[-1]
            dir_name = "/".join(pth[:-1])
            book_path = "{0}/_rels/{1}.rels".format (dir_name, f_name)
            book_xml = archive.read(book_path)
            Book = parse_books(book_xml)

            range_xml = archive.read(d['path'])
            Book.links = list(parse_ranges(range_xml))
            yield Book


def write_external_link(links):
    """Serialise links to ranges in a single external worbook"""
    root = Element("{%s}externalLink" % SHEET_MAIN_NS)
    book =  SubElement(root, "{%s}externalBook" % SHEET_MAIN_NS, {'{%s}id' % REL_NS:'rId1'})
    external_ranges = SubElement(book, "{%s}definedNames" % SHEET_MAIN_NS)
    for l in links:
        external_ranges.append(Element("{%s}definedName" % SHEET_MAIN_NS, dict(l)))
    return root


def write_external_book_rel(book):
    """Serialise link to external file"""
    root = RelationshipList()
    root.append(book)
    return root.to_tree()
