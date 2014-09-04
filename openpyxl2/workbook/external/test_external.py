# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest
from openpyxl2.tests.helper import compare_xml, get_xml

from openpyxl2.reader.workbook import read_rels
from openpyxl2.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK_RELS,
    PKG_REL_NS,
    REL_NS,
)
from openpyxl2.xml.functions import tostring


def test_read_external_ref(datadir):
    datadir.chdir()
    archive = ZipFile(BytesIO(), "w")
    with open("[Content_Types].xml") as src:
        archive.writestr(ARC_CONTENT_TYPES, src.read())
    with open("workbook.xml.rels") as src:
        archive.writestr(ARC_WORKBOOK_RELS, src.read())
    rels = read_rels(archive)
    for _, pth in rels:
        if pth['type'] == '%s/externalLink' % REL_NS:
            assert pth['path'] == 'xl/externalLinks/externalLink1.xml'


def test_read_external_link(datadir):
    from openpyxl2.workbook.external import parse_books
    datadir.chdir()
    with open("externalLink1.xml.rels") as src:
        xml = src.read()
    book = parse_books(xml)
    assert book.Id == 'rId1'


def test_read_external_ranges(datadir):
    from openpyxl2.workbook.external import parse_ranges
    datadir.chdir()
    with open("externalLink1.xml") as src:
        xml = src.read()
    names = tuple(parse_ranges(xml))
    assert names[0].name == 'B2range'
    assert names[0].refersTo == "='Sheet1'!$A$1:$A$10"


def test_dict_external_book():
    from openpyxl2.workbook.external import ExternalBook
    book = ExternalBook('rId1', "book1.xlsx")
    assert dict(book) == {'Id':'rId1', 'Target':'book1.xlsx',
                          'TargetMode':'External',
                          'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'}


def test_dict_external_range():
    from openpyxl2.workbook.external import ExternalRange
    rng = ExternalRange("something_special", "='Sheet1'!$A$1:$B$2")
    assert dict(rng) == {'name':'something_special', 'refersTo':"='Sheet1'!$A$1:$B$2"}


def test_write_external_link():
    from openpyxl2.workbook.external import ExternalRange
    from openpyxl2.workbook.external.writer import write_external_link
    link1 = ExternalRange('r1', 'over_there!$A$1:$B$2')
    link2 = ExternalRange('r2', 'somewhere_else!$C$10:$D$12')
    links = [link1, link2]
    el = write_external_link(links)
    xml = get_xml(el)
    expected = """
    <externalLink xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <externalBook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1">
        <definedNames>
          <definedName name="r1" refersTo="over_there!$A$1:$B$2"/>
          <definedName name="r2" refersTo="somewhere_else!$C$10:$D$12"/>
        </definedNames>
      </externalBook>
    </externalLink>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_external_book_rel():
    from openpyxl2.workbook.external import ExternalBook
    from openpyxl2.workbook.external.writer import write_external_book_rel
    book = ExternalBook("rId1", "book2.xlsx")
    rel = write_external_book_rel(book)
    xml = tostring(rel)
    expected = """
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath" Target="book2.xlsx" TargetMode="External"/>
</Relationships>

"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_read_archive(datadir):
    from openpyxl2.reader.workbook import read_rels
    from openpyxl2.workbook.external import detect_external_links
    datadir.chdir()
    archive = ZipFile("book1.xlsx")
    rels = read_rels(archive)
    books = detect_external_links(rels, archive)
    book = tuple(books)[0]
    assert book.Target == "book2.xlsx"

    expected = ["='Sheet1'!$A$1:$A$10", ]
    for link, exp in zip(book.links, expected):
        assert link.refersTo == exp


def test_load_workbook(datadir):
    datadir.chdir()
    from openpyxl2 import load_workbook
    wb = load_workbook('book1.xlsx')
    assert len(wb._external_links) == 1
