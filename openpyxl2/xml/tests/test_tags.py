# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
import pytest

from openpyxl2.xml.functions import start_tag, end_tag, tag, XMLGenerator
from openpyxl2.xml.constants import SHEET_MAIN_NS


@pytest.fixture(scope="class")
def doc():
    return BytesIO()

@pytest.fixture(scope="class")
def root(doc):
    return XMLGenerator(doc)


class TestSimpleTag:

    def test_start_tag(self, doc, root):
        start_tag(root, "start")
        assert doc.getvalue() == "<start>"

    def test_end_tag(self, doc, root):
        """"""
        end_tag(root, "blah")
        assert doc.getvalue() == "<start></blah>"


class TestTagBody:

    def test_start_tag(self, doc, root):
        start_tag(root, "start", body="just words")
        assert doc.getvalue() == "<start>just words"

    def test_end_tag(self, doc, root):
        end_tag(root, "end")
        assert doc.getvalue() == "<start>just words</end>"


def test_start_tag_attrs(doc, root):
    start_tag(root, "start", {'width':"10"})
    assert doc.getvalue() == """<start width="10">"""


def test_tag(doc, root):
    t = tag(root, "start", {'height':"10"}, "words")
    assert doc.getvalue() == """<start height="10">words</start>"""
