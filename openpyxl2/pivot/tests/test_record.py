from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl
import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml

@pytest.fixture
def Record():
    from ..record import Record
    return Record


class TestRecord:

    def test_ctor(self, Record):
        field = Record()
        xml = tostring(field.to_tree())
        expected = """
        <r />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Record):
        src = """
        <r />
        """
        node = fromstring(src)
        field = Record.from_tree(node)
        assert field == Record()


@pytest.fixture
def Error():
    from ..record import Error
    return Error


class TestError:

    def test_ctor(self, Error):
        error = Error()
        xml = tostring(error.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Error):
        src = """
        <root />
        """
        node = fromstring(src)
        error = Error.from_tree(node)
        assert error == Error()


@pytest.fixture
def Boolean():
    from ..record import Boolean
    return Boolean


class TestBoolean:

    def test_ctor(self, Boolean):
        boolean = Boolean()
        xml = tostring(boolean.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Boolean):
        src = """
        <root />
        """
        node = fromstring(src)
        boolean = Boolean.from_tree(node)
        assert boolean == Boolean()


@pytest.fixture
def Missing():
    from ..record import Missing
    return Missing


class TestMissing:

    def test_ctor(self, Missing):
        missing = Missing()
        xml = tostring(missing.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Missing):
        src = """
        <root />
        """
        node = fromstring(src)
        missing = Missing.from_tree(node)
        assert missing == Missing()


@pytest.fixture
def Number():
    from ..record import Number
    return Number


class TestNumber:

    def test_ctor(self, Number):
        number = Number()
        xml = tostring(number.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Number):
        src = """
        <root />
        """
        node = fromstring(src)
        number = Number.from_tree(node)
        assert number == Number()


@pytest.fixture
def Text():
    from ..record import Text
    return Text


class TestText:

    def test_ctor(self, Text):
        text = Text()
        xml = tostring(text.to_tree())
        expected = """
        <root />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Text):
        src = """
        <root />
        """
        node = fromstring(src)
        text = Text.from_tree(node)
        assert text == Text()