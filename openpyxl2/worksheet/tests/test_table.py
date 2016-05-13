from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl
import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def TableColumn():
    from ..table import TableColumn
    return TableColumn


class TestTableColumn:

    def test_ctor(self, TableColumn):
        col = TableColumn(id=1, name="Column1")
        xml = tostring(col.to_tree())
        expected = """
        <tableColumn id="1" name="Column1"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableColumn):
        src = """
        <tableColumn id="1" name="Column1"/>
        """
        node = fromstring(src)
        col = TableColumn.from_tree(node)
        assert col == TableColumn(id=1, name="Column1")


@pytest.fixture
def Table():
    from ..table import Table
    return Table


class TestTable:

    def test_ctor(self, Table, TableColumn):
        table = Table(id=1, displayName="A_Sample_Table", ref="A1:F10",
                      )
        table.tableColumns.append(TableColumn(id=1, name="Column1"))
        xml = tostring(table.to_tree())
        expected = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           displayName="A_Sample_Table" name="A_Sample_Table" id="1" ref="A1:F10">
        <tableColumns count="1">
          <tableColumn id="1" name="Column1" />
        </tableColumns>
        </table>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Table):
        src = """
        <table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            id="1" name="Table1" displayName="Table1" ref="A1:AA27">
        </table>
        """
        node = fromstring(src)
        table = Table.from_tree(node)
        assert table == Table(id=1, displayName="Table1", name="Table1",
                              ref="A1:AA27")


@pytest.fixture
def TableFormula():
    from ..table import TableFormula
    return TableFormula


class TestTableFormula:

    def test_ctor(self, TableFormula):
        formula = TableFormula()
        formula.text = "=A1*4"
        xml = tostring(formula.to_tree())
        expected = """
        <tableFormula>=A1*4</tableFormula>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableFormula):
        src = """
        <tableFormula>=A1*4</tableFormula>
        """
        node = fromstring(src)
        formula = TableFormula.from_tree(node)
        assert formula.text == "=A1*4"


@pytest.fixture
def TableStyleInfo():
    from ..table import TableStyleInfo
    return TableStyleInfo


class TestTableInfo:

    def test_ctor(self, TableStyleInfo):
        info = TableStyleInfo(name="TableStyleMedium12")
        xml = tostring(info.to_tree())
        expected = """
        <tableStyleInfo name="TableStyleMedium12" />
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, TableStyleInfo):
        src = """
        <tableStyleInfo name="TableStyleLight1" showRowStripes="1" />
        """
        node = fromstring(src)
        info = TableStyleInfo.from_tree(node)
        assert info == TableStyleInfo(name="TableStyleLight1", showRowStripes=True)


@pytest.fixture
def XMLColumnProps():
    from ..table import XMLColumnProps
    return XMLColumnProps


class TestXMLColumnPr:

    def test_ctor(self, XMLColumnProps):
        col = XMLColumnProps(mapId="1", xpath="/xml/foo/element", xmlDataType="string")
        xml = tostring(col.to_tree())
        expected = """
        <xmlColumnPr mapId="1" xpath="/xml/foo/element" xmlDataType="string"/>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, XMLColumnProps):
        src = """
        <xmlColumnPr mapId="1" xpath="/xml/foo/element" xmlDataType="string"/>
        """
        node = fromstring(src)
        col = XMLColumnProps.from_tree(node)
        assert col == XMLColumnProps(mapId="1", xpath="/xml/foo/element", xmlDataType="string")
