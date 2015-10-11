from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl
import pytest

from openpyxl2.xml.functions import fromstring, tostring, Element
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def Sequence():
    from ..sequence import Sequence

    return Sequence


@pytest.fixture
def Dummy(Sequence):

    class Dummy:

        value = Sequence(expected_type=int, name="value")

    return Dummy


class TestSequence:

    @pytest.mark.parametrize("value", [list(), tuple()])
    def test_valid_ctor(self, Dummy, value):
        dummy = Dummy()
        dummy.value = value
        assert dummy.value == list(value)

    @pytest.mark.parametrize("value", ["", b"", dict(), 1, None])
    def test_invalid_container(self, Dummy, value):
        dummy = Dummy()
        with pytest.raises(TypeError):
            dummy.value = value


class TestPrimitive:

    def test_to_tree(self, Dummy):

        dummy = Dummy()
        dummy.value = [1, '2', 3]

        root = Element("root")
        for node in Dummy.value.to_tree(dummy.value, "el"):
            root.append(node)

        xml = tostring(root)
        expected = """
        <root>
          <el>1</el>
          <el>2</el>
          <el>3</el>
        </root>
        """
        diff = compare_xml(xml, expected)
        assert diff is None, diff


    def test_from_xml(self, Dummy):
        src = """
        <root>
          <el>1</el>
          <el>2</el>
          <el>3</el>
        </root>
        """
        node = fromstring(src)

        dummy = Dummy()
        desc = Dummy.value
        dummy.value = list(desc.from_tree(node))
        assert dummy.value == [1, 2, 3]
