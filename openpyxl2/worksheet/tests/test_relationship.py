from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import pytest
from openpyxl2.tests.helper import compare_xml
from openpyxl2.tests.schema import rel_src, XMLSchema
from openpyxl2.xml.functions import tostring


@pytest.fixture
def Relationship():
    from ..relationship import Relationship
    return Relationship


def test_ctor(Relationship):
    rel = Relationship("drawing", "drawings.xml", "external", "4")
    expected = """<Relationship id="4" target="drawings.xml" targetMode="external" type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" />
    """
    assert dict(rel) == {'id': '4', 'target': 'drawings.xml', 'targetMode':
                         'external', 'type':
                         'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'}
    xml = tostring(rel.to_tree())
    schema = XMLSchema(file=rel_src)
    schema.validate(rel.to_tree())

    diff = compare_xml(xml, expected)
    assert diff is None, diff
