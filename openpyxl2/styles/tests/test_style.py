from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl2.xml.functions import tostring
from openpyxl2.tests.helper import compare_xml


@pytest.fixture
def StyleId():
    from .. style import StyleId
    return StyleId


def test_ctor(StyleId):
    style = StyleId()
    assert dict(style) == {'borderId': 0, 'fillId': 0, 'fontId': 0,
                           'numFmtId': 0, 'xfId': 0, 'alignmentId':0, 'protectionId':0}


def test_protection(StyleId):
    style = StyleId(protectionId=1)
    assert style.applyProtection is True


def test_alignment(StyleId):
    style = StyleId(alignmentId=1)
    assert style.applyAlignment is True


def test_serialise(StyleId):
    style = StyleId()
    xml = tostring(style.to_tree())
    expected = """
     <xf borderId="0" fillId="0" fontId="0" numFmtId="0" xfId="0" />
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
