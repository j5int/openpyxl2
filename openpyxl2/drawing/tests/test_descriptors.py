from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest

from openpyxl2.xml.functions import fromstring, tostring
from openpyxl2.tests.helper import compare_xml


def test_color_descriptor():
    from ..colors import ColorChoiceDescriptor

    class DummyStyle(object):

        value = ColorChoiceDescriptor('value')

    style = DummyStyle()
    style.value = "efefef"
    style.value.RGB == "efefef"
