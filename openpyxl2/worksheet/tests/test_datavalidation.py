from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from io import BytesIO

import pytest

# package imports
from .. datavalidation import (
    DataValidation,
    ValidationType
    )

from openpyxl2.workbook import Workbook
from openpyxl2.tests.helper import get_xml, compare_xml

# There are already unit-tests in test_cell.py that test out the
# coordinate_from_string method.  This should be the only way the
# collapse_cell_addresses method can throw, so we don't bother using invalid
# cell coordinates in the test-data here.
COLLAPSE_TEST_DATA = [
    (["A1"], "A1"),
    (["A1", "B1"], "A1 B1"),
    (["A1", "A2", "A3", "A4", "B1", "B2", "B3", "B4"], "A1:A4 B1:B4"),
    (["A2", "A4", "A3", "A1", "A5"], "A1:A5"),
]
@pytest.mark.parametrize("cells, expected",
                         COLLAPSE_TEST_DATA)
def test_collapse_cell_addresses(cells, expected):
    from .. datavalidation import collapse_cell_addresses
    assert collapse_cell_addresses(cells) == expected


def test_list_validation():
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Fish"')
    assert dv.formula1, '"Dog,Cat == Fish"'
    dv_dict = dict(dv)
    assert dv_dict['type'] == 'list'
    assert dv_dict['allowBlank'] == '0'
    assert dv_dict['showErrorMessage'] == '1'
    assert dv_dict['showInputMessage'] == '1'


def test_error_message():
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Fish"')
    dv.set_error_message('You done bad')
    dv_dict = dict(dv)
    assert dv_dict['errorTitle'] == 'Validation Error'
    assert dv_dict['error'] == 'You done bad'


def test_prompt_message():
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Fish"')
    dv.set_prompt_message('Please enter a value')
    dv_dict = dict(dv)
    assert dv_dict['promptTitle'] == 'Validation Prompt'
    assert dv_dict['prompt'] == 'Please enter a value'


def test_writer_validation():
    from .. datavalidation import writer
    wb = Workbook()
    ws = wb.active
    dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Fish"')
    dv.add_cell(ws['A1'])

    xml = get_xml(writer(dv))
    expected = """
    <dataValidation allowBlank="0" showErrorMessage="1" showInputMessage="1" sqref="A1" type="list">
      <formula1>&quot;Dog,Cat,Fish&quot;</formula1>
      <formula2>None</formula2>
    </dataValidation>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
