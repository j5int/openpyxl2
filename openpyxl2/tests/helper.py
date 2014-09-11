from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Python stdlib imports
from io import BytesIO
import os.path
from sys import version_info
from lxml.doctestcompare import LXMLOutputChecker, PARSE_XML

# package imports

# constants
DATADIR = os.path.abspath(os.path.join(os.path.dirname(__file__), 'data'))

def get_xml(xml_node):

    io = BytesIO()
    if version_info[0] >= 3 and version_info[1] >= 2:
        ElementTree(xml_node).write(io, encoding='UTF-8', xml_declaration=False)
        ret = str(io.getvalue(), 'utf-8')
        ret = ret.replace('utf-8', 'UTF-8', 1)
    else:
        ElementTree(xml_node).write(io, encoding='UTF-8')
        ret = io.getvalue()
    io.close()
    return ret.replace('\n', '')


def compare_xml(generated, expected):
    """Use doctest checking from lxml for comparing XML trees. Returns diff if the two are not the same"""
    checker = LXMLOutputChecker()

    class DummyDocTest():
        pass

    ob = DummyDocTest()
    ob.want = expected

    check = checker.check_output(expected, generated, PARSE_XML)
    if check is False:
        diff = checker.output_difference(ob, generated, PARSE_XML)
        return diff
