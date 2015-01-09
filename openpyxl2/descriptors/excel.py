from __future__ import absolute_import
#copyright openpyxl 2010-2015

"""
Excel specific descriptors
"""

from openpyxl2.compat import basestring
from . import MatchPattern, MinMax, Integer


class HexBinary(MatchPattern):

    pattern = "[0-9a-fA-F]+$"


class UniversalMeasure(MatchPattern):

    pattern = "[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)"


class TextPoint(MinMax):
    """
    Size in hundredths of points.
    In theory other units of measurement can be used but these are unbounded
    """
    expected_type = int

    min = -400000
    max = 400000


class Coordinate(MinMax):

    """union of unqualified coordinate and universal measure types
    see worksheet properties for universal measure
    """

    min= -27273042329600
    max= 27273042316900


class Percentage(MatchPattern):

    pattern = "((100)|([0-9][0-9]?))(\.[0-9][0-9]?)?%"
