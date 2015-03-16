"""
Collection of utility primitives for charts.
"""

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    MinMax
)


class GapAmount(Serialisable):

    # needs to serialise to %
    val = MinMax(min=0, max=500)

    def __init__(self,
                 val=150,
                ):
        self.val = val


class Overlap(Serialisable):

    # needs to serialise to %

    val = MinMax(min=0, max=150)

    def __init__(self,
                 val=None,
                ):
        self.val = val
