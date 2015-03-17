from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    Integer,
    Set,
    MinMax,
    Bool,
)

from openpyxl2.descriptors.excel import ExtensionList


class Skip(Serialisable):

    val = Typed(expected_type=Integer(), )

    def __init__(self,
                 val=None,
                ):
        self.val = val


class LblOffset(Serialisable):

    # need to serialise to %
    val = MinMax(min=0, max=1000)

    def __init__(self,
                 val=None,
                ):
        self.val = val


class LblAlgn(Serialisable):

    val = Set(values=(['ctr', 'l', 'r']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class CatAx(Serialisable):

    auto = Bool(nested=True, allow_none=True)
    lblAlgn = Typed(expected_type=LblAlgn, allow_none=True)
    lblOffset = Typed(expected_type=LblOffset, allow_none=True)
    tickLblSkip = Typed(expected_type=Skip, allow_none=True)
    tickMarkSkip = Typed(expected_type=Skip, allow_none=True)
    noMultiLvlLbl = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('auto', 'lblAlgn', 'lblOffset', 'tickLblSkip', 'tickMarkSkip', 'noMultiLvlLbl', 'extLst')

    def __init__(self,
                 auto=None,
                 lblAlgn=None,
                 lblOffset=None,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 noMultiLvlLbl=None,
                 extLst=None,
                ):
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl
        self.extLst = extLst

