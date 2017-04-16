#Autogenerated schema
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Typed,
    DateTime,
    Bool,
    Float,
    String,
    Integer,
    Sequence,
)

from openpyxl2.descriptors.excel import HexBinary, ExtensionList
from openpyxl2.descriptors.nested import (
    NestedInteger,
    NestedBool,
)


class X(Serialisable):

    v = Integer(allow_none=True)

    def __init__(self,
                 v=0,
                ):
        self.v = v


class Tuple(Serialisable):

    fld = Integer()
    hier = Integer()
    item = Integer()

    def __init__(self,
                 fld=None,
                 hier=None,
                 item=None,
                ):
        self.fld = fld
        self.hier = hier
        self.item = item


class TupleList(Serialisable):

    c = Integer(allow_none=True)
    tpl = Typed(expected_type=Tuple, )

    __elements__ = ('tpl',)

    def __init__(self,
                 c=None,
                 tpl=None,
                ):
        self.c = c
        self.tpl = tpl


class SharedItem(Serialisable):

    pass


class Missing(SharedItem):

    tagname = "m"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=X)
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary()
    fc = HexBinary()
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=False,
                 un=False,
                 st=False,
                 b=False,
                ):
        self.tpls = tpls
        self.x = x
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b

class Number(SharedItem):

    tagname = "n"

    tpls = Sequence(expected_type=TupleList)
    x = Sequence(expected_type=X)
    v = Float()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary()
    fc = HexBinary()
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=(),
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=False,
                 un=False,
                 st=False,
                 b=False,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Error(SharedItem):

    tagname = "e"

    tpls = Typed(expected_type=TupleList, allow_none=True)
    x = Sequence(expected_type=X)
    v = String()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)
    _in = Integer(allow_none=True)
    bc = HexBinary()
    fc = HexBinary()
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 tpls=None,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=False,
                 un=False,
                 st=False,
                 b=False,
                ):
        self.tpls = tpls
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b


class Boolean(SharedItem):

    tagname = "b"

    x = Sequence(expected_type=X)
    v = Bool()
    u = Bool(allow_none=True)
    f = Bool(allow_none=True)
    c = String(allow_none=True)
    cp = Integer(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 x=(),
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                ):
        self.x = x
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp


class Text(SharedItem):

    tagname = "s"

    v = String(allow_none=True)
    u = Bool()
    f = Bool()
    c = String()
    cp = Integer()
    _in = Integer(allow_none=True)
    bc = HexBinary()
    fc = HexBinary()
    i = Bool(allow_none=True)
    un = Bool(allow_none=True)
    st = Bool(allow_none=True)
    b = Bool(allow_none=True)
    tpls = Typed(expected_type=TupleList, allow_none=True)
    x = NestedInteger(allow_none=True)

    __elements__ = ('tpls', 'x')

    def __init__(self,
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 _in=None,
                 bc=None,
                 fc=None,
                 i=None,
                 un=None,
                 st=None,
                 b=None,
                 tpls=None,
                 x=None,
                ):
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self._in = _in
        self.bc = bc
        self.fc = fc
        self.i = i
        self.un = un
        self.st = st
        self.b = b
        self.tpls = tpls
        self.x = x


class PivotDateTime(Serialisable):

    v = DateTime()
    u = Bool()
    f = Bool()
    c = String()
    cp = Integer()
    x = NestedInteger(allow_none=True)

    __elements__ = ('x',)

    def __init__(self,
                 v=None,
                 u=None,
                 f=None,
                 c=None,
                 cp=None,
                 x=None,
                ):
        self.v = v
        self.u = u
        self.f = f
        self.c = c
        self.cp = cp
        self.x = x


class Record(Serialisable):

    tagname = "r"

    # some elements are choice
    m = Typed(expected_type=Missing, allow_none=True)
    n = Typed(expected_type=Number, allow_none=True)
    b = NestedBool(allow_none=True)
    e = Typed(expected_type=Error, allow_none=True)
    s = Typed(expected_type=Text, allow_none=True)
    d = Typed(xpected_type=PivotDateTime, allow_none=True)
    x = NestedInteger(allow_none=True,)

    __elements__ = ('m', 'n', 'b', 'e', 's', 'd', 'x')

    def __init__(self,
                 m=None,
                 n=None,
                 b=None,
                 e=None,
                 s=None,
                 d=None,
                 x=None,
                ):
        self.m = m
        self.n = n
        self.b = b
        self.e = e
        self.s = s
        self.d = d
        self.x = x


class PivotCacheRecordList(Serialisable):

    count = Integer()
    r = Typed(expected_type=Record, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('r', )

    def __init__(self,
                 count=None,
                 r=None,
                 extLst=None,
                ):
        self.count = count
        self.r = r
        self.extLst = extLst
