from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Sequence
)
from openpyxl2.descriptors.excel import (
    Relation,
)

class ExternalReference(Serialisable):

    tagname = "externalReference"

    id = Relation()

    def __init__(self, id):
        self.id = id


class ExternalReferenceList(Serialisable):

    tagname = "externalReferences"

    externalReference = Sequence(expected_type=ExternalReference, )

    __elements__ = ('externalReference',)

    def __init__(self,
                 externalReference=None,
                ):
        self.externalReference = externalReference
