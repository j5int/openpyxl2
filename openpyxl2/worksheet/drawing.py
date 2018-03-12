from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors.excel import Relation


class Drawing(Serialisable):

    tagname = "drawing"

    id = Relation()

    def __init__(self, id=None):
        self.id = id
