from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from openpyxl2.xml.constants import CHART_NS

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors.excel import Relation


class ChartRelation(Serialisable):

    tagname = "chart"
    namespace = CHART_NS

    id = Relation()

    def __init__(self, id):
        self.id = id
