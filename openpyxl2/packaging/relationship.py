from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors import String, Set, NoneSet, Alias
from openpyxl2.descriptors.serialisable import Serialisable

from openpyxl2.xml.constants import REL_NS, PKG_REL_NS
from openpyxl2.xml.functions import Element, SubElement, tostring


class Relationship(Serialisable):
    """Represents many kinds of relationships."""
    # TODO: Use this object for workbook relationships as well as
    # worksheet relationships

    tagname = "Relationship"

    Type = String()
    type = Alias('Type')
    Target = String()
    target = Alias('Target')
    TargetMode = String(allow_none=True)
    targetMode = Alias('TargetMode')
    Id = String()
    id = Alias('Id')


    def __init__(self, type, target=None, targetMode=None, id=None):
        self.type = "%s/%s" % (REL_NS, type)
        self.target = target
        self.targetMode = targetMode
        self.id = id
