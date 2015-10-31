from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl2.descriptors import (
    String,
    Set,
    NoneSet,
    Alias,
    Sequence,
)
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
    Id = String(allow_none=True)
    id = Alias('Id')


    def __init__(self,
                 type=None,
                 target=None,
                 targetMode=None,
                 id=None,
                 Id=None,
                 Type=None,
                 Target=None,
                 ):
        if type is not None:
            Type = "%s/%s" % (REL_NS, type)
        self.Type = Type
        if target is not None:
            Target = target
        self.Target = Target
        self.targetMode = targetMode
        if id is not None:
            Id = id
        self.Id = Id


class RelationshipList(Serialisable):

    tagname = "Relationships"

    Relationship = Sequence(expected_type=Relationship)


    def __init__(self, Relationship=()):
        self.Relationship = Relationship


    def append(self, value):
        values = self.Relationship[:]
        values.append(value)
        self.Relationship = values


    def __len__(self):
        return len(self.Relationship)


    def __bool__(self):
        return bool(self.Relationship)


    def to_tree(self):
        tree = Element("Relationships", xmlns=PKG_REL_NS)
        for idx, rel in enumerate(self.Relationship, 1):
            if not rel.id:
                rel.id = "rId{0}".format(idx)
            tree.append(rel.to_tree())

        return tree
