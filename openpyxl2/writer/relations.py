from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl2.xml.functions import Element
from openpyxl2.xml.constants import (
    PKG_REL_NS,
)
from openpyxl2.packaging.relationship import Relationship


def write_rels(worksheet, comments_id=None, vba_controls_id=None):
    """Write relationships for the worksheet to xml."""
    root = Element('Relationships', xmlns=PKG_REL_NS)
    rels = worksheet._rels

    if worksheet._comment_count > 0:

        rel = Relationship(type="comments", id="comments",
                           target='/comments%s.xml' % comments_id)
        rels.append(rel)

        rel = Relationship(type="vmlDrawing", id="commentsvml",
                           target='/drawings/commentsDrawing%s.vml' % comments_id)
        rels.append(rel)

    if worksheet.vba_controls is not None:
        rel = Relationship("vmlDrawing", id=worksheet.vba_controls,
                           target='/drawings/vmlDrawing%s.vml' % vba_controls_id)
        rels.append(rel)

    for idx, rel in enumerate(rels, 1):
        if rel.id is None:
            rel.id = "rId{0}".format(idx)
        root.append(rel.to_tree())

    return root
