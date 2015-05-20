from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl2.xml.functions import Element, SubElement
from openpyxl2.xml.constants import (
    COMMENTS_NS,
    PKG_REL_NS,
    REL_NS,
    VML_NS,
)
from openpyxl2.packaging.relationship import Relationship


def write_rels(worksheet, drawing_id, comments_id, vba_controls_id):
    """Write relationships for the worksheet to xml."""
    root = Element('Relationships', xmlns=PKG_REL_NS)
    rels = worksheet.relationships

    if worksheet._charts or worksheet._images:
        rel = Relationship(type="drawing", id="rId1",
                           target='../drawings/drawing%s.xml' % drawing_id)
        rels.append(rel)

    if worksheet._comment_count > 0:

        rel = Relationship(type="comments", id="comments",
                           target='../comments%s.xml' % comments_id)
        rels.append(rel)

        rel = Relationship("type", target='../drawings/commentsDrawing%s.vml' % comments_id, id="commentsvml")
        rel.type = VML_NS
        rels.append(rel)

    if worksheet.vba_controls is not None:
        rel = Relationship("type", target='../drawings/vmlDrawing%s.vml' %
                           vba_controls_id, id=worksheet.vba_controls)
        rel.type = VML_NS
        rels.append(rel)

    for idx, rel in enumerate(rels, 1):
        rel.id = "rId{0}".format(idx)
        root.append(rel.to_tree())

    return root
