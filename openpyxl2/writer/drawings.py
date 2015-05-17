# coding=UTF-8
from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl2.xml.functions import Element, SubElement, tostring
from openpyxl2.xml.constants import (
    DRAWING_NS,
    SHEET_DRAWING_NS,
    CHART_NS,
    REL_NS,
    CHART_DRAWING_NS,
    PKG_REL_NS
)
from openpyxl2.compat.strings import safe_string
from openpyxl2.chart.chart_drawing import (
    OneCellAnchor,
    TwoCellAnchor,
    AbsoluteAnchor,
)
from openpyxl2.chart.graphic import PictureFrame
from openpyxl2.chart.fill import Blip
from openpyxl2.utils.units import pixels_to_EMU


class DrawingWriter(object):
    """ one main drawing file per sheet """

    def __init__(self, sheet):
        self._sheet = sheet

    def write(self):
        """ write drawings for one sheet in one file """

        root = Element("{%s}wsDr" % SHEET_DRAWING_NS)

        for idx, chart in enumerate(self._sheet._charts):
            self._write_chart(root, chart, idx+1)

        for idx, img in enumerate(self._sheet._images):
            self._write_image(root, img, idx+1)

        return tostring(root)

    def _write_chart(self, node, chart, idx):
        """Add a chart"""
        drawing = chart.drawing

        # we only support absolute anchor atm (TODO: oneCellAnchor, twoCellAnchor
        x, y, w, h = drawing.get_emu_dimensions()
        anchor = SubElement(node, '{%s}absoluteAnchor' % SHEET_DRAWING_NS)
        SubElement(anchor, '{%s}pos' % SHEET_DRAWING_NS, {'x':str(x), 'y':str(y)})
        SubElement(anchor, '{%s}ext' % SHEET_DRAWING_NS, {'cx':str(w), 'cy':str(h)})

        # graph frame
        frame = SubElement(anchor, '{%s}graphicFrame' % SHEET_DRAWING_NS, {'macro':''})

        name = SubElement(frame, '{%s}nvGraphicFramePr' % SHEET_DRAWING_NS)
        SubElement(name, '{%s}cNvPr'% SHEET_DRAWING_NS, {'id':'%s' % (idx + 1), 'name':'Chart %s' % idx})
        SubElement(name, '{%s}cNvGraphicFramePr' % SHEET_DRAWING_NS)

        frm = SubElement(frame, '{%s}xfrm'  % SHEET_DRAWING_NS)
        # no transformation
        SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
        SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

        graph = SubElement(frame, '{%s}graphic' % DRAWING_NS)
        data = SubElement(graph, '{%s}graphicData' % DRAWING_NS, {'uri':CHART_NS})
        SubElement(data, '{%s}chart' % CHART_NS, {'{%s}id' % REL_NS:'rId%s' % idx })

        SubElement(anchor, '{%s}clientData' % SHEET_DRAWING_NS)
        return node

    def _write_anchor(self, node, drawing):
        x, y, w, h = drawing.get_emu_dimensions()

        if drawing.anchortype == "oneCell":
            anchor = OneCellAnchor()
            anchor._from.col = drawing.anchorcol
            anchor._from.row = drawing.anchorrow

        else:
            anchor = AbsoluteAnchor()
            anchor.pos.x = pixels_to_EMU(drawing.left)
            anchor.pos.y = pixels_to_EMU(drawing.top)

        anchor.ext.width = pixels_to_EMU(drawing._width)
        anchor.ext.height = pixels_to_EMU(drawing._height)

        node.append(anchor.to_tree())
        return anchor


    def _write_image(self, node, img, idx):
        anchor = self._write_anchor(node, img.drawing)

        pic = PictureFrame()
        pic.nvPicPr.cNvPr.desc = "Picture 1"
        pic.nvPicPr.cNvPr.id = 1
        pic.nvPicPr.cNvPicPr.picLocks
        pic.blipFill.blip = Blip()
        pic.blipFill.blip.embed = "rId1"
        pic.blipFill.blip.cstate = "print"

        pic.spPr.noFill = True
        pic.spPr.ln.prstDash = None
        pic.spPr.ln.w = 1
        pic.spPr.ln.noFill = True

        anchor.pic = pic
        node.append(anchor.to_tree())
        return anchor

    def write_rels(self, chart_id, image_id):

        root = Element("{%s}Relationships" % PKG_REL_NS)
        i = 0
        for i, chart in enumerate(self._sheet._charts):
            attrs = {'Id' : 'rId%s' % (i + 1),
                'Type' : '%s/chart' % REL_NS,
                'Target' : '../charts/chart%s.xml' % (chart_id + i) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        for j, img in enumerate(self._sheet._images):
            attrs = {'Id' : 'rId%s' % (i + j + 1),
                'Type' : '%s/image' % REL_NS,
                'Target' : '../media/image%s.png' % (image_id + j) }
            SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        return tostring(root)


class ShapeWriter(object):
    """ one file per shape """

    def __init__(self, shapes):

        self._shapes = shapes

    def write(self, shape_id):

        root = Element('{%s}userShapes' % CHART_NS)

        for shape in self._shapes:
            anchor = SubElement(root, '{%s}relSizeAnchor' % CHART_DRAWING_NS)

            xstart, ystart, xend, yend = shape.coordinates

            _from = SubElement(anchor, '{%s}from' % CHART_DRAWING_NS)
            SubElement(_from, '{%s}x' % CHART_DRAWING_NS).text = str(xstart)
            SubElement(_from, '{%s}y' % CHART_DRAWING_NS).text = str(ystart)

            _to = SubElement(anchor, '{%s}to' % CHART_DRAWING_NS)
            SubElement(_to, '{%s}x' % CHART_DRAWING_NS).text = str(xend)
            SubElement(_to, '{%s}y' % CHART_DRAWING_NS).text = str(yend)

            sp = SubElement(anchor, '{%s}sp' % CHART_DRAWING_NS, {'macro':'', 'textlink':''})
            nvspr = SubElement(sp, '{%s}nvSpPr' % CHART_DRAWING_NS)
            SubElement(nvspr, '{%s}cNvPr' % CHART_DRAWING_NS, {'id':str(shape_id), 'name':'shape %s' % shape_id})
            SubElement(nvspr, '{%s}cNvSpPr' % CHART_DRAWING_NS)

            sppr = SubElement(sp, '{%s}spPr' % CHART_DRAWING_NS)
            frm = SubElement(sppr, '{%s}xfrm' % DRAWING_NS,)
            # no transformation
            SubElement(frm, '{%s}off' % DRAWING_NS, {'x':'0', 'y':'0'})
            SubElement(frm, '{%s}ext' % DRAWING_NS, {'cx':'0', 'cy':'0'})

            prstgeom = SubElement(sppr, '{%s}prstGeom' % DRAWING_NS, {'prst':str(shape.style)})
            SubElement(prstgeom, '{%s}avLst' % DRAWING_NS)

            fill = SubElement(sppr, '{%s}solidFill' % DRAWING_NS, )
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.color})

            border = SubElement(sppr, '{%s}ln' % DRAWING_NS, {'w':str(shape._border_width)})
            sf = SubElement(border, '{%s}solidFill' % DRAWING_NS)
            SubElement(sf, '{%s}srgbClr' % DRAWING_NS, {'val':shape.border_color})

            self._write_style(sp)
            self._write_text(sp, shape)

            shape_id += 1

        return tostring(root)

    def _write_text(self, node, shape):
        """ write text in the shape """

        tx_body = SubElement(node, '{%s}txBody' % CHART_DRAWING_NS)
        SubElement(tx_body, '{%s}bodyPr' % DRAWING_NS, {'vertOverflow':'clip'})
        SubElement(tx_body, '{%s}lstStyle' % DRAWING_NS)
        p = SubElement(tx_body, '{%s}p' % DRAWING_NS)
        if shape.text:
            r = SubElement(p, '{%s}r' % DRAWING_NS)
            rpr = SubElement(r, '{%s}rPr' % DRAWING_NS, {'lang':'en-US'})
            fill = SubElement(rpr, '{%s}solidFill' % DRAWING_NS)
            SubElement(fill, '{%s}srgbClr' % DRAWING_NS, {'val':shape.text_color})

            SubElement(r, '{%s}t' % DRAWING_NS).text = shape.text
        else:
            SubElement(p, '{%s}endParaRPr' % DRAWING_NS, {'lang':'en-US'})

    def _write_style(self, node):
        """ write style theme """

        style = SubElement(node, '{%s}style' % CHART_DRAWING_NS)

        ln_ref = SubElement(style, '{%s}lnRef' % DRAWING_NS, {'idx':'2'})
        scheme_clr = SubElement(ln_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})
        SubElement(scheme_clr, '{%s}shade' % DRAWING_NS, {'val':'50000'})

        fill_ref = SubElement(style, '{%s}fillRef' % DRAWING_NS, {'idx':'1'})
        SubElement(fill_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        effect_ref = SubElement(style, '{%s}effectRef' % DRAWING_NS, {'idx':'0'})
        SubElement(effect_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'accent1'})

        font_ref = SubElement(style, '{%s}fontRef' % DRAWING_NS, {'idx':'minor'})
        SubElement(font_ref, '{%s}schemeClr' % DRAWING_NS, {'val':'lt1'})
