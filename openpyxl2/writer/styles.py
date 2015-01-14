from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Write the shared style table."""

# package imports

from openpyxl2.compat import safe_string
from openpyxl2.utils.indexed_list import IndexedList
from openpyxl2.xml.functions import (
    Element,
    SubElement,
    ConditionalElement,
    tostring,
    )
from openpyxl2.xml.constants import SHEET_MAIN_NS

from openpyxl2.styles import DEFAULTS
from openpyxl2.styles import numbers
from openpyxl2.styles.fills import GradientFill, PatternFill


class StyleWriter(object):

    def __init__(self, workbook):
        self.wb = workbook
        self._root = Element('styleSheet', {'xmlns': SHEET_MAIN_NS})

    @property
    def styles(self):
        return self.wb._cell_styles

    @property
    def fonts(self):
        return self.wb._fonts

    @property
    def fills(self):
        return self.wb._fills

    @property
    def borders(self):
        return self.wb._borders

    @property
    def number_formats(self):
        return self.wb._number_formats

    @property
    def alignments(self):
        return self.wb._alignments

    @property
    def protections(self):
        return self.wb._protections

    def write_table(self):
        self._write_number_formats()
        self._write_fonts()
        self._write_fills()
        self._write_borders()

        self._write_named_styles()
        self._write_cell_styles()
        self._write_style_names()
        self._write_conditional_styles()
        self._write_table_styles()

        return tostring(self._root)

    def _write_color(self, node, color, key='color'):
        """
        Convert colors encoded as RGB, theme, indexed, auto with tint
        """
        attrs = dict(color)
        SubElement(node, key, attrs)


    def _write_number_formats(self):
        node = SubElement(self._root, 'numFmts', count= "%d" % len(self.number_formats))
        for idx, nf in enumerate(self.number_formats, 164):
            SubElement(node, 'numFmt', {'numFmtId':'%d' % idx,
                                        'formatCode':'%s' % nf}
                       )


    def _write_fonts(self):
        fonts_node = SubElement(self._root, 'fonts', count="%d" % len(self.fonts))
        for font in self.fonts:
            fonts_node.append(font.serialise())


    def _write_fills(self):
        fills_node = SubElement(self._root, 'fills', count="%d" % len(self.fills))
        for fill in self.fills:
            fill_node = SubElement(fills_node, 'fill')
            fill_node.append(fill.serialise())

    def _write_borders(self):
        """Write the child elements for an individual border section"""
        borders_node = SubElement(self._root, 'borders', count="%d" % len(self.borders))
        for border in self.borders:
            border_node = SubElement(borders_node, 'border', dict(border))
            for tag, elem in border.children:
                side = SubElement(border_node, tag, dict(elem))
                if elem.color is not None:
                    self._write_color(side, elem.color)

    def _write_named_styles(self):
        cell_style_xfs = SubElement(self._root, 'cellStyleXfs', {'count':'1'})
        SubElement(cell_style_xfs, 'xf',
            {'numFmtId':"0", 'fontId':"0", 'fillId':"0", 'borderId':"0"})

    def _write_cell_styles(self):
        """ write styles combinations based on ids found in tables """
        # writing the cellXfs
        cell_xfs = SubElement(self._root, 'cellXfs',
                              count='%d' % len(self.styles))

        # default
        def _get_default_vals():
            return dict(numFmtId='0', fontId='0', fillId='0',
                        xfId='0', borderId='0')

        for st in self.styles:
            vals = _get_default_vals()

            if st.font != 0:
                vals['fontId'] = "%d" % (st.font)
                vals['applyFont'] = '1'

            if st.border != 0:
                vals['borderId'] = "%d" % (st.border)
                vals['applyBorder'] = '1'

            if st.fill != 0:
                vals['fillId'] =  "%d" % (st.fill)
                vals['applyFill'] = '1'

            if st.number_format != 0:
                vals['numFmtId'] = '%d' % st.number_format
                vals['applyNumberFormat'] = '1'

            node = SubElement(cell_xfs, 'xf', vals)

            if st.alignment != 0:
                node.set("applyProtection", '1')
                al = self.alignments[st.alignment]
                el = al.serialise()
                node.append(el)

            if st.protection != 0:
                node.set('applyProtection', '1')
                prot = self.protections[st.protection]
                el = prot.serialise()
                node.append(el)


    def _write_style_names(self):
        cell_styles = SubElement(self._root, 'cellStyles', {'count':'1'})
        SubElement(cell_styles, 'cellStyle',
            {'name':"Normal", 'xfId':"0", 'builtinId':"0"})

    def _write_conditional_styles(self):
        if self.wb.style_properties and 'dxf_list' in self.wb.style_properties:
            dxfs = SubElement(self._root, 'dxfs', {'count': str(len(self.wb.style_properties['dxf_list']))})
            for d in self.wb.style_properties['dxf_list']:
                dxf = SubElement(dxfs, 'dxf')
                if 'font' in d and d['font'] is not None:
                    font_node = SubElement(dxf, 'font')
                    if d['font'].color is not None:
                        self._write_color(font_node, d['font'].color)
                    ConditionalElement(font_node, 'b', d['font'].bold, 'val')
                    ConditionalElement(font_node, 'i', d['font'].italic, 'val')
                    ConditionalElement(font_node, 'u', d['font'].underline,
                                       {'val': d['font'].underline})
                    ConditionalElement(font_node, 'strike', d['font'].strikethrough)


                if 'fill' in d:
                    f = d['fill']
                    fill = SubElement(dxf, 'fill')
                    if f.fill_type:
                        node = SubElement(fill, 'patternFill', {'patternType': f.fill_type})
                    else:
                        node = SubElement(fill, 'patternFill')
                    if f.start_color != DEFAULTS.fill.start_color:
                        self._write_color(node, f.start_color, 'fgColor')

                    if f.end_color != DEFAULTS.fill.end_color:
                        self._write_color(node, f.end_color, 'bgColor')

                if 'border' in d:
                    borders = d['border']
                    border = SubElement(dxf, 'border')
                    # caution: respect this order
                    for side in ('left', 'right', 'top', 'bottom'):
                        obj = getattr(borders, side)
                        if obj.border_style is None or obj.border_style == 'none':
                            node = SubElement(border, side)
                        else:
                            node = SubElement(border, side, {'style': obj.border_style})
                            self._write_color(node, obj.color)
        else:
            dxfs = SubElement(self._root, 'dxfs', {'count': '0'})
        return dxfs

    def _write_table_styles(self):

        SubElement(self._root, 'tableStyles',
            {'count':'0', 'defaultTableStyle':'TableStyleMedium9',
            'defaultPivotStyle':'PivotStyleLight16'})
