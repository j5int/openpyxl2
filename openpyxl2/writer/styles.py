from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

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
        self._write_number_format()
        self._write_font()

        fills_node = SubElement(self._root, 'fills', count=len(self.fills))
        for fill in self.fills:
            self._write_fill(fills_node, fill)

        borders_node = SubElement(self._root, 'borders', count=len(self.borders))
        for border in self.borders:
            self._write_border(borders_node, border)

        self._write_cell_style_xfs()
        self._write_cell_xfs(number_format_node, fonts_node, fills_node, borders_node)
        self._write_cell_style()
        self._write_dxfs()
        self._write_table_styles()

        return tostring(self._root)

    def _write_color(self, node, color, key='color'):
        """
        Convert colors encoded as RGB, theme, indexed, auto with tint
        """
        attrs = dict(color)
        SubElement(node, key, attrs)

    def _write_font(self):
        fonts_node = SubElement(self._root, 'fonts', count="%d" % len(self.fonts))
        for font in self.fonts:
            node = SubElement(fonts_node, "font")

        # if present vertAlign has to be at the start otherwise Excel has a problem
        ConditionalElement(node, "vertAlign", font.vertAlign, {'val':font.vertAlign})
        SubElement(node, 'sz', {'val':str(font.size)})
        self._write_color(node, font.color)
        SubElement(node, 'name', {'val':font.name})
        SubElement(node, 'family', {'val': '%d' % font.family})

        # boolean attrs
        for attr in ("b", "i", "outline", "shadow", "condense"):
            ConditionalElement(node, attr, getattr(font, attr))

        # Don't write the 'scheme' element because it appears to prevent
        # the font name from being applied in Excel.
        #SubElement(font_node, 'scheme', {'val':'minor'})

        ConditionalElement(node, "u", font.underline=='single')
        ConditionalElement(node, "charset", font.charset, {'val':str(font.charset)})

    def _write_pattern_fill(self, node, fill):
        node = SubElement(node, 'patternFill', {'patternType':
                                                fill.fill_type})
        if fill.start_color != DEFAULTS.fill.start_color:
            self._write_color(node, fill.start_color, 'fgColor')
        if fill.end_color != DEFAULTS.fill.end_color:
            self._write_color(node, fill.end_color, 'bgColor')

    def _write_gradient_fill(self, node, fill):
        node = SubElement(node, 'gradientFill', dict(fill))
        for idx, color in enumerate(fill.stop):
            stop = SubElement(node, "stop", {"position":safe_string(idx)})
            self._write_color(stop, color)

    def _write_fill(self, node, fill):
        fill_node = SubElement(node, 'fill')
        if isinstance(fill, PatternFill):
            self._write_pattern_fill(fill_node, fill)
        else:
            self._write_gradient_fill(fill_node, fill)

    def _write_border(self, node, border):
        """Write the child elements for an individual border section"""
        border_node = SubElement(node, 'border', dict(border))
        for tag, elem in border.children:
            side = SubElement(border_node, tag, dict(elem))
            if elem.color is not None:
                self._write_color(side, elem.color)

    def _write_cell_style_xfs(self):
        cell_style_xfs = SubElement(self._root, 'cellStyleXfs', {'count':'1'})
        SubElement(cell_style_xfs, 'xf',
            {'numFmtId':"0", 'fontId':"0", 'fillId':"0", 'borderId':"0"})

    def _write_cell_xfs(self, number_format_node, fonts_node, fills_node, borders_node):
        """ write styles combinations based on ids found in tables """
        # writing the cellXfs
        cell_xfs = SubElement(self._root, 'cellXfs',
            {'count':'%d' % len(self.styles)})

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

            if st.alignment != DEFAULTS.alignment:
                vals['applyAlignment'] = '1'

            if st.protection != DEFAULTS.protection:
                vals['applyProtection'] = '1'

            node = SubElement(cell_xfs, 'xf', vals)

            if st.alignment != DEFAULTS.alignment:
                self._write_alignment(node, st.alignment)

            if st.protection != DEFAULTS.protection:
                self._write_protection(node, st.protection)


    def _write_number_format(self):
        node = SubElement(self._root, 'numFmts', count= "%d" % len(self.number_formats))
        for idx, nf in enumerate(self.number_formats):
            SubElement(node, 'numFmt', {'numFmtId':'%d' % idx,
                                        'formatCode':'%s' % nf}
                       )


    def _write_alignment(self, node, alignment):
        alignment = self.wb._alignments[alignment]
        values = dict(alignment)
        if values.get('horizontal', 'general') == 'general':
            del values['horizontal']
        if values.get('vertical', 'bottom') == 'bottom':
            del values['vertical']
        SubElement(node, 'alignment', values)


    def _write_protection(self, node, protection):
        SubElement(node, 'protection', dict(protection))


    def _write_cell_style(self):
        cell_styles = SubElement(self._root, 'cellStyles', {'count':'1'})
        SubElement(cell_styles, 'cellStyle',
            {'name':"Normal", 'xfId':"0", 'builtinId':"0"})

    def _write_dxfs(self):
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
