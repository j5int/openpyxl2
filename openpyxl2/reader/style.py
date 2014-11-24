from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Read shared style definitions"""

# package imports
from openpyxl2.compat import OrderedDict, zip
from openpyxl2.utils.indexed_list import IndexedList
from openpyxl2.utils.exceptions import MissingNumberFormat
from openpyxl2.styles import (
    Style,
    numbers,
    Font,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Protection,
    Alignment,
    borders,
)
from openpyxl2.styles.colors import COLOR_INDEX, Color
from openpyxl2.styles.proxy import StyleId
from openpyxl2.xml.functions import fromstring, safe_iterator, localname
from openpyxl2.xml.constants import SHEET_MAIN_NS
from copy import deepcopy


class SharedStylesParser(object):

    def __init__(self, xml_source):
        self.root = fromstring(xml_source)
        self.shared_styles = IndexedList()
        self.cell_styles = IndexedList()
        self.cond_styles = []
        self.style_prop = {}
        self.color_index = COLOR_INDEX
        self.font_list = IndexedList()
        self.fill_list = IndexedList()
        self.border_list = IndexedList()
        self.alignments = IndexedList()
        self.protections = IndexedList()

    def parse(self):
        self.parse_custom_num_formats()
        self.parse_color_index()
        self.style_prop['color_index'] = self.color_index
        self.font_list = IndexedList(self.parse_fonts())
        self.fill_list = IndexedList(self.parse_fills())
        self.border_list = IndexedList(self.parse_borders())
        self.parse_dxfs()
        self.parse_cell_styles()

    def parse_custom_num_formats(self):
        """Read in custom numeric formatting rules from the shared style table"""
        custom_formats = {}
        num_fmts = self.root.find('{%s}numFmts' % SHEET_MAIN_NS)
        if num_fmts is not None:
            num_fmt_nodes = safe_iterator(num_fmts, '{%s}numFmt' % SHEET_MAIN_NS)
            for num_fmt_node in num_fmt_nodes:
                fmt_id = int(num_fmt_node.get('numFmtId'))
                fmt_code = num_fmt_node.get('formatCode').lower()
                custom_formats[fmt_id] = fmt_code
        self.custom_num_formats = custom_formats

    def parse_color_index(self):
        """Read in the list of indexed colors"""
        colors = self.root.find('{%s}colors' % SHEET_MAIN_NS)
        if colors is not None:
            indexedColors = colors.find('{%s}indexedColors' % SHEET_MAIN_NS)
            if indexedColors is not None:
                color_nodes = safe_iterator(indexedColors, '{%s}rgbColor' % SHEET_MAIN_NS)
                self.color_index = IndexedList([node.get('rgb') for node in color_nodes])

    def parse_dxfs(self):
        """Read in the dxfs effects - used by conditional formatting."""
        dxf_list = []
        dxfs = self.root.find('{%s}dxfs' % SHEET_MAIN_NS)
        if dxfs is not None:
            nodes = dxfs.findall('{%s}dxf' % SHEET_MAIN_NS)
            for dxf in nodes:
                dxf_item = {}
                font_node = dxf.find('{%s}font' % SHEET_MAIN_NS)
                if font_node is not None:
                    dxf_item['font'] = self.parse_font(font_node)
                fill_node = dxf.find('{%s}fill' % SHEET_MAIN_NS)
                if fill_node is not None:
                    dxf_item['fill'] = self.parse_fill(fill_node)
                border_node = dxf.find('{%s}border' % SHEET_MAIN_NS)
                if border_node is not None:
                    dxf_item['border'] = self.parse_border(border_node)
                dxf_list.append(dxf_item)
        self.cond_styles = dxf_list

    def parse_fonts(self):
        """Read in the fonts"""
        fonts = self.root.find('{%s}fonts' % SHEET_MAIN_NS)
        if fonts is not None:
            for node in safe_iterator(fonts, '{%s}font' % SHEET_MAIN_NS):
                yield self.parse_font(node)

    def parse_font(self, font_node):
        """Read individual font"""
        font = {}
        for child in safe_iterator(font_node):
            if child is not font_node:
                tag = localname(child)
                font[tag] = child.get("val", True)
        underline = font_node.find('{%s}u' % SHEET_MAIN_NS)
        if underline is not None:
            font['u'] = underline.get('val', 'single')
        color = font_node.find('{%s}color' % SHEET_MAIN_NS)
        if color is not None:
            font['color'] = Color(**dict(color.attrib))
        return Font(**font)

    def parse_fills(self):
        """Read in the list of fills"""
        fills = self.root.find('{%s}fills' % SHEET_MAIN_NS)
        if fills is not None:
            for fill_node in safe_iterator(fills, '{%s}fill' % SHEET_MAIN_NS):
                yield self.parse_fill(fill_node)

    def parse_fill(self, fill_node):
        """Read individual fill"""
        pattern = fill_node.find('{%s}patternFill' % SHEET_MAIN_NS)
        gradient = fill_node.find('{%s}gradientFill' % SHEET_MAIN_NS)
        if pattern is not None:
            return self.parse_pattern_fill(pattern)
        if gradient is not None:
            return self.parse_gradient_fill(gradient)

    def parse_pattern_fill(self, node):
        fill = dict(node.attrib)
        for child in safe_iterator(node):
            if child is not node:
                tag = localname(child)
                fill[tag] = Color(**dict(child.attrib))
        return PatternFill(**fill)

    def parse_gradient_fill(self, node):
        fill = dict(node.attrib)
        color_nodes = safe_iterator(node, "{%s}color" % SHEET_MAIN_NS)
        fill['stop'] = tuple(Color(**dict(node.attrib)) for node in color_nodes)
        return GradientFill(**fill)

    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.find('{%s}borders' % SHEET_MAIN_NS)
        if borders is not None:
            for border_node in safe_iterator(borders, '{%s}border' % SHEET_MAIN_NS):
                yield self.parse_border(border_node)

    def parse_border(self, border_node):
        """Read individual border"""
        border = dict(border_node.attrib)

        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            node = border_node.find('{%s}%s' % (SHEET_MAIN_NS, side))
            if node is not None:
                bside = dict(node.attrib)
                color = node.find('{%s}color' % SHEET_MAIN_NS)
                if color is not None:
                    bside['color'] = Color(**dict(color.attrib))
                border[side] = Side(**bside)
        return Border(**border)


    def parse_named_styles(self):
        ns = OrderedDict()
        _styles = safe_iterator(self.root, "{%s}cellStyleXfs" % SHEET_MAIN_NS)
        _names = safe_iterator(self.root, "{%s}cellStyles" % SHEET_MAIN_NS)


    def parse_cell_styles(self):
        node = self.root.find('{%s}cellXfs' % SHEET_MAIN_NS)
        if node is not None:
            self.shared_styles, self.cell_styles = self._parse_cell_xfs(node)


    def _parse_xfs(self, node):
        """Read styles from the shared style table"""
        _styles  = []
        _style_ids = []

        builtin_formats = numbers.BUILTIN_FORMATS
        xfs = safe_iterator(node, '{%s}xf' % SHEET_MAIN_NS)
        for index, xf in enumerate(xfs):
            _style = {}

            alignmentId = protectionId = 0
            numFmtId = int(xf.get("numFmtId", 0))
            fontId = int(xf.get("fontId", 0))
            fillId = int(xf.get("fillId", 0))
            borderId = int(xf.get("borderId", 0))

            if numFmtId < 164:
                format_code = builtin_formats.get(numFmtId, 'General')
            else:
                fmt_code = self.custom_num_formats.get(numFmtId)
                if fmt_code is not None:
                    format_code = fmt_code
                else:
                    raise MissingNumberFormat('%s' % numFmtId)
            _style['number_format'] = format_code

            if bool_attrib(xf, 'applyAlignment'):
                al = xf.find('{%s}alignment' % SHEET_MAIN_NS)
                if al is not None:
                    alignment = Alignment(**al.attrib)
                    alignmentId = self.alignments.add(alignment)
                    _style['alignment'] = alignment

            if bool_attrib(xf, 'applyFont'):
                _style['font'] = self.font_list[fontId]

            if bool_attrib(xf, 'applyFill'):
                _style['fill'] = self.fill_list[fillId]

            if bool_attrib(xf, 'applyBorder'):
                _style['border'] = self.border_list[borderId]

            if bool_attrib(xf, 'applyProtection'):
                prot = xf.find('{%s}protection' % SHEET_MAIN_NS)
                if prot is not None:
                    protection = Protection(**prot.attrib)
                    protectionId = self.alignments.add(protection)
                    _style['protection'] = protection

            _styles.append(Style(**_style))
            _style_ids.append(StyleId(alignmentId, borderId, fillId, fontId, numFmtId, protectionId))

        return IndexedList(_styles), IndexedList(_style_ids)


def read_style_table(xml_source):
    p = SharedStylesParser(xml_source)
    p.parse()
    return p


def bool_attrib(element, attr):
    """
    Cast an XML attribute that should be a boolean to a Python equivalent
    None, 'f', '0' and 'false' all cast to False, everything else to true
    """
    value = element.get(attr)
    if not value or value in ("false", "f", "0"):
        return False
    return True
