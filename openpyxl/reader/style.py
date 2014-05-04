from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

"""Read shared style definitions"""

# package imports
from openpyxl.collections import IndexedList
from openpyxl.xml.functions import fromstring, safe_iterator, localname
from openpyxl.exceptions import MissingNumberFormat
from openpyxl.styles import (
    Style,
    NumberFormat,
    Font,
    PatternFill,
    GradientFill,
    Borders,
    Protection,
    Alignment
)
from openpyxl.styles.border import Border
from openpyxl.styles import borders
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.xml.constants import SHEET_MAIN_NS
from copy import deepcopy


class SharedStylesParser(object):

    def __init__(self, xml_source):
        self.root = fromstring(xml_source)
        self.style_prop = {'table': {},
                           'list': IndexedList()}
        self.color_index = COLOR_INDEX

    def parse(self):
        self.parse_custom_num_formats()
        self.parse_color_index()
        self.style_prop['color_index'] = self.color_index
        self.font_list = list(self.parse_fonts())
        self.fill_list = list(self.parse_fills())
        self.border_list = list(self.parse_borders())
        self.parse_dxfs()
        self.parse_cell_xfs()

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
                self.color_index = [node.get('rgb') for node in color_nodes]

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
        self.style_prop['dxf_list'] = dxf_list

    def _get_relevant_color(self, color):
        """Utility method for getting the color from different attributes"""
        value = None
        idx = color.get('indexed')
        theme = color.get('theme')
        tint = color.get('tint')
        rgb = color.get('rgb')
        if idx is not None:
            idx = int(idx)
            if idx < len(self.color_index):
                value = self.color_index[int(idx)]
        elif theme is not None:
            if tint is not None:
                value = 'theme:%s:%s' % (theme, tint)
            else:
                value = 'theme:%s:' % theme  # prefix color with theme
        elif rgb is not None:
            value = rgb
        return value

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
            index = self._get_relevant_color(color)
            if index is not None:
                font['color'] = Color(index)
            else:
                del font['color']
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
        fill = dict(node.items())
        for child in safe_iterator(node):
            if child is not node:
                tag = localname(child)
                color = self._get_relevant_color(child)
                if color is not None:
                    fill[tag] = Color(color)
        return PatternFill(**fill)

    def parse_gradient_fill(self, node):
        fill = dict(node.items())
        color_nodes = safe_iterator(node, "{%s}color" % SHEET_MAIN_NS)
        fill['stop'] = [Color(self._get_relevant_color(node)) for node in color_nodes]
        return GradientFill(**fill)

    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.find('{%s}borders' % SHEET_MAIN_NS)
        if borders is not None:
            for border_node in safe_iterator(borders, '{%s}border' % SHEET_MAIN_NS):
                yield self.parse_border(border_node)

    def parse_border(self, border_node):
        """Read individual border"""
        border = {}
        up = border_node.get('diagonalup') and borders.DIAGONAL_UP
        down = border_node.get('diagonaldown') and borders.DIAGONAL_DOWN
        if up and down:
            direction = borders.DIAGONAL_BOTH
        elif up:
            direction = up
        elif down:
            direction = down

        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            node = border_node.find('{%s}%s' % (SHEET_MAIN_NS, side))
            if node is not None:
                bside = {}
                if node.get('style') is not None:
                    bside['border_style'] = node.get('style')
                color = node.find('{%s}color' % SHEET_MAIN_NS)
                if color is not None:
                    index = self._get_relevant_color(color)
                    if index is not None:
                        bside['color'] = Color(index)
                    # Ignore 'auto'
                border[side] = Border(**bside)
        return Borders(**border)

    def parse_cell_xfs(self):
        """Read styles from the shared style table"""
        cell_xfs = self.root.find('{%s}cellXfs' % SHEET_MAIN_NS)
        styles_list = self.style_prop['list']

        if cell_xfs is None:  # can happen on bad OOXML writers (e.g. Gnumeric)
            return

        builtin_formats = NumberFormat._BUILTIN_FORMATS
        cell_xfs_nodes = safe_iterator(cell_xfs, '{%s}xf' % SHEET_MAIN_NS)
        for index, cell_xfs_node in enumerate(cell_xfs_nodes):
            _style = {}

            number_format_id = int(cell_xfs_node.get('numFmtId'))
            if number_format_id < 164:
                format_code = builtin_formats.get(number_format_id, 'General')
            else:
                fmt_code = self.custom_num_formats.get(number_format_id)
                if fmt_code is not None:
                    format_code = fmt_code
                else:
                    raise MissingNumberFormat('%s' % number_format_id)
            _style['number_format'] = NumberFormat(format_code=format_code)

            if bool(cell_xfs_node.get('applyAlignment')):
                alignment = {}
                al = cell_xfs_node.find('{%s}alignment' % SHEET_MAIN_NS)
                if al is not None:
                    for key in ('horizontal', 'vertical', 'indent'):
                        _value = al.get(key)
                        if _value is not None:
                            alignment[key] = _value
                    alignment['wrap_text'] = bool(al.get('wrapText'))
                    alignment['shrink_to_fit'] = bool(al.get('shrinkToFit'))
                    text_rotation = al.get('textRotation')
                    if text_rotation is not None:
                        alignment['text_rotation'] = int(text_rotation)
                    # ignore justifyLastLine option when horizontal = distributed
                _style['alignment'] = Alignment(**alignment)

            if bool(cell_xfs_node.get('applyFont')):
                _style['font'] = self.font_list[int(cell_xfs_node.get('fontId'))].copy()

            if bool(cell_xfs_node.get('applyFill')):
                _style['fill'] = self.fill_list[int(cell_xfs_node.get('fillId'))].copy()

            if bool(cell_xfs_node.get('applyBorder')):
                _style['borders'] = self.border_list[int(cell_xfs_node.get('borderId'))].copy()

            if bool(cell_xfs_node.get('applyProtection')):
                protection = {}
                prot = cell_xfs_node.find('{%s}protection' % SHEET_MAIN_NS)
                # Ignore if there are no protection sub-nodes
                if prot is not None:
                    protection['locked'] = bool(prot.get('locked'))
                    protection['hidden'] = bool(prot.get('hidden'))
                _style['protection'] = Protection(**protection)

            self.style_prop['table'][index] = styles_list.add(Style(**_style))


def read_style_table(xml_source):
    p = SharedStylesParser(xml_source)
    p.parse()
    return p.style_prop
