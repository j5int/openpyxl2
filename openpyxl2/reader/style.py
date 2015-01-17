from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read shared style definitions"""

# package imports
from openpyxl2.compat import OrderedDict, zip
from openpyxl2.utils.indexed_list import IndexedList
from openpyxl2.utils.exceptions import MissingNumberFormat
from openpyxl2.styles import (
    Style,
    numbers,
    Font,
    Fill,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Protection,
    Alignment,
    borders,
)
from openpyxl2.formatting.conditional import ConditionalFormat
from openpyxl2.styles.colors import COLOR_INDEX, Color
from openpyxl2.styles.proxy import StyleId
from openpyxl2.styles.named_styles import NamedStyle
from openpyxl2.xml.functions import fromstring, safe_iterator, localname
from openpyxl2.xml.constants import SHEET_MAIN_NS, ARC_STYLE
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
        for node in self.root.findall("{%s}dxfs/{%s}dxf" % (SHEET_MAIN_NS, SHEET_MAIN_NS) ):
            self.cond_styles.append(ConditionalFormat.create(node))


    def parse_fonts(self):
        """Read in the fonts"""
        fonts = self.root.findall('{%s}fonts/{%s}font' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for node in fonts:
            yield Font.create(node)


    def parse_fills(self):
        """Read in the list of fills"""
        fills = self.root.findall('{%s}fills/{%s}fill' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for fill in fills:
            yield Fill.create(fill)

    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.findall('{%s}borders/{%s}border' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for border_node in borders:
            yield Border.create(border_node)


    def parse_named_styles(self):
        """
        Extract named styles
        """
        ns = []
        styles_node = self.root.find("{%s}cellStyleXfs" % SHEET_MAIN_NS)
        _styles, _ids = self._parse_xfs(styles_node)

        for _name, idx in self._parse_style_names():
            _id = _ids[idx]
            style = NamedStyle(name=_name)
            style.border = self.border_list[_id.border]
            style.fill = self.fill_list[_id.fill]
            style.font = self.font_list[_id.font]
            if _id.alignment:
                style.alignment = self.alignments[_id.alignment]
            if _id.protection:
                style.protection = self.protections[_id.protection]
            ns.append(style)
        self.named_styles = IndexedList(ns)


    def _parse_style_names(self):
        names_node = self.root.find("{%s}cellStyles" % SHEET_MAIN_NS)
        for _name in names_node:
            yield _name.get("name"), int(_name.get("xfId"))


    def parse_cell_styles(self):
        """
        Extract individual cell styles
        """
        node = self.root.find('{%s}cellXfs' % SHEET_MAIN_NS)
        if node is not None:
            self.shared_styles, self.cell_styles = self._parse_xfs(node)


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


def read_style_table(archive):
    if ARC_STYLE in archive.namelist():
        xml_source = archive.read(ARC_STYLE)
    else:
        return
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
