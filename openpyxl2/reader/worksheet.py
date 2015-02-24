from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Reader for a single worksheet."""
from io import BytesIO

# compatibility imports
from openpyxl2.xml.functions import iterparse

# package imports
from openpyxl2.cell import Cell
from openpyxl2.worksheet import Worksheet, ColumnDimension, RowDimension
from openpyxl2.worksheet.page import PageMargins, PrintOptions, PageSetup
from openpyxl2.worksheet.protection import SheetProtection
from openpyxl2.worksheet.views import SheetView
from openpyxl2.xml.constants import SHEET_MAIN_NS, REL_NS
from openpyxl2.xml.functions import safe_iterator
from openpyxl2.styles import Color
from openpyxl2.formatting import ConditionalFormatting
from openpyxl2.worksheet.properties import parse_sheetPr
from openpyxl2.utils import (
    coordinate_from_string,
    get_column_letter,
    column_index_from_string,
    coordinate_to_tuple,
    )


def _get_xml_iter(xml_source):
    """
    Possible inputs: strings, bytes, members of zipfile, temporary file
    Always return a file like object
    """
    if not hasattr(xml_source, 'read'):
        try:
            xml_source = xml_source.encode("utf-8")
        except (AttributeError, UnicodeDecodeError):
            pass
        return BytesIO(xml_source)
    else:
        try:
            xml_source.seek(0)
        except:
            pass
        return xml_source


def _cast_number(value):
    "Convert numbers as string to an int or float"
    try:
        return int(value)
    except ValueError:
        return float(value)



class WorkSheetParser(object):

    COL_TAG = '{%s}col' % SHEET_MAIN_NS
    ROW_TAG = '{%s}row' % SHEET_MAIN_NS
    CELL_TAG = '{%s}c' % SHEET_MAIN_NS
    VALUE_TAG = '{%s}v' % SHEET_MAIN_NS
    FORMULA_TAG = '{%s}f' % SHEET_MAIN_NS
    MERGE_TAG = '{%s}mergeCell' % SHEET_MAIN_NS
    INLINE_STRING = "{%s}is/{%s}t" % (SHEET_MAIN_NS, SHEET_MAIN_NS)
    INLINE_RICHTEXT = "{%s}is/{%s}r/{%s}t" % (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS)

    def __init__(self, wb, title, xml_source, shared_strings):
        self.ws = wb.create_sheet(title=title)
        self.source = xml_source
        self.shared_strings = shared_strings
        self.guess_types = wb._guess_types
        self.data_only = wb.data_only
        self.styles = [v._asdict() for v in self.ws.parent._cell_styles]

    def parse(self):
        dispatcher = {
            '{%s}mergeCells' % SHEET_MAIN_NS: self.parse_merge,
            '{%s}col' % SHEET_MAIN_NS: self.parse_column_dimensions,
            '{%s}row' % SHEET_MAIN_NS: self.parse_row_dimensions,
            '{%s}printOptions' % SHEET_MAIN_NS: self.parse_print_options,
            '{%s}pageMargins' % SHEET_MAIN_NS: self.parse_margins,
            '{%s}pageSetup' % SHEET_MAIN_NS: self.parse_page_setup,
            '{%s}headerFooter' % SHEET_MAIN_NS: self.parse_header_footer,
            '{%s}conditionalFormatting' % SHEET_MAIN_NS: self.parser_conditional_formatting,
            '{%s}autoFilter' % SHEET_MAIN_NS: self.parse_auto_filter,
            '{%s}sheetProtection' % SHEET_MAIN_NS: self.parse_sheet_protection,
            '{%s}dataValidations' % SHEET_MAIN_NS: self.parse_data_validation,
            '{%s}sheetPr' % SHEET_MAIN_NS: self.parse_properties,
            '{%s}legacyDrawing' % SHEET_MAIN_NS: self.parse_legacy_drawing,
            '{%s}sheetViews' % SHEET_MAIN_NS: self.parse_sheet_views,
                      }
        tags = dispatcher.keys()
        stream = _get_xml_iter(self.source)
        it = iterparse(stream, tag=tags)

        for _, element in it:
            tag_name = element.tag
            if tag_name in dispatcher:
                dispatcher[tag_name](element)
                element.clear()

        self.ws._current_row = self.ws.max_row

        # Handle parsed conditional formatting rules together.
        if len(self.ws.conditional_formatting.parse_rules):
            self.ws.conditional_formatting.update(self.ws.conditional_formatting.parse_rules)

    def parse_cell(self, element):
        value = element.find(self.VALUE_TAG)
        if value is not None:
            value = value.text
        formula = element.find(self.FORMULA_TAG)
        data_type = element.get('t', 'n')
        coordinate = element.get('r')
        style_id = element.get('s')

        # assign formula to cell value unless only the data is desired
        if formula is not None and not self.data_only:
            data_type = 'f'
            if formula.text:
                value = "=" + formula.text
            else:
                value = "="
            formula_type = formula.get('t')
            if formula_type:
                self.ws.formula_attributes[coordinate] = {'t': formula_type}
                si = formula.get('si')  # Shared group index for shared formulas
                if si:
                    self.ws.formula_attributes[coordinate]['si'] = si
                ref = formula.get('ref')  # Range for shared formulas
                if ref:
                    self.ws.formula_attributes[coordinate]['ref'] = ref


        style = {}
        if style_id is not None:
            style_id = int(style_id)
            style = self.styles[style_id]

        row, column = coordinate_to_tuple(coordinate)
        cell = Cell(self.ws, row=row, col_idx=column, **style)
        self.ws._cells[(row, column)] = cell

        if value is not None:
            if data_type == 'n':
                value = _cast_number(value)
            elif data_type == 'b':
                value = bool(int(value))
            elif data_type == 's':
                value = self.shared_strings[int(value)]
            elif data_type == 'str':
                data_type = 's'

        else:
            if data_type == 'inlineStr':
                data_type = 's'
                child = element.find(self.INLINE_STRING)
                if child is None:
                    child = element.find(self.INLINE_RICHTEXT)
                if child is not None:
                    value = child.text

        if self.guess_types or value is None:
            cell.value = value
        else:
            cell._value=value
            cell.data_type=data_type


    def parse_merge(self, element):
        for mergeCell in safe_iterator(element, ('{%s}mergeCell' % SHEET_MAIN_NS)):
            self.ws.merge_cells(mergeCell.get('ref'))

    def parse_column_dimensions(self, col):
        attrs = dict(col.attrib)
        column = get_column_letter(int(attrs['min']))
        attrs['index'] = column
        dim = ColumnDimension(self.ws, **attrs)
        self.ws.column_dimensions[column] = dim


    def parse_row_dimensions(self, row):
        attrs = dict(row.attrib)
        if set(attrs) - set(['r', 'span']):
            attrs['worksheet'] = self.ws
            dim = RowDimension(**attrs)
            self.ws.row_dimensions[dim.index] = dim

        for cell in safe_iterator(row, self.CELL_TAG):
            self.parse_cell(cell)


    def parse_print_options(self, element):
        self.ws.print_options = PrintOptions(**element.attrib)

    def parse_margins(self, element):
        self.page_margins = PageMargins(**element.attrib)

    def parse_page_setup(self, element):
        id_key = '{%s}id' % REL_NS
        if id_key in element.attrib.keys():
            element.attrib['id'] = element.attrib.pop(id_key)

        self.ws.page_setup = PageSetup(**element.attrib)

    def parse_header_footer(self, element):
        oddHeader = element.find('{%s}oddHeader' % SHEET_MAIN_NS)
        if oddHeader is not None and oddHeader.text is not None:
            self.ws.header_footer.setHeader(oddHeader.text)
        oddFooter = element.find('{%s}oddFooter' % SHEET_MAIN_NS)
        if oddFooter is not None and oddFooter.text is not None:
            self.ws.header_footer.setFooter(oddFooter.text)

    def parser_conditional_formatting(self, element):
        range_string = element.get('sqref')
        cfRules = element.findall('{%s}cfRule' % SHEET_MAIN_NS)
        if range_string not in self.ws.conditional_formatting.parse_rules:
            self.ws.conditional_formatting.parse_rules[range_string] = []
        for cfRule in cfRules:
            if not cfRule.get('type') or cfRule.get('type') == 'dataBar':
                # dataBar conditional formatting isn't supported, as it relies on the complex <extLst> tag
                continue
            rule = {'type': cfRule.get('type')}
            for attr in ConditionalFormatting.rule_attributes:
                if cfRule.get(attr) is not None:
                    if attr == 'priority':
                        rule[attr] = int(cfRule.get(attr))
                    else:
                        rule[attr] = cfRule.get(attr)

            formula = cfRule.findall('{%s}formula' % SHEET_MAIN_NS)
            for f in formula:
                if 'formula' not in rule:
                    rule['formula'] = []
                rule['formula'].append(f.text)

            colorScale = cfRule.find('{%s}colorScale' % SHEET_MAIN_NS)
            if colorScale is not None:
                rule['colorScale'] = {'cfvo': [], 'color': []}
                cfvoNodes = colorScale.findall('{%s}cfvo' % SHEET_MAIN_NS)
                for node in cfvoNodes:
                    cfvo = {}
                    if node.get('type') is not None:
                        cfvo['type'] = node.get('type')
                    if node.get('val') is not None:
                        cfvo['val'] = node.get('val')
                    rule['colorScale']['cfvo'].append(cfvo)
                colorNodes = colorScale.findall('{%s}color' % SHEET_MAIN_NS)
                for color in colorNodes:
                    attrs = dict(color.items())
                    color = Color(**attrs)
                    rule['colorScale']['color'].append(color)

            iconSet = cfRule.find('{%s}iconSet' % SHEET_MAIN_NS)
            if iconSet is not None:
                rule['iconSet'] = {'cfvo': []}
                for iconAttr in ConditionalFormatting.icon_attributes:
                    if iconSet.get(iconAttr) is not None:
                        rule['iconSet'][iconAttr] = iconSet.get(iconAttr)
                cfvoNodes = iconSet.findall('{%s}cfvo' % SHEET_MAIN_NS)
                for node in cfvoNodes:
                    cfvo = {}
                    if node.get('type') is not None:
                        cfvo['type'] = node.get('type')
                    if node.get('val') is not None:
                        cfvo['val'] = node.get('val')
                    rule['iconSet']['cfvo'].append(cfvo)

            self.ws.conditional_formatting.parse_rules[range_string].append(rule)

    def parse_auto_filter(self, element):
        self.ws.auto_filter.ref = element.get("ref")
        for fc in safe_iterator(element, '{%s}filterColumn' % SHEET_MAIN_NS):
            filters = fc.find('{%s}filters' % SHEET_MAIN_NS)
            if filters is None:
                continue
            vals = [f.get("val") for f in safe_iterator(filters, '{%s}filter' % SHEET_MAIN_NS)]
            blank = filters.get("blank")
            self.ws.auto_filter.add_filter_column(fc.get("colId"), vals, blank=blank)
        for sc in safe_iterator(element, '{%s}sortCondition' % SHEET_MAIN_NS):
            self.ws.auto_filter.add_sort_condition(sc.get("ref"), sc.get("descending"))

    def parse_sheet_protection(self, element):
        values = element.attrib
        self.ws.protection = SheetProtection(**values)
        password = values.get("password")
        if password is not None:
            self.ws.protection.set_password(password, True)

    def parse_data_validation(self, element):
        from openpyxl2.worksheet.datavalidation import parser
        for tag in safe_iterator(element, "{%s}dataValidation" % SHEET_MAIN_NS):
            dv = parser(tag)
            self.ws._data_validations.append(dv)


    def parse_properties(self, element):
        self.ws.sheet_properties = parse_sheetPr(element)


    def parse_legacy_drawing(self, element):
        self.ws.vba_controls = element.get("r:id")


    def parse_sheet_views(self, element):
        for el in element.findall("{%s}sheetView" % SHEET_MAIN_NS):
            # according to the specification the last view wins
            pass
        self.ws.sheet_view = SheetView.from_tree(el)


def fast_parse(xml_source, parent, sheet_title, shared_strings):
    parser = WorkSheetParser(parent, sheet_title, xml_source, shared_strings)
    parser.parse()
    return parser.ws
