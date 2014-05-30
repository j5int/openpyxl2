from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

# Experimental writer of worksheet data using lxml incremental API

from lxml.etree import xmlfile, Element, SubElement

from openpyxl2.compat import (
    iterkeys,
    itervalues,
    safe_string,
    iteritems
)
from openpyxl2.cell import (
    column_index_from_string,
    coordinate_from_string
)
from openpyxl2.xml.constants import PKG_REL_NS

from .worksheet import row_sort


def write_cols(xf, worksheet, style_table=None):
    """Write worksheet columns to xml.

    style_table is ignored but required
    for compatibility with the dumped worksheet <cols> may never be empty -
    spec says must contain at least one child
    """
    cols = []
    for label, dimension in iteritems(worksheet.column_dimensions):
        dimension.style = worksheet._styles.get(label)
        col_def = dict(dimension)
        if col_def == {}:
            continue
        idx = column_index_from_string(label)
        cols.append((idx, col_def))

    if not cols:
        return

    with xf.element('cols'):
        for idx, col_def in sorted(cols):
            v = "%d" % idx
            cmin = col_def.get('min') or v
            cmax = col_def.get('max') or v
            col_def.update({'min': cmin, 'max': cmax})
            c = Element('col', col_def)
            xf.write(c)


def write_worksheet_data(xf, worksheet, string_table, style_table=None):
    """Write worksheet data to xml."""

    # Ensure a blank cell exists if it has a style
    for styleCoord in iterkeys(worksheet._styles):
        if isinstance(styleCoord, str) and COORD_RE.search(styleCoord):
            worksheet.cell(styleCoord)

    # create rows of cells
    cells_by_row = {}
    for cell in itervalues(worksheet._cells):
        cells_by_row.setdefault(cell.row, []).append(cell)

    with xf.element("sheetData"):
        for row_idx in sorted(cells_by_row):
            # row meta data
            row_dimension = worksheet.row_dimensions[row_idx]
            row_dimension.style = worksheet._styles.get(row_idx)
            attrs = {'r': '%d' % row_idx,
                     'spans': '1:%d' % worksheet.max_column}
            attrs.update(dict(row_dimension))

            with xf.element("row", attrs):

                row_cells = cells_by_row[row_idx]
                for cell in sorted(row_cells, key=row_sort):
                    write_cell(xf, worksheet, cell, string_table)


def write_cell(xf, worksheet, cell, string_table):
    coordinate = cell.coordinate
    attributes = {'r': coordinate}
    cell_style = worksheet._styles.get(coordinate)
    if cell_style is not None:
        attributes['s'] = '%d' % cell_style

    if cell.data_type != 'f':
        attributes['t'] = cell.data_type

    value = cell.internal_value

    if value in ('', None):
        with xf.element("c", attributes):
            return

    with xf.element('c', attributes):
        if cell.data_type == 'f':
            shared_formula = worksheet.formula_attributes.get(coordinate, {})
            if shared_formula is not None:
                if (shared_formula.get('t') == 'shared'
                    and 'ref' not in shared_formula):
                    value = None
            with xf.element('f', shared_formula):
                if value is not None:
                    xf.write(value[1:])
                    value = None

        if cell.data_type == 's':
            value = string_table.index(value)
        with xf.element("v") as v:
            if value is not None:
                xf.write(safe_string(value))


def write_autofilter(xf, worksheet):
    auto_filter = worksheet.auto_filter
    el = Element('autoFilter', {'ref': auto_filter.ref})
    if (auto_filter.filter_columns
        or auto_filter.sort_conditions):
        for col_id, filter_column in sorted(auto_filter.filter_columns.items()):
            fc = SubElement(el, 'filterColumn', {'colId': str(col_id)})
            attrs = {}
            if filter_column.blank:
                attrs = {'blank': '1'}
            flt = SubElement(fc, 'filters', attrs)
            for val in filter_column.vals:
                SubElement(flt, 'filter', {'val': val})
        if auto_filter.sort_conditions:
            srt = SubElement(el,  'sortState', {'ref': auto_filter.ref})
            for sort_condition in auto_filter.sort_conditions:
                sort_attr = {'ref': sort_condition.ref}
                if sort_condition.descending:
                    sort_attr['descending'] = '1'
                SubElement(srt, 'sortCondtion', sort_attr)
    xf.write(el)


def write_sheetviews(xf, worksheet):
    views = Element('sheetViews')
    view = SubElement(views, 'sheetView', {'workbookViewId': '0'})
    selectionAttrs = {}
    topLeftCell = worksheet.freeze_panes
    if topLeftCell:
        colName, row = coordinate_from_string(topLeftCell)
        column = column_index_from_string(colName)
        pane = 'topRight'
        paneAttrs = {}
        if column > 1:
            paneAttrs['xSplit'] = str(column - 1)
        if row > 1:
            paneAttrs['ySplit'] = str(row - 1)
            pane = 'bottomLeft'
            if column > 1:
                pane = 'bottomRight'
        paneAttrs.update(dict(topLeftCell=topLeftCell,
                              activePane=pane,
                              state='frozen'))
        SubElement(view, 'pane', paneAttrs)
        selectionAttrs['pane'] = pane
        if row > 1 and column > 1:
            SubElement(view, 'selection', {'pane': 'topRight'})
            SubElement(view, 'selection', {'pane': 'bottomLeft'})

    selectionAttrs.update({'activeCell': worksheet.active_cell,
                           'sqref': worksheet.selected_cell})

    SubElement(view, 'selection', selectionAttrs)
    xf.write(views)


def write_format(xf, worksheet):
    attrs = {'defaultRowHeight': '15', 'baseColWidth': '10'}
    dimensions_outline = [dim.outline_level
                          for dim in itervalues(worksheet.column_dimensions)]
    if dimensions_outline:
        outline_level = max(dimensions_outline)
        if outline_level:
            attrs['outlineLevelCol'] = str(outline_level)
    with xf.element('sheetFormatPr', attrs):
        pass


def write_mergecells(xf, worksheet):
    """Write merged cells to xml."""
    merge = Element('mergeCells')
    if worksheet._merged_cells:
        merge.set("count", str(len(worksheet._merged_cells)))
    for range_string in worksheet._merged_cells:
        attrs = {'ref': range_string}
        SubElement(merge, 'mergeCell', attrs)
    xf.write(merge)


def write_datavalidation(xf, worksheet):
    """ Write data validation(s) to xml."""
    # Filter out "empty" data-validation objects (i.e. with 0 cells)
    required_dvs = [x for x in worksheet._data_validations
                    if len(x.cells) or len(x.ranges)]
    if not required_dvs:
        return

    dvs = Element('dataValidations', {'count': str(len(required_dvs))})
    for data_validation in required_dvs:
        dv = SubElement(dvs, 'dataValidation',
                        data_validation.generate_attributes_map())
        if data_validation.formula1:
            SubElement(dv, 'formula1').text = data_validation.formula1
        if data_validation.formula2:
            SubElement(dv, 'formula2').text = data_validation.formula2
    xf.write(dvs)


def write_header_footer(xf, worksheet):
    header = worksheet.header_footer.getHeader()
    footer = worksheet.header_footer.getFooter()
    if header or footer:
        tag = Element('headerFooter')
        if header:
            SubElement(tag, 'oddHeader').text = header
        if worksheet.header_footer.hasFooter():
            SubElement(tag, 'oddFooter').text = footer
        xf.write(tag)


def write_pagebreaks(xf, worksheet):
    breaks = worksheet.page_breaks
    if breaks:
        tag = Element( 'rowBreaks', {'count': str(len(breaks)),
                                     'manualBreakCount': str(len(breaks))})
        for b in breaks:
            SubElement(tag, 'brk', {'id': str(b), 'man': 'true', 'max': '16383',
                             'min': '0'})


def write_rels(xf, worksheet, drawing_id, comments_id):
    """Write relationships for the worksheet to xml."""
    root = Element('{%s}Relationships' % PKG_REL_NS)
    for rel in worksheet.relationships:
        attrs = {'Id': rel.id, 'Type': rel.type, 'Target': rel.target}
        if rel.target_mode:
            attrs['TargetMode'] = rel.target_mode
        SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
    if worksheet._charts or worksheet._images:
        attrs = {'Id': 'rId1',
                 'Type': '%s/drawing' % REL_NS,
                 'Target': '../drawings/drawing%s.xml' % drawing_id}
        SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
    if worksheet._comment_count > 0:
        # there's only one comments sheet per worksheet,
        # so there's no reason to call the Id rIdx
        attrs = {'Id': 'comments',
                 'Type': COMMENTS_NS,
                 'Target': '../comments%s.xml' % comments_id}
        SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
        attrs = {'Id': 'commentsvml',
                 'Type': VML_NS,
                 'Target': '../drawings/commentsDrawing%s.vml' % comments_id}
        SubElement(root, '{%s}Relationship' % PKG_REL_NS, attrs)
    xf.write(root)
