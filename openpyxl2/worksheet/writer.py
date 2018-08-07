from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from collections import defaultdict
from io import BytesIO
from operator import itemgetter
from openpyxl2.xml.functions import xmlfile
from openpyxl2.xml.constants import SHEET_MAIN_NS
from openpyxl2.compat import unicode

from openpyxl2.packaging.relationship import Relationship, RelationshipList
from openpyxl2.styles.differential import DifferentialStyle
from .dimensions import SheetDimension
from .hyperlink import HyperlinkList
from .merge import MergeCell, MergeCells
from .related import Related
from .table import TablePartList

from openpyxl2.writer.etree_worksheet import write_cell


class WorksheetWriter:


    def __init__(self, ws, out=None):
        self.ws = ws
        if out is None:
            out = BytesIO()
        self.out = out
        self._rels = RelationshipList()
        self._hyperlinks = []
        self.xf = self.write()
        next(self.xf) # start generator


    def write_properties(self):
        props = self.ws.sheet_properties
        self.xf.send(props.to_tree())


    def write_dimensions(self):
        """
        Write worksheet size if known
        """
        ref = getattr(self.ws, 'calculate_dimension')
        if ref:
            dim = SheetDimension(ref())
            self.xf.send(dim.to_tree())


    def write_format(self):
        self.ws.sheet_format.outlineLevelCol = self.ws.column_dimensions.max_outline
        fmt = self.ws.sheet_format
        self.xf.send(fmt.to_tree())


    def write_views(self):
        views = self.ws.views
        self.xf.send(views.to_tree())


    def write_cols(self):
        cols = self.ws.column_dimensions
        self.xf.send(cols.to_tree())


    def write_top(self):
        """
        Write all elements up to rows:
        properties
        dimensions
        views
        format
        cols
        """
        self.write_properties()
        self.write_dimensions()
        self.write_views()
        self.write_format()
        self.write_cols()


    def rows(self):
        """Return all rows, and any cells that they contain"""
        # order cells by row
        rows = defaultdict(list)
        for (row, col), cell in self.ws._cells.items():
            rows[row].append((col, cell))

        # add empty rows if styling has been applied
        for row in set(self.ws.row_dimensions.keys()) - set(rows.keys()):
            rows[row] = []

        return sorted(rows.items())


    def write_rows(self):
        xf = self.xf.send(True)

        with xf.element("sheetData"):
            for row_idx, row in self.rows():
                row = sorted(row, key=itemgetter(0))
                self.write_row(xf, row, row_idx)


    def write_row(self, xf, row, row_idx):
        max_column = self.ws.max_column

        attrs = {'r': '%d' % row_idx}
        dims = self.ws.row_dimensions
        attrs.update(dims.get(row_idx, {}))

        with xf.element("row", attrs):

            for col, cell in row:
                if cell._comment is not None:
                    comment = CommentRecord.from_cell(cell)
                    worksheet._comments.append(comment)
                if (
                    cell._value is None
                    and not cell.has_style
                    and not cell._comment
                    ):
                    continue
                el = write_cell(xf, self.ws, cell, cell.has_style)


    def write_protection(self):
        prot = self.ws.protection
        if prot:
            self.xf.send(prot.to_tree())


    def write_filter(self):
        flt = self.ws.auto_filter
        if flt:
            self.xf.send(flt.to_tree())


    def write_sort(self):
        """
        As per discusion with the OOXML Working Group global sort state is not required.
        openpyxl never reads it from existing files
        """
        pass


    def write_merged_cells(self):
        merged = self.ws.merged_cells
        if merged:
            cells = [MergeCell(str(ref)) for ref in self.ws.merged_cells]
            self.xf.send(MergeCells(mergeCell=cells).to_tree())


    def write_formatting(self):
        df = DifferentialStyle()
        wb = self.ws.parent
        for cf in self.ws.conditional_formatting:
            for rule in cf.rules:
                if rule.dxf and rule.dxf != df:
                    rule.dxfId = wb._differential_styles.add(rule.dxf)
            self.xf.send(cf.to_tree())


    def write_validations(self):
        dv = self.ws.data_validations
        if dv:
            self.xf.send(dv.to_tree())


    def write_hyperlinks(self):
        links = HyperlinkList()

        for link in self._hyperlinks:
            if link.target:
                rel = Relationship(type="hyperlink", TargetMode="External", Target=link.target)
                self._rels.append(rel)
                link.id = rel.id
            links.hyperlink.append(link)

        if links:
            self.xf.send(links.to_tree())


    def write_print(self):
        print_options = self.ws.print_options
        if print_options:
            self.xf.send(print_options.to_tree())


    def write_margins(self):
        margins = self.ws.page_margins
        if margins:
            self.xf.send(margins.to_tree())


    def write_page(self):
        setup = self.ws.page_setup
        if setup:
            self.xf.send(setup.to_tree())


    def write_header(self):
        hf = self.ws.HeaderFooter
        if hf:
            self.xf.send(hf.to_tree())


    def write_breaks(self):
        brk = self.ws.page_breaks
        if brk:
            self.xf.send(brk.to_tree())


    def write_drawings(self):
        if self.ws._charts or self.ws._images:
            rel = Relationship(type="drawing", Target="")
            self._rels.append(rel)
            drawing = Related()
            drawing.id = rel.id
            self.xf.send(drawing.to_tree("drawing"))


    def write_legacy(self):
        """
        Comments & VBA controls use VML and require an additional element
        that is no longer in the specification.
        """
        if (self.ws.legacy_drawing is not None or self.ws._comments):
            legacy = Related(id="anysvml")
            self.xf.send(legacy.to_tree("legacyDrawing"))


    def write_tables(self):
        tables = TablePartList()

        for table in self.ws._tables:
            if not table.tableColumns:
                table._initialise_columns()
                if table.headerRowCount:
                    row = self.ws[table.ref][0]
                    for cell, col in zip(row, table.tableColumns):
                        if cell.data_type != "s":
                            warn("File may not be readable: column headings must be strings.")
                        col.name = unicode(cell.value)
            rel = Relationship(Type=table._rel_type, Target="")
            self._rels.append(rel)
            table._rel_id = rel.Id
            tables.append(Related(id=rel.Id))

        if tables:
            self.xf.send(tables.to_tree())


    def write(self):
        with xmlfile(self.out) as xf:
            with xf.element("worksheet", xmlns=SHEET_MAIN_NS):
                try:
                    while True:
                        el = (yield)
                        if el is True:
                            yield xf
                        else:
                            xf.write(el)
                except GeneratorExit:
                    pass


    def write_tail(self):
        """
        Write all elements after the rows
        calc properties
        protection
        protected ranges #
        scenarios #
        filters
        sorts # always ignored
        data consolidation #
        custom views #
        merged cells
        phonetic properties #
        conditional formatting
        data validation
        hyperlinks
        print options
        page margins
        page setup
        header
        row breaks
        col breaks
        custom properties #
        cell watches #
        ignored errors #
        smart tags #
        drawing
        drawingHF #
        background #
        OLE objects #
        controls #
        web publishing #
        tables
        """
        self.write_protection()
        self.write_filter()
        self.write_merged_cells()
        self.write_formatting()
        self.write_validations()
        self.write_hyperlinks()
        self.write_print()
        self.write_margins()
        self.write_page()
        self.write_header()
        self.write_breaks()
        self.write_drawings()
        self.write_legacy()
        self.write_tables()