from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

"""Write worksheets to xml representations."""

# Python stdlib imports
from io import BytesIO
from warnings import warn

# package imports
from openpyxl2.xml.functions import xmlfile
from openpyxl2.xml.constants import SHEET_MAIN_NS
from openpyxl2.compat import unicode

from openpyxl2.styles.differential import DifferentialStyle
from openpyxl2.packaging.relationship import Relationship, RelationshipList
from openpyxl2.worksheet.merge import MergeCells, MergeCell
from openpyxl2.worksheet.properties import WorksheetProperties
from openpyxl2.worksheet.hyperlink import (
    Hyperlink,
    HyperlinkList,
)
from openpyxl2.worksheet.related import Related
from openpyxl2.worksheet.table import TablePartList
from openpyxl2.worksheet.header_footer import HeaderFooter
from openpyxl2.worksheet.dimensions import (
    SheetFormatProperties,
    SheetDimension,
)


def write_worksheet(worksheet):
    """Write a worksheet to an xml file."""

    ws = worksheet

    from openpyxl2.worksheet.writer import WorksheetWriter
    writer = WorksheetWriter(ws)
    writer.write_top()
    writer.write_rows()
    writer.write_tail()
    writer.xf.close()
    ws._rels = writer._rels
    ws._hyperlinks = writer._hyperlinks
    return writer.out.getvalue()


def _add_table_headers(ws):
    """
    Check if tables have tableColumns and create them and autoFilter if necessary.
    Column headers will be taken from the first row of the table.
    """

    tables = TablePartList()

    for table in ws._tables:
        if not table.tableColumns:
            table._initialise_columns()
            if table.headerRowCount:
                row = ws[table.ref][0]
                for cell, col in zip(row, table.tableColumns):
                    if cell.data_type != "s":
                        warn("File may not be readable: column headings must be strings.")
                    col.name = unicode(cell.value)
        rel = Relationship(Type=table._rel_type, Target="")
        ws._rels.append(rel)
        table._rel_id = rel.Id
        tables.append(Related(id=rel.Id))

    return tables
