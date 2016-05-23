from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Write a .xlsx file."""

# Python stdlib imports
from io import BytesIO
from re import match
from zipfile import ZipFile, ZIP_DEFLATED

# package imports
from openpyxl2.xml.constants import (
    ARC_SHARED_STRINGS,
    ARC_CONTENT_TYPES,
    ARC_ROOT_RELS,
    ARC_WORKBOOK_RELS,
    ARC_APP, ARC_CORE,
    ARC_THEME,
    ARC_STYLE,
    ARC_WORKBOOK,
    PACKAGE_WORKSHEETS,
    PACKAGE_CHARTSHEETS,
    PACKAGE_DRAWINGS,
    PACKAGE_CHARTS,
    PACKAGE_IMAGES,
    PACKAGE_XL
    )
from openpyxl2.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl2.xml.functions import tostring, fromstring, Element
from openpyxl2.packaging.manifest import write_content_types
from openpyxl2.packaging.relationship import (
    get_rels_path,
    RelationshipList,
    Relationship,
)
from openpyxl2.packaging.extended import ExtendedProperties

from openpyxl2.writer.strings import write_string_table
from openpyxl2.writer.workbook import (
    write_root_rels,
    write_workbook_rels,
    write_workbook,
)
from openpyxl2.writer.theme import write_theme
from .relations import write_rels
from openpyxl2.writer.worksheet import write_worksheet
from openpyxl2.styles.stylesheet import write_stylesheet

from openpyxl2.comments.writer import CommentWriter

ARC_VBA = ('xl/vba', r'xl/drawings/.*vmlDrawing\d\.vml', 'xl/ctrlProps', 'customUI',
           'xl/activeX', r'xl/media/.*\.emf')


class ExcelWriter(object):
    """Write a workbook object to an Excel file."""

    comment_writer = CommentWriter

    def __init__(self, workbook, archive):
        self.archive = archive
        self.workbook = workbook
        self.vba_modified = set()
        self._tables = []
        self._charts = []
        self._images = []
        self._drawings = []
        self._comments = []
        self.as_template = False


    def write_data(self):
        """Write the various xml files into the zip archive."""
        # cleanup all worksheets
        archive = self.archive

        archive.writestr(ARC_ROOT_RELS, write_root_rels(self.workbook))
        props = ExtendedProperties()
        archive.writestr(ARC_APP, tostring(props.to_tree()))

        archive.writestr(ARC_CORE, tostring(self.workbook.properties.to_tree()))
        if self.workbook.loaded_theme:
            archive.writestr(ARC_THEME, self.workbook.loaded_theme)
        else:
            archive.writestr(ARC_THEME, write_theme())

        self._write_worksheets()
        self._write_comments()
        self._write_chartsheets()
        self._write_images()
        self._write_charts()

        self._write_string_table()
        self._write_external_links()

        stylesheet = write_stylesheet(self.workbook)
        archive.writestr(ARC_STYLE, tostring(stylesheet))

        archive.writestr(ARC_WORKBOOK, write_workbook(self.workbook))
        archive.writestr(ARC_WORKBOOK_RELS, write_workbook_rels(self.workbook))

        if self.workbook.vba_archive:
            vba_archive = self.workbook.vba_archive
            for name in set(vba_archive.namelist()) - self.vba_modified:
                for s in ARC_VBA:
                    if match(s, name):
                        archive.writestr(name, vba_archive.read(name))
                        break

        exts = []
        for n in archive.namelist():
            if "media" in n:
                exts.append(n)
        manifest = write_content_types(self.workbook, as_template=self.as_template, exts=exts)
        archive.writestr(ARC_CONTENT_TYPES, tostring(manifest.to_tree()))


    def _write_string_table(self):
        self.archive.writestr(ARC_SHARED_STRINGS,
                write_string_table(self.workbook.shared_strings))


    def _write_images(self):
        for img in self._images:
            buf = BytesIO()
            img.image.save(buf, format='PNG')
            self.archive.writestr(img._path, buf.getvalue())


    def _write_charts(self):
        for chart in self._charts:
            self.archive.writestr(chart._path, tostring(chart._write()))


    def _write_drawing(self, drawing):
        """
        Write a drawing
        """
        self._drawings.append(drawing)
        drawing_id = len(self._drawings)
        for chart in drawing.charts:
            self._charts.append(chart)
            chart._id = len(self._charts)
        for img in drawing.images:
            self._images.append(img)
            img._id = len(self._images)
        drawingpath = "{0}/drawing{1}.xml".format(PACKAGE_DRAWINGS, drawing_id)
        self.archive.writestr(drawingpath, tostring(drawing._write()))
        self.archive.writestr("{0}/_rels/drawing{1}.xml.rels".format(PACKAGE_DRAWINGS,
                                                                drawing_id), tostring(drawing._write_rels()))
        return drawingpath


    def _write_chartsheets(self):
        for idx, sheet in enumerate(self.workbook.chartsheets, 1):

            sheet._path = "sheet{0}.xml".format(idx)
            arc_path = "{0}/{1}".format(PACKAGE_CHARTSHEETS, sheet._path)
            rels_path = get_rels_path(arc_path)
            xml = tostring(sheet.to_tree())

            self.archive.writestr(arc_path, xml)

            if sheet._charts:
                drawing = SpreadsheetDrawing()
                drawing.charts = sheet._charts
                drawingpath = self._write_drawing(self.archive, drawing)

                rel = Relationship(type="drawing", Target="/" + drawingpath)
                rels = RelationshipList()
                rels.append(rel)
                tree = rels.to_tree()

                self.archive.writestr(rels_path, tostring(tree))


    def _write_comments(self):
        for idx, cw in enumerate(self._comments, 1):

            self.archive.writestr('xl/comments%d.xml' % idx,
                                  cw.write_comments())
            if cw.vml is not None:
                vml = cw.write_comments_vml(cw.vml)
                vml_path = cw.vml_path
            else:
                vml = Element("xml")
                vml_path = 'xl/drawings/commentsDrawing%d.vml' % idx
                cw.write_comments_vml(vml)
            self.archive.writestr(vml_path, vml)


    def _write_worksheets(self):

        for idx, sheet in enumerate(self.workbook.worksheets, 1):

            xml = sheet._write()
            sheet._path = "sheet{0}.xml".format(idx)
            arc_path = "{0}/{1}".format(PACKAGE_WORKSHEETS, sheet._path)
            rels_path = get_rels_path(arc_path)

            self.archive.writestr(arc_path, xml)

            if sheet._charts or sheet._images:
                drawing = SpreadsheetDrawing()
                drawing.charts = sheet._charts
                drawing.images = sheet._images
                drawingpath = self._write_drawing(drawing)

                for r in sheet._rels.Relationship:
                    if "drawing" in r.Type:
                        r.Target = "/" + drawingpath

            if sheet._comments:
                cw = self.comment_writer(sheet)

                if sheet.legacy_drawing is not None:
                    vml = fromstring(self.workbook.vba_archive.read(sheet.legacy_drawing))
                    cw.vml = vml
                    cw.vml_path = sheet.legacy_drawing
                    # Record this file so we don't write it again when we dump out vba_archive
                    self.vba_modified.add(sheet.legacy_drawing)

                self._comments.append(cw)

            for t in sheet._tables:
                self._tables.append(t)
                t.id = len(self._tables)
                t._write(self.archive)
                sheet._rels[t._rel_id].Target = t.path

            if (sheet._rels
                or sheet._comments
                or sheet.legacy_drawing is not None):
                rels = write_rels(sheet, comments_id=len(self._comments))

                self.archive.writestr(rels_path, tostring(rels))


    def _write_external_links(self):
        """Write links to external workbooks"""
        wb = self.workbook
        for idx, link in enumerate(wb._external_links, 1):

            link._path = "{0}{1}.xml".format(link._rel_type, idx)

            arc_path = "{0}/{1}s/{2}".format(PACKAGE_XL, link._rel_type, link._path)
            rels_path = get_rels_path(arc_path)

            xml = link.to_tree()
            self.archive.writestr(arc_path, tostring(xml))
            rels = RelationshipList()
            rels.append(link.file_link)
            self.archive.writestr(rels_path, tostring(rels.to_tree()))


    def save(self, filename):
        """Write data into the archive."""
        self.write_data()
        self.archive.close()


def save_workbook(workbook, filename, as_template=False):
    """Save the given workbook on the filesystem under the name filename.

    :param workbook: the workbook to save
    :type workbook: :class:`openpyxl.workbook.Workbook`

    :param filename: the path to which save the workbook
    :type filename: string

    :rtype: bool

    """
    archive = ZipFile(filename, 'w', ZIP_DEFLATED, allowZip64=True)
    writer = ExcelWriter(workbook, archive)
    writer.as_template = as_template
    writer.save(filename)
    return True


def save_virtual_workbook(workbook, as_template=False):
    """Return an in-memory workbook, suitable for a Django response."""
    temp_buffer = BytesIO()
    archive = ZipFile(temp_buffer, 'w', ZIP_DEFLATED, allowZip64=True)

    writer = ExcelWriter(workbook, archive)
    writer.as_template = as_template

    try:
        writer.write_data()
    finally:
        archive.close()

    virtual_workbook = temp_buffer.getvalue()
    temp_buffer.close()
    return virtual_workbook
