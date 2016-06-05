from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""
File manifest
"""
import mimetypes
import os.path

from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import String, Sequence
from openpyxl2.xml.functions import fromstring
from openpyxl2.xml.constants import (
    ARC_CORE,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK,
    ARC_APP,
    ARC_THEME,
    ARC_STYLE,
    ARC_SHARED_STRINGS,
    EXTERNAL_LINK,
    THEME_TYPE,
    STYLES_TYPE,
    XLSX,
    XLSM,
    XLTM,
    XLTX,
    WORKSHEET_TYPE,
    COMMENTS_TYPE,
    SHARED_STRINGS,
    DRAWING_TYPE,
    CHART_TYPE,
    CHARTSHAPE_TYPE,
    CHARTSHEET_TYPE,
    CONTYPES_NS
)
from openpyxl2.xml.functions import tostring

# initialise mime-types
mimetypes.init()
mimetypes.add_type('application/xml', ".xml")
mimetypes.add_type('application/vnd.openxmlformats-package.relationships+xml', ".rels")
mimetypes.add_type("application/vnd.ms-office.activeX", ".bin")
mimetypes.add_type("application/vnd.openxmlformats-officedocument.vmlDrawing", ".vml")
mimetypes.add_type("image/x-emf", ".emf")


class FileExtension(Serialisable):

    tagname = "Default"

    Extension = String()
    ContentType = String()

    def __init__(self, Extension, ContentType):
        self.Extension = Extension
        self.ContentType = ContentType


class Override(Serialisable):

    tagname = "Override"

    PartName = String()
    ContentType = String()

    def __init__(self, PartName, ContentType):
        self.PartName = PartName
        self.ContentType = ContentType


DEFAULT_TYPES = [
    FileExtension("rels", "application/vnd.openxmlformats-package.relationships+xml"),
    FileExtension("xml", "application/xml"),
]

DEFAULT_OVERRIDE = [
    Override("/" + ARC_WORKBOOK, XLSX), # Workbook
    Override("/" + ARC_SHARED_STRINGS, SHARED_STRINGS), # Shared strings
    Override("/" + ARC_STYLE, STYLES_TYPE), # Styles
    Override("/" + ARC_THEME, THEME_TYPE), # Theme
    Override("/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml"),
    Override("/docProps/app.xml", "application/vnd.openxmlformats-officedocument.extended-properties+xml")
]


class Manifest(Serialisable):

    tagname = "Types"

    Default = Sequence(expected_type=FileExtension, unique=True)
    Override = Sequence(expected_type=Override, unique=True)
    path = "[Content_Types].xml"

    __elements__ = ("Default", "Override")

    def __init__(self,
                 Default=(),
                 Override=(),
                 ):
        if not Default:
            Default = DEFAULT_TYPES
        self.Default = Default
        if not Override:
            Override = DEFAULT_OVERRIDE
        self.Override = Override


    @property
    def filenames(self):
        return [part.PartName for part in self.Override]


    @property
    def extensions(self):
        exts = set([os.path.splitext(part.PartName)[-1] for part in self.Override])
        return [(ext[1:], mimetypes.types_map[ext]) for ext in sorted(exts)]


    def to_tree(self):
        """
        Custom serialisation method to allow setting a default namespace
        """
        defaults = [t.Extension for t in self.Default]
        for ext, mime in self.extensions:
            if ext not in defaults:
                mime = FileExtension(ext, mime)
                self.Default.append(mime)
        tree = super(Manifest, self).to_tree()
        tree.set("xmlns", CONTYPES_NS)
        return tree


    def __contains__(self, content_type):
        """
        Check whether a particular content type is contained
        """
        for t in self.Override:
            if t.ContentType == content_type:
                return True


    def find(self, content_type):
        """
        Find specific content-type
        """
        for t in self.Override:
            if t.ContentType == content_type:
                return t


    def append(self, obj):
        """
        Add content object to the package manifest
        # needs a contract...
        """
        ct = Override(PartName=obj.path, ContentType=obj.mime_type)
        self.Override.append(ct)


    def _write(self, archive, workbook, as_template=False):
        """
        Write manifest to the archive
        """
        self._write_content_types(workbook, as_template=as_template, filenames=archive.namelist())
        archive.writestr(self.path, tostring(self.to_tree()))


    def _write_content_types(self, workbook, as_template=False, filenames=None):

        for n in self.filenames:
            if n.endswith('.vml'):
                ext = FileExtension("vml", mimetypes.types_map[".vml"])
                self.manifest.Default.append(ext)
                break

        if filenames is not None:
            for fn in filenames:
                ext = os.path.splitext(fn)[-1]
                if not ext:
                    continue
                mime = mimetypes.types_map[ext]
                fe = FileExtension(ext[1:], mime)
                self.Default.append(fe)

        if workbook.vba_archive:
            node = fromstring(workbook.vba_archive.read(ARC_CONTENT_TYPES))
            mf = Manifest.from_tree(node)
            for override in mf.Override:
                if override not in self.Override:
                    self.Override.append(override)

        # templates
        for part in self.Override:
            if part.PartName == "/xl/workbook.xml":
                ct = as_template and XLTX or XLSX
                if workbook.vba_archive:
                    ct = as_template and XLTM or XLSM
                part.ContentType = ct

        return self
