from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

"""Write the workbook global settings to the archive."""

from copy import copy

from openpyxl2.utils import absolute_coordinate, quote_sheetname
from openpyxl2.xml.constants import (
    ARC_APP,
    ARC_CORE,
    ARC_WORKBOOK,
    PKG_REL_NS,
    CUSTOMUI_NS,
    ARC_ROOT_RELS,
)
from openpyxl2.xml.functions import tostring, fromstring

from openpyxl2.worksheet import Worksheet
from openpyxl2.chartsheet import Chartsheet
from openpyxl2.packaging.relationship import Relationship, RelationshipList
from openpyxl2.workbook.defined_name import DefinedName
from openpyxl2.workbook.external_reference import ExternalReference
from openpyxl2.workbook.parser import ChildSheet, WorkbookPackage
from openpyxl2.workbook.properties import CalcProperties, WorkbookProperties
from openpyxl2.workbook.views import BookView


def write_root_rels(workbook):
    """Write the relationships xml."""

    rels = RelationshipList()

    rel = Relationship(type="officeDocument", Target=ARC_WORKBOOK)
    rels.append(rel)

    rel = Relationship(Target=ARC_CORE, Type="%s/metadata/core-properties" % PKG_REL_NS)
    rels.append(rel)

    rel = Relationship(type="extended-properties", Target=ARC_APP)
    rels.append(rel)

    if workbook.vba_archive is not None:
        # See if there was a customUI relation and reuse it
        xml = fromstring(workbook.vba_archive.read(ARC_ROOT_RELS))
        root_rels = RelationshipList.from_tree(xml)
        custom_ui = list(root_rels.find(CUSTOMUI_NS))
        if custom_ui:
            rels.append(custom_ui[0])

    return tostring(rels.to_tree())


def get_active_sheet(wb):
    """
    Return the index of the active sheet.
    If the sheet set to active is hidden return the next visible sheet
    """
    idx = wb._active_sheet_index
    sheet = wb.active
    if sheet.sheet_state == "visible":
        return idx

    for idx, sheet in enumerate(wb._sheets[idx:], idx):
        if sheet.sheet_state == "visible":
            wb.active = idx
            return idx

    raise IndexError("At least one sheet must be visible")


def write_workbook(workbook):
    """Write the core workbook xml."""

    wb = workbook
    wb.rels = RelationshipList()

    root = WorkbookPackage()

    props = WorkbookProperties()
    if wb.code_name is not None:
        props.codeName = wb.code_name
    root.workbookPr = props

    # book views
    active = get_active_sheet(wb)
    view = BookView(activeTab=active)
    root.bookViews =[view]

    # worksheets
    for idx, sheet in enumerate(wb._sheets, 1):
        sheet_node = ChildSheet(name=sheet.title, sheetId=idx, id="rId{0}".format(idx))
        rel = Relationship(
            type=sheet._rel_type,
            Target='{0}s/{1}'.format(sheet._rel_type, sheet._path)
        )
        wb.rels.append(rel)

        if not sheet.sheet_state == 'visible':
            if len(wb._sheets) == 1:
                raise ValueError("The only worksheet of a workbook cannot be hidden")
            sheet_node.state = sheet.sheet_state
        root.sheets.append(sheet_node)

    # external references
    if wb._external_links:
        # need to match a counter with a workbook's relations
        counter = len(wb._sheets) + 3 # strings, styles, theme
        if wb.vba_archive:
            counter += 1
        for idx, _ in enumerate(wb._external_links, counter+1):
            ext = ExternalReference(id="rId{0}".format(idx))
            root.externalReferences.append(ext)

    # Defined names
    defined_names = copy(wb.defined_names) # don't add special defns to workbook itself.

    # Defined names -> autoFilter
    for idx, sheet in enumerate(wb.worksheets):
        auto_filter = sheet.auto_filter.ref
        if auto_filter:
            name = DefinedName(name='_FilterDatabase', localSheetId=idx, hidden=True)
            name.value = "{0}!{1}".format(quote_sheetname(sheet.title),
                                          absolute_coordinate(auto_filter)
                                          )
            defined_names.append(name)

        # print titles
        if sheet.print_titles:
            name = DefinedName(name="PrintTitles", localSheetId=idx)
            name.value = quote_sheetname(sheet.print_titles)
            defined_names.append(name)

        # print areas
        if sheet.print_area:
            name = DefinedName(name="PrintArea", localSheetId=idx)
            name.value = "{0}!{1}".format(quote_sheetname(sheet.title), sheet.print_area)
            defined_names.append(name)

    root.definedNames = defined_names

    root.calcPr = CalcProperties(calcId=124519, fullCalcOnLoad=True)

    return tostring(root.to_tree())


def write_workbook_rels(workbook):
    """Write the workbook relationships xml."""
    wb = workbook

    external_links = workbook._external_links
    if external_links:
        for idx, link in enumerate(external_links, 1):
            ext =  Relationship(type='externalLink',
                                Target='externalLinks/externalLink{0}.xml'.format(idx)
                                )
            wb.rels.append(ext)

    strings =  Relationship(type='sharedStrings', Target='sharedStrings.xml')
    wb.rels.append(strings)

    styles =  Relationship(type='styles', Target='styles.xml')
    wb.rels.append(styles)

    theme =  Relationship(type='theme', Target='theme/theme1.xml')
    wb.rels.append(theme)

    if workbook.vba_archive:
        vba =  Relationship(type='vbaProject', Target='vbaProject.bin')
        vba.type ='http://schemas.microsoft.com/office/2006/relationships/vbaProject'
        wb.rels.append(vba)

    return tostring(wb.rels.to_tree())
