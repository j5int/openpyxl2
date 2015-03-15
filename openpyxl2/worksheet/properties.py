from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Worksheet Properties"""

from openpyxl2.compat import safe_string
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import String, Bool, Typed
from openpyxl2.styles.colors import ColorDescriptor
from openpyxl2.xml.constants import SHEET_MAIN_NS
from openpyxl2.xml.functions import Element
from openpyxl2.styles.colors import Color


class Outline(Serialisable):

    tag = "{%s}outlinePr" % SHEET_MAIN_NS

    applyStyles = Bool(allow_none=True)
    summaryBelow = Bool(allow_none=True)
    summaryRight = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)


    def __init__(self,
                 applyStyles=None,
                 summaryBelow=None,
                 summaryRight=None,
                 showOutlineSymbols=None
                 ):
        self.applyStyles = applyStyles
        self.summaryBelow = summaryBelow
        self.summaryRight = summaryRight
        self.showOutlineSymbols = showOutlineSymbols


class PageSetupPr(Serialisable):

    tag = "{%s}pageSetUpPr" % SHEET_MAIN_NS

    autoPageBreaks = Bool(allow_none=True)
    fitToPage = Bool(allow_none=True)

    def __init__(self, autoPageBreaks=None, fitToPage=None):
        self.autoPageBreaks = autoPageBreaks
        self.fitToPage = fitToPage


class WorksheetProperties(Serialisable):

    tag = "{%s}sheetPr" % SHEET_MAIN_NS

    codeName = String(allow_none=True)
    enableFormatConditionsCalculation = Bool(allow_none=True)
    filterMode = Bool(allow_none=True)
    published = Bool(allow_none=True)
    syncHorizontal = Bool(allow_none=True)
    syncRef = String(allow_none=True)
    syncVertical = Bool(allow_none=True)
    transitionEvaluation = Bool(allow_none=True)
    transitionEntry = Bool(allow_none=True)
    tabColor = ColorDescriptor(allow_none=True)
    outlinePr = Typed(expected_type=Outline, allow_none=True)
    pageSetUpPr = Typed(expected_type=PageSetupPr, allow_none=True)


    def __init__(self,
                 codeName=None,
                 enableFormatConditionsCalculation=None,
                 filterMode=None,
                 published=None,
                 syncHorizontal=None,
                 syncRef=None,
                 syncVertical=None,
                 transitionEvaluation=None,
                 transitionEntry=None,
                 tabColor=None,
                 outlinePr=None,
                 pageSetUpPr=None,
                 ):
        """ Attributes """
        self.codeName = codeName
        self.enableFormatConditionsCalculation = enableFormatConditionsCalculation
        self.filterMode = filterMode
        self.published = published
        self.syncHorizontal = syncHorizontal
        self.syncRef = syncRef
        self.syncVertical = syncVertical
        self.transitionEvaluation = transitionEvaluation
        self.transitionEntry = transitionEntry
        """ Elements """
        self.tabColor = tabColor
        self.outlinePr = outlinePr
        self.pageSetUpPr = pageSetUpPr


def parse_sheetPr(node):
    props = WorksheetProperties(**node.attrib)

    outline = node.find(Outline.tag)
    if outline is not None:
        props.outlinePr = Outline(**outline.attrib)

    page_setup = node.find(PageSetupPr.tag)
    if page_setup is not None:
        props.pageSetUpPr = PageSetupPr(**page_setup.attrib)

    tab_color = node.find('{%s}tabColor' % SHEET_MAIN_NS)
    if tab_color is not None:
        props.tabColor = Color(**dict(tab_color.attrib))

    return props


def write_sheetPr(props):

    attributes = {}
    for k, v in dict(props).items():
        if not isinstance(v, dict):
            attributes[k] = v

    el = Element(props.tag, attributes)

    outline = props.outlinePr
    if outline:
        el.append(Element(outline.tag, dict(outline)))

    page_setup = props.pageSetUpPr
    if page_setup:
        el.append(Element(page_setup.tag, dict(page_setup)))

    if props.tabColor:
        el.append(Element('{%s}tabColor' % SHEET_MAIN_NS, rgb=props.tabColor.value))

    return el
