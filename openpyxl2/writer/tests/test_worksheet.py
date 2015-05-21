from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import datetime
import decimal
from io import BytesIO

import pytest

from openpyxl2.xml.functions import tostring, xmlfile
from openpyxl2 import Workbook

from .. worksheet import write_worksheet

from openpyxl2.tests.helper import compare_xml
from openpyxl2.worksheet.properties import PageSetupProperties


@pytest.fixture
def worksheet():
    from openpyxl2 import Workbook
    wb = Workbook()
    return wb.active


@pytest.fixture
def DummyWorksheet():

    class DummyWorksheet:

        def __init__(self):
            self._styles = {}
            self.column_dimensions = {}
            self.parent = Workbook()

    return DummyWorksheet()


@pytest.fixture
def write_cols():
    from .. worksheet import write_cols
    return write_cols


@pytest.fixture
def ColumnDimension():
    from openpyxl2.worksheet.dimensions import ColumnDimension
    return ColumnDimension


def test_no_cols(write_cols, DummyWorksheet):

    cols = write_cols(DummyWorksheet)
    assert cols is None


def test_col_widths(write_cols, ColumnDimension, DummyWorksheet):
    ws = DummyWorksheet
    ws.column_dimensions['A'] = ColumnDimension(worksheet=ws, width=4)
    cols = write_cols(ws)
    xml = tostring(cols)
    expected = """<cols><col width="4" min="1" max="1" customWidth="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_col_style(write_cols, ColumnDimension, DummyWorksheet):
    from openpyxl2.styles import Font
    ws = DummyWorksheet
    cd = ColumnDimension(worksheet=ws)
    ws.column_dimensions['A'] = cd
    cd.font = Font(color="FF0000")
    cols = write_cols(ws)
    xml = tostring(cols)
    expected = """<cols><col max="1" min="1" style="1"></col></cols>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_lots_cols(write_cols, ColumnDimension, DummyWorksheet):
    from openpyxl2.styles import Font
    ws = DummyWorksheet
    from openpyxl2.cell import get_column_letter
    for i in range(1, 15):
        label = get_column_letter(i)
        cd = ColumnDimension(worksheet=ws)
        cd.font = Font(name=label)
        dict(cd)  # create style_id in order for test
        ws.column_dimensions[label] = cd
    cols = write_cols(ws)
    xml = tostring(cols)
    expected = """<cols>
   <col max="1" min="1" style="1"></col>
   <col max="2" min="2" style="2"></col>
   <col max="3" min="3" style="3"></col>
   <col max="4" min="4" style="4"></col>
   <col max="5" min="5" style="5"></col>
   <col max="6" min="6" style="6"></col>
   <col max="7" min="7" style="7"></col>
   <col max="8" min="8" style="8"></col>
   <col max="9" min="9" style="9"></col>
   <col max="10" min="10" style="10"></col>
   <col max="11" min="11" style="11"></col>
   <col max="12" min="12" style="12"></col>
   <col max="13" min="13" style="13"></col>
   <col max="14" min="14" style="14"></col>
 </cols>
"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


@pytest.fixture
def write_format():
    from .. worksheet import write_format
    return write_format


def test_sheet_format(write_format, ColumnDimension, DummyWorksheet):
    fmt = write_format(DummyWorksheet)
    xml = tostring(fmt)
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10"/>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_format(write_format, ColumnDimension, DummyWorksheet):
    worksheet = DummyWorksheet
    worksheet.column_dimensions['A'] = ColumnDimension(worksheet=worksheet,
                                                       outline_level=1)
    fmt = write_format(worksheet)
    xml = tostring(fmt)
    expected = """<sheetFormatPr defaultRowHeight="15" baseColWidth="10" outlineLevelCol="1" />"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


def test_outline_cols(write_cols, ColumnDimension, DummyWorksheet):
    worksheet = DummyWorksheet
    worksheet.column_dimensions['A'] = ColumnDimension(worksheet=worksheet,
                                                       outline_level=1)
    cols = write_cols(worksheet)
    xml = tostring(cols)
    expected = """<cols><col max="1" min="1" outlineLevel="1"/></cols>"""
    diff = compare_xml(expected, xml)
    assert diff is None, diff


@pytest.fixture
def write_rows():
    from .. etree_worksheet import write_rows
    return write_rows



@pytest.mark.parametrize("value, expected",
                         [
                             (9781231231230, """<c t="n" r="A1"><v>9781231231230</v></c>"""),
                             (decimal.Decimal('3.14'), """<c t="n" r="A1"><v>3.14</v></c>"""),
                             (1234567890, """<c t="n" r="A1"><v>1234567890</v></c>"""),
                             ("=sum(1+1)", """<c r="A1"><f>sum(1+1)</f><v></v></c>"""),
                             (True, """<c t="b" r="A1"><v>1</v></c>"""),
                             ("Hello", """<c t="s" r="A1"><v>0</v></c>"""),
                             ("", """<c r="A1" t="s"></c>"""),
                             (None, """<c r="A1" t="n"></c>"""),
                             (datetime.date(2011, 12, 25), """<c r="A1" t="n" s="1"><v>40902</v></c>"""),
                         ])
def test_write_cell(worksheet, value, expected):
    from openpyxl2.cell import Cell
    from .. etree_worksheet import write_cell
    ws = worksheet
    ws['A1'] = value

    el = write_cell(ws, ws['A1'])
    xml = tostring(el)
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_formula(worksheet, write_rows):
    ws = worksheet

    ws['F1'] = 10
    ws['F2'] = 32
    ws['F3'] = '=F1+F2'
    ws['A4'] = '=A1+A2+A3'
    ws.formula_attributes['A4'] = {'t': 'shared', 'ref': 'A4:C4', 'si': '0'}
    ws['B4'] = '=1'
    ws.formula_attributes['B4'] = {'t': 'shared', 'si': '0'}
    ws['C4'] = '=1'
    ws.formula_attributes['C4'] = {'t': 'shared', 'si': '0'}

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)

    xml = out.getvalue()
    expected = """
    <sheetData>
      <row r="1" spans="1:6">
        <c r="F1" t="n">
          <v>10</v>
        </c>
      </row>
      <row r="2" spans="1:6">
        <c r="F2" t="n">
          <v>32</v>
        </c>
      </row>
      <row r="3" spans="1:6">
        <c r="F3">
          <f>F1+F2</f>
          <v></v>
        </c>
      </row>
      <row r="4" spans="1:6">
        <c r="A4">
          <f ref="A4:C4" si="0" t="shared">A1+A2+A3</f>
          <v></v>
        </c>
        <c r="B4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
        <c r="C4">
          <f si="0" t="shared"></f>
          <v></v>
        </c>
      </row>
    </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_height(worksheet, write_rows):
    from openpyxl2.worksheet.dimensions import RowDimension
    ws = worksheet
    ws['F1'] = 10

    ws.row_dimensions[1] = RowDimension(ws, height=30)
    ws.row_dimensions[2] = RowDimension(ws, height=30)

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws)
    xml = out.getvalue()
    expected = """
     <sheetData>
       <row customHeight="1" ht="30" r="1" spans="1:6">
         <c r="F1" t="n">
           <v>10</v>
         </c>
       </row>
       <row customHeight="1" ht="30" r="2" spans="1:6"></row>
     </sheetData>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_get_rows_to_write(worksheet):
    from .. etree_worksheet import get_rows_to_write

    ws = worksheet
    ws['A10'] = "test"
    ws.row_dimensions[10] = None
    ws.row_dimensions[2] = None

    cells_by_row = get_rows_to_write(ws)

    assert cells_by_row == [
        (2, []),
        (10, [(1, ws['A10'])])
    ]

@pytest.fixture
def write_autofilter():
    from .. worksheet import write_autofilter
    return write_autofilter


def test_auto_filter(worksheet, write_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'
    af = write_autofilter(ws)
    xml = tostring(af)
    expected = """<autoFilter ref="A1:F1"></autoFilter>"""
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_auto_filter_filter_column(worksheet, write_autofilter):
    ws = worksheet
    ws.auto_filter.ref = 'A1:F1'
    ws.auto_filter.add_filter_column(0, ["0"], blank=True)

    af = write_autofilter(ws)
    xml = tostring(af)
    expected = """
    <autoFilter ref="A1:F1">
      <filterColumn colId="0">
        <filters blank="1">
          <filter val="0"></filter>
        </filters>
      </filterColumn>
    </autoFilter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_auto_filter_sort_condition(worksheet, write_autofilter):
    ws = worksheet
    ws.cell('A1').value = 'header'
    ws.cell('A2').value = 1
    ws.cell('A3').value = 0
    ws.auto_filter.ref = 'A2:A3'
    ws.auto_filter.add_sort_condition('A2:A3', descending=True)

    af = write_autofilter(ws)
    xml = tostring(af)
    expected = """
    <autoFilter ref="A2:A3">
    <sortState ref="A2:A3">
      <sortCondtion descending="1" ref="A2:A3"></sortCondtion>
    </sortState>
    </autoFilter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_auto_filter_worksheet(worksheet, write_worksheet):
    worksheet.auto_filter.ref = 'A1:F1'
    xml = write_worksheet(worksheet, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr>
        <outlinePr summaryBelow="1" summaryRight="1"/>
        <pageSetUpPr/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <autoFilter ref="A1:F1"/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_merge(worksheet):
    from .. worksheet import write_mergecells

    ws = worksheet
    ws.cell('A1').value = 'Cell A1'
    ws.cell('B1').value = 'Cell B1'

    ws.merge_cells('A1:B1')
    merge = write_mergecells(ws)
    xml = tostring(merge)
    expected = """
      <mergeCells count="1">
        <mergeCell ref="A1:B1"/>
      </mergeCells>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_no_merge(worksheet):
    from .. worksheet import write_mergecells

    merge = write_mergecells(worksheet)
    assert merge is None


def test_header_footer(worksheet):
    ws = worksheet
    ws.header_footer.left_header.text = "Left Header Text"
    ws.header_footer.center_header.text = "Center Header Text"
    ws.header_footer.center_header.font_name = "Arial,Regular"
    ws.header_footer.center_header.font_size = 6
    ws.header_footer.center_header.font_color = "445566"
    ws.header_footer.right_header.text = "Right Header Text"
    ws.header_footer.right_header.font_name = "Arial,Bold"
    ws.header_footer.right_header.font_size = 8
    ws.header_footer.right_header.font_color = "112233"
    ws.header_footer.left_footer.text = "Left Footer Text\nAnd &[Date] and &[Time]"
    ws.header_footer.left_footer.font_name = "Times New Roman,Regular"
    ws.header_footer.left_footer.font_size = 10
    ws.header_footer.left_footer.font_color = "445566"
    ws.header_footer.center_footer.text = "Center Footer Text &[Path]&[File] on &[Tab]"
    ws.header_footer.center_footer.font_name = "Times New Roman,Bold"
    ws.header_footer.center_footer.font_size = 12
    ws.header_footer.center_footer.font_color = "778899"
    ws.header_footer.right_footer.text = "Right Footer Text &[Page] of &[Pages]"
    ws.header_footer.right_footer.font_name = "Times New Roman,Italic"
    ws.header_footer.right_footer.font_size = 14
    ws.header_footer.right_footer.font_color = "AABBCC"

    from .. worksheet import write_header_footer
    hf = write_header_footer(ws)
    xml = tostring(hf)
    expected = """
    <headerFooter>
      <oddHeader>&amp;L&amp;"Calibri,Regular"&amp;K000000Left Header Text&amp;C&amp;"Arial,Regular"&amp;6&amp;K445566Center Header Text&amp;R&amp;"Arial,Bold"&amp;8&amp;K112233Right Header Text</oddHeader>
      <oddFooter>&amp;L&amp;"Times New Roman,Regular"&amp;10&amp;K445566Left Footer Text_x000D_And &amp;D and &amp;T&amp;C&amp;"Times New Roman,Bold"&amp;12&amp;K778899Center Footer Text &amp;Z&amp;F on &amp;A&amp;R&amp;"Times New Roman,Italic"&amp;14&amp;KAABBCCRight Footer Text &amp;P of &amp;N</oddFooter>
    </headerFooter>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_no_header(worksheet):
    from .. worksheet import write_header_footer

    hf = write_header_footer(worksheet)
    assert hf is None


def test_hyperlink(worksheet):
    from .. worksheet import write_hyperlinks

    ws = worksheet
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com"

    hyper = write_hyperlinks(ws)
    assert len(worksheet._rels) == 1
    assert worksheet._rels[0].target == "http://test.com"
    xml = tostring(hyper)
    expected = """
    <hyperlinks xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <hyperlink r:id="rId1" ref="A1"/>
    </hyperlinks>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_no_hyperlink(worksheet):
    from .. worksheet import write_hyperlinks

    l = write_hyperlinks(worksheet)
    assert l is None


@pytest.mark.xfail
@pytest.mark.pil_required
def test_write_hyperlink_image_rels(Workbook, Image, datadir):
    datadir.chdir()
    wb = Workbook()
    ws = wb.create_sheet()
    ws.cell('A1').value = "test"
    ws.cell('A1').hyperlink = "http://test.com/"
    i = Image("plain.png")
    ws.add_image(i)
    raise ValueError("Resulting file is invalid")
    # TODO write integration test with duplicate relation ids then fix


@pytest.fixture
def worksheet_with_cf(worksheet):
    from openpyxl2.formatting import ConditionalFormatting
    worksheet.conditional_formating = ConditionalFormatting()
    return worksheet


@pytest.fixture
def write_conditional_formatting():
    from .. worksheet import write_conditional_formatting
    return write_conditional_formatting


def test_conditional_formatting_customRule(worksheet_with_cf, write_conditional_formatting):
    ws = worksheet_with_cf
    from openpyxl2.formatting.rule import Rule

    ws.conditional_formatting.add('C1:C10',
                                  Rule(**{'type': 'expression',
                                          'formula': ['ISBLANK(C1)'], 'stopIfTrue': '1'}
                                       )
                                  )
    cfs = write_conditional_formatting(ws)
    xml = b""
    for cf in cfs:
        xml += tostring(cf)

    diff = compare_xml(xml, """
    <conditionalFormatting sqref="C1:C10">
      <cfRule type="expression" stopIfTrue="1" priority="1">
        <formula>ISBLANK(C1)</formula>
      </cfRule>
    </conditionalFormatting>
    """)
    assert diff is None, diff


def test_conditional_font(worksheet_with_cf, write_conditional_formatting):
    """Test to verify font style written correctly."""

    # Create cf rule
    from openpyxl2.styles import PatternFill, Font, Color
    from openpyxl2.formatting.rule import CellIsRule

    redFill = PatternFill(start_color=Color('FFEE1111'),
                   end_color=Color('FFEE1111'),
                   patternType='solid')
    whiteFont = Font(color=Color("FFFFFFFF"))

    ws = worksheet_with_cf
    ws.conditional_formatting.add('A1:A3',
                                  CellIsRule(operator='equal',
                                             formula=['"Fail"'],
                                             stopIfTrue=False,
                                             font=whiteFont,
                                             fill=redFill))

    cfs = write_conditional_formatting(ws)
    xml = b""
    for cf in cfs:
        xml += tostring(cf)
    diff = compare_xml(xml, """
    <conditionalFormatting sqref="A1:A3">
      <cfRule operator="equal" priority="1" type="cellIs" dxfId="0" stopIfTrue="0">
        <formula>"Fail"</formula>
      </cfRule>
    </conditionalFormatting>
    """)
    assert diff is None, diff


def test_formula_rule(worksheet_with_cf, write_conditional_formatting):
    from openpyxl2.formatting.rule import FormulaRule

    ws = worksheet_with_cf
    ws.conditional_formatting.add('C1:C10',
                                  FormulaRule(
                                      formula=['ISBLANK(C1)'],
                                      stopIfTrue=True)
                                  )
    cfs = write_conditional_formatting(ws)
    xml = b""
    for cf in cfs:
        xml += tostring(cf)

    diff = compare_xml(xml, """
    <conditionalFormatting sqref="C1:C10">
      <cfRule type="expression" stopIfTrue="1" priority="1">
        <formula>ISBLANK(C1)</formula>
      </cfRule>
    </conditionalFormatting>
    """)
    assert diff is None, diff


@pytest.fixture
def write_worksheet():
    from .. worksheet import write_worksheet
    return write_worksheet


def test_write_empty(worksheet, write_worksheet):
    ws = worksheet
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr>
        <outlinePr summaryRight="1" summaryBelow="1"/>
        <pageSetUpPr/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection sqref="A1" activeCell="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_vba(worksheet, write_worksheet):
    ws = worksheet
    ws.vba_code = {"codeName":"Sheet1"}
    ws.vba_controls = "rId2"
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr codeName="Sheet1">
        <outlinePr summaryBelow="1" summaryRight="1"/>
        <pageSetUpPr/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
      <legacyDrawing r:id="rId2"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_protection(worksheet, write_worksheet):
    ws = worksheet
    ws.protection.enable()
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
      <sheetPr>
        <outlinePr summaryBelow="1" summaryRight="1"/>
        <pageSetUpPr/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <sheetProtection sheet="1" objects="0" selectLockedCells="0" selectUnlockedCells="0" scenarios="0" formatCells="1" formatColumns="1" formatRows="1" insertColumns="1" insertRows="1" insertHyperlinks="1" deleteColumns="1" deleteRows="1" sort="1" autoFilter="1" pivotTables="1"/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_comments(worksheet, write_worksheet):
    ws = worksheet
    worksheet._comment_count = 1
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr>
        <outlinePr summaryBelow="1" summaryRight="1"/>
        <pageSetUpPr/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection activeCell="A1" sqref="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins bottom="1" footer="0.5" header="0.5" left="0.75" right="0.75" top="1"/>
      <legacyDrawing r:id="commentsvml"></legacyDrawing>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff

def test_write_with_tab_color(worksheet, write_worksheet):
    ws = worksheet
    ws.sheet_properties.tabColor = "F0F0F0"
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr>
        <outlinePr summaryRight="1" summaryBelow="1"/>
        <pageSetUpPr/>
        <tabColor rgb="00F0F0F0"/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection sqref="A1" activeCell="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff


def test_write_with_fit_to_page(worksheet, write_worksheet):
    ws = worksheet
    ws.page_setup.fitToPage = True
    ws.page_setup.autoPageBreaks = False
    xml = write_worksheet(ws, None)
    expected = """
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <sheetPr>
        <outlinePr summaryRight="1" summaryBelow="1"/>
        <pageSetUpPr fitToPage="1" autoPageBreaks="0"/>
      </sheetPr>
      <dimension ref="A1:A1"/>
      <sheetViews>
        <sheetView workbookViewId="0">
          <selection sqref="A1" activeCell="A1"/>
        </sheetView>
      </sheetViews>
      <sheetFormatPr baseColWidth="10" defaultRowHeight="15"/>
      <sheetData/>
      <pageMargins left="0.75" right="0.75" top="1" bottom="1" header="0.5" footer="0.5"/>
    </worksheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
