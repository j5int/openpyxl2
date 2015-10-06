from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Alias,
    Typed,
    Sequence
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.utils.indexed_list import IndexedList
from openpyxl2.xml.constants import ARC_STYLE, SHEET_MAIN_NS
from openpyxl2.xml.functions import fromstring

from .colors import ColorList, COLOR_INDEX
from .differential import DifferentialStyleList
from .table import TableStyleList
from .borders import BorderList
from .fills import FillList
from .fonts import FontList
from .numbers import (
    NumberFormatList,
    BUILTIN_FORMATS,
    BUILTIN_FORMATS_REVERSE
)
from .alignment import Alignment
from .protection import Protection
from .named_styles import (
    NamedCellStyleList,
    NamedStyle,
    NamedCellStyle
)
from .cell_style import CellStyleList, CellStyle


class Stylesheet(Serialisable):

    tagname = "styleSheet"

    numFmts = Typed(expected_type=NumberFormatList, allow_none=True)
    fonts = Typed(expected_type=FontList, allow_none=True)
    fills = Typed(expected_type=FillList, allow_none=True)
    borders = Typed(expected_type=BorderList, allow_none=True)
    cellStyleXfs = Typed(expected_type=CellStyleList, allow_none=True)
    cellXfs = Typed(expected_type=CellStyleList, allow_none=True)
    cellStyles = Typed(expected_type=NamedCellStyleList, allow_none=True)
    dxfs = Typed(expected_type=DifferentialStyleList, allow_none=True)
    tableStyles = Typed(expected_type=TableStyleList, allow_none=True)
    colors = Typed(expected_type=ColorList, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs',
                    'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors')

    def __init__(self,
                 numFmts=None,
                 fonts=None,
                 fills=None,
                 borders=None,
                 cellStyleXfs=None,
                 cellXfs=None,
                 cellStyles=None,
                 dxfs=None,
                 tableStyles=None,
                 colors=None,
                 extLst=None,
                ):
        if numFmts is None:
            numFmts = NumberFormatList()
        self.numFmts = numFmts
        if fonts is None:
            fonts = FontList()
        self.fonts = fonts
        if fills is None:
            fills = FillList()
        self.fills = fills
        if borders is None:
            borders = BorderList()
        self.borders = borders
        self.cellStyleXfs = cellStyleXfs
        if cellXfs is None:
            cellXfs = CellStyleList()
        self.cellXfs = cellXfs
        if cellStyles is None:
            cellStyles = NamedCellStyleList()
        self.cellStyles = cellStyles
        if dxfs is None:
            dxfs = DifferentialStyleList()
        self.dxfs = dxfs
        self.tableStyles = tableStyles
        self.colors = colors

        self.cell_styles = self.cellXfs._to_array()
        self.alignments = self.cellXfs.alignments
        self.protections = self.cellXfs.prots
        self._normalise_numbers()
        self.named_styles =  self._merge_named_styles()


    @classmethod
    def from_tree(cls, node):
        # strip all attribs
        attrs = dict(node.attrib)
        for k in attrs:
            del node.attrib[k]
        return  super(Stylesheet, cls).from_tree(node)


    def _merge_named_styles(self):
        """
        Merge named style names "cellStyles" with their associated styles "cellStyleXfs"
        """
        named_styles = {}
        custom = self.custom_formats
        formats = self.number_formats
        for name, style in self.cellStyles.names.items():
            xf = self.cellStyleXfs[style.xfId]
            style.font = self.fonts[xf.fontId]
            style.fill = self.fills[xf.fillId]
            style.border = self.borders[xf.borderId]
            if xf.numFmtId in custom:
                fmt = custom[xf.numFmtId]
                style.numFmtId = formats.index(fmt) + 164
            if xf.alignment:
                style.alignment = xf.alignment
            if xf.protection:
                style.protection = xf.protection
            named_styles[name] = style
        return named_styles


    def _split_named_styles(self, wb):
        """
        Convert NamedStyle into separate CellStyle and Xf objects
        """
        names = []
        xfs = []
        for idx, style in enumerate(wb._named_styles.values()):
            name = NamedCellStyle(
                name=style.name,
                builtinId=style.builtinId,
                hidden=style.hidden,
                xfId = idx
            )
            names.append(name)

            xf = CellStyle()
            xf.fontId =  wb._fonts.add(style.font)
            xf.borderId = wb._borders.add(style.border)
            xf.fillId =  wb._fills.add(style.fill)
            fmt = style.number_format
            if fmt in BUILTIN_FORMATS_REVERSE:
                fmt = BUILTIN_FORMATS_REVERSE[fmt]
            else:
                fmt = wb._number_formats.add(style.number_format) + 164
            xf.numFmtId = fmt
            xfs.append(xf)

        self.cellStyles.cellStyle = names
        self.cellStyleXfs = CellStyleList(xf=xfs)


    @property
    def number_formats(self):
        fmts = [n.formatCode for n in self.numFmts.numFmt]
        return IndexedList(fmts)


    @property
    def custom_formats(self):
        return dict([(n.numFmtId, n.formatCode) for n in self.numFmts.numFmt])


    def _normalise_numbers(self):
        """
        Rebase numFmtIds with a floor of 164
        """
        custom = self.custom_formats
        formats = self.number_formats
        for style in self.cell_styles:
            if style.numFmtId in custom:
                fmt = custom[style.numFmtId]
                style.numFmtId = formats.index(fmt) + 164


def apply_stylesheet(archive, wb):
    """
    Add styles to workbook if present
    """
    try:
        src = archive.read(ARC_STYLE)
    except KeyError:
        return wb
    node = fromstring(src)
    stylesheet = Stylesheet.from_tree(node)

    wb._cell_styles = stylesheet.cell_styles
    wb._named_styles = stylesheet.named_styles
    wb._borders = IndexedList(stylesheet.borders.border)
    wb._fonts = IndexedList(stylesheet.fonts.font)
    wb._fills = IndexedList(stylesheet.fills.fill)
    wb._differential_styles = IndexedList(stylesheet.dxfs.dxf)
    wb._number_formats = stylesheet.number_formats
    wb._protections = stylesheet.protections
    wb._alignments = stylesheet.alignments
    if stylesheet.colors is not None:
        wb._colors = stylesheet.colors.index


def write_stylesheet(wb):
    stylesheet = Stylesheet()
    stylesheet.fonts.font = wb._fonts
    stylesheet.fills.fill = wb._fills
    stylesheet.borders.border = wb._borders
    stylesheet.dxfs.dxf = wb._differential_styles

    from .numbers import NumberFormat
    fmts = []
    for idx, code in enumerate(wb._number_formats, 164):
        fmt = NumberFormat(idx, code)
        fmts.append(fmt)

    stylesheet.numFmts.numFmt = fmts

    xfs = []
    for style in wb._cell_styles:
        xf = CellStyle.from_array(style)

        if style.applyAlignment:
            xf.alignment = wb._alignments[style.alignmentId]
            xf.applyAlignment = True

        if style.applyProtection:
            xf.protection = wb._protections[style.protectionId]
        xfs.append(xf)
    stylesheet.cellXfs = CellStyleList(xf=xfs)

    stylesheet._split_named_styles(wb)
    stylesheet.tableStyles = TableStyleList()

    tree = stylesheet.to_tree()
    tree.set("xmlns", SHEET_MAIN_NS)
    return tree
