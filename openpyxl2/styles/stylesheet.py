from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.descriptors import (
    Alias,
    Typed,
    Sequence
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.utils.indexed_list import IndexedList

from .colors import ColorList, COLOR_INDEX
from .differential import DifferentialStyleList
from .table import TableStyleList
from .borders import BorderList
from .fills import FillList
from .fonts import FontList
from .numbers import NumberFormatList
from .alignment import Alignment
from .protection import Protection
from .named_styles import NamedCellStyleList, NamedStyle
from .cell_style import CellStyleList


class Stylesheet(Serialisable):

    tagname = "stylesheet"

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
        self.cellXfs = cellXfs
        self.cellStyles = cellStyles
        if dxfs is None:
            dxfs = DifferentialStyleList()
        self.dxfs = dxfs
        self.tableStyles = tableStyles
        self.colors = colors


    @classmethod
    def from_tree(cls, node):
        # strip all attribs
        attrs = dict(node.attrib)
        for k in attrs:
            del node.attrib[k]
        self = super(Stylesheet, cls).from_tree(node)
        # convert objects where necessary
        cell_styles = self.cellXfs
        self.cell_styles = cell_styles._to_array()
        self.alignments = cell_styles.alignments
        self.protections = cell_styles.prots
        self.named_styles =  self._merge_named_styles()
        return self


    def _merge_named_styles(self):
        """
        Merge named style names "cellStyles" with their associated styles "cellStyleXfs"
        """
        named_styles = {}
        for name, style in self.cellStyles.names.items():
            xf = self.cellStyleXfs[style.xfId]
            style.font = self.fonts[xf.fontId]
            style.fill = self.fills[xf.fillId]
            style.border = self.borders[xf.borderId]
            if xf.numFmtId > 164:
                style.number_format = self.numFmts[xf.numFmtId]
            if xf.alignment:
                style.alignment = xf.alignment
            if xf.protection:
                style.protection = xf.protection
            named_styles[name] = style
        return named_styles


    @property
    def color_index(self):
        if self.colors:
            return self.colors.index
        return COLOR_INDEX


    @property
    def font_list(self):
        return IndexedList(self.fonts.font)


    @property
    def fill_list(self):
        return IndexedList(self.fills.fill)


    @property
    def border_list(self):
        return IndexedList(self.borders.border)


    @property
    def differential_list(self):
        return IndexedList(self.dxfs.dxf)

    @property
    def number_formats(self):
        return IndexedList(self.numFmts.numFmt)


from openpyxl2.xml.constants import ARC_STYLE
from openpyxl2.xml.functions import fromstring

def read_stylesheet(archive):
    try:
        src = archive.read(ARC_STYLE)
    except KeyError:
        return
    node = fromstring(src)
    return Stylesheet.from_tree(node)
