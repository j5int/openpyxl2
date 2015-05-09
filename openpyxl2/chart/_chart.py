from __future__ import absolute_import

from openpyxl2.descriptors import Typed, Integer
from openpyxl2.descriptors.serialisable import Serialisable
from openpyxl2.xml.constants import CHART_NS

from .data_source import AxDataSource, NumRef
from .legend import Legend
from .reference import Reference
from .series_factory import SeriesFactory
from .series import attribute_mapping


class AxId(Serialisable):

    val = Integer()

    def __init__(self, val):
        self.val = val


class ChartBase(Serialisable):

    """
    Base class for all charts
    """

    legend = Typed(expected_type=Legend, allow_none=True)

    _series_type = ""
    ser = ()
    anchor = "E15" # default anchor position
    _shapes = ()

    __elements__ = ()

    def __init__(self):
        self._charts = [self]
        self.legend = Legend()


    def __iadd__(self, other):
        """
        Combine the chart with another one
        """
        if not isinstance(other, ChartBase):
            raise TypeError("Only other charts can be added")
        self._charts.append(other)
        return self


    def to_tree(self, tagname=None, idx=None):
        if self.ser is not None:
            for s in self.ser:
                s.__elements__ = attribute_mapping[self._series_type]
        return super(ChartBase, self).to_tree(tagname, idx)


    def _write(self):
        from .chartspace import ChartSpace, ChartContainer, PlotArea
        plot = PlotArea()
        for chart in self._charts:
            setattr(plot, chart.tagname, chart)
        for axis in ("x_axis", "y_axis", 'z_axis'):
            axis = getattr(self, axis, None)
            if axis is None:
                continue
            setattr(plot, axis.tagname, axis)
        container = ChartContainer(plotArea=plot, legend=self.legend)
        cs = ChartSpace(chart=container)
        tree = cs.to_tree()
        tree.set("xmlns", CHART_NS)
        return tree


    @property
    def axId(self):
        x = getattr(self, "x_axis", None)
        y = getattr(self, "y_axis", None)
        z = getattr(self, "z_axis", None)
        ids = [AxId(axis.axId) for axis in (x, y, z) if axis]

        return ids


    def set_categories(self, labels):
        """
        Set the categories / x-axis values
        """
        if not isinstance(labels, Reference):
            labels = Reference(range_string=labels)
        for s in self.ser:
            s.cat = AxDataSource(numRef=NumRef(f=labels))


    def add_data(self, data, from_rows=False, title_from_data=False, labels=None):
        """
        Add a range of data in a single pass.
        The default is to treat each column as a data series.
        """
        if not isinstance(data, Reference):
            data = Reference(range_string=data)
        if from_rows:
            values = data.rows
        else:
            values = data.cols
        for v in values:
            range_string = "{0}!{1}:{2}".format(data.sheetname, v[0], v[-1])
            series = SeriesFactory(range_string, title_from_data=title_from_data)
            self.ser.append(series)
        if labels is not None:
            self.set_categories(labels)

