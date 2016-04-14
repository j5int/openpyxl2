Charts
======

.. warning::

    Openpyxl currently supports chart creation within a worksheet only. Charts in
    existing workbooks will be lost.


Chart types
-----------

The following charts are available:

.. toctree::

    area
    bar
    bubble
    line
    scatter
    pie
    doughnut
    radar
    stock
    surface


Creating a chart
----------------

Charts are composed of at least one series of one or more data points. Series
themselves are comprised of references to cell ranges.

.. :: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook()
>>> ws = wb.active
>>> for i in range(10):
...     ws.append([i])
>>>
>>> from openpyxl2[.]chart import BarChart, Reference, Series
>>> values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
>>> chart = BarChart()
>>> chart.add_data(values)
>>> ws.add_chart(chart, "E15")
>>> wb.save("SampleChart.xlsx")


By default the top-left corner of a chart is anchored to cell E15 and the
size is 15 x 7.5 cm (approximately 5 columns by 14 rows). This can be changed
by setting the `anchor`, `width` and `height` properties of the chart. The
actual size will depend on operating system and device. Other anchors are
possible see :mod:`openpyxl2[.]drawing.spreadsheet_drawing` for further information.


Working with axes
-----------------

.. toctree::

    limits_and_scaling
    secondary


Change the chart layout
-----------------------

.. toctree::

    chart_layout


Styling charts
--------------

.. toctree::

    pattern


Advanced charts
---------------

Charts can be combined to create new charts:

.. toctree::

    gauge


Using chartsheets
-----------------

Charts can be added to special worksheets called chartsheets:

.. toctree::

    chartsheet
