from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

import operator
from openpyxl2.compat import accumulate


def dataframe_to_rows(df, index=True, header=True):
    """
    Convert a Pandas dataframe into something suitable for passing into a worksheet
    """
    import numpy
    from pandas import Timestamp
    blocks = df._data.blocks
    ncols = sum(b.shape[0] for b in blocks)
    data = [None] * ncols

    for b in blocks:
        values = b.values

        if b.dtype.type == numpy.datetime64:
            values = numpy.array([Timestamp(v) for v in values.ravel()])
            values = values.reshape(b.shape)

        result = values.tolist()

        for col_loc, col in zip(b.mgr_locs, result):
            data[col_loc] = col

    if header:
        if df.columns.nlevels > 1:
            rows = expand_levels(df.columns.levels)
        else:
            rows = [list(df.columns.values)]
        for row in rows:
            n = []
            for v in row:
                if isinstance(v, numpy.datetime64):
                    v = Timestamp(v)
                n.append(v)
            row = n
            yield [None]*index + row

    for idx, v in enumerate(df.index):
        yield [v]*index + [data[j][idx] for j in range(ncols)]


def expand_levels(levels):
    """
    Multiindexes need expanding so that subtitles repeat
    """
    widths = (len(s) for s in levels)
    widths = list(accumulate(widths, operator.mul))
    size = max(widths)

    for level, width in zip(levels, widths):
        padding = int(size/width) # how wide a title should be
        repeat = int(width/len(level)) # how often a title is repeated
        row = []
        for v in level:
            title = [None]*padding
            title[0] = v
            row.extend(title)
        row = row*repeat
        yield row
