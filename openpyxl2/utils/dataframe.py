from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from itertools import accumulate
import operator
import numpy
from pandas import Timestamp


def dataframe_to_rows(df, index=True, header=True):
    """
    Convert a Pandas dataframe into something suitable for passing into a worksheet
    """
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
        values = list(df.columns.values)
        if df.columns.dtype.type == numpy.datetime64:
            values = [Timestamp(v) for v in values]
        yield [None]*index + values

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
