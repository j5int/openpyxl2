from __future__ import absolute_import
# Copyright (c) 2010-2018 openpyxl

from io import BytesIO
from warnings import warn

from openpyxl2.xml.functions import fromstring
from openpyxl2.xml.constants import IMAGE_NS
from openpyxl2.packaging.relationship import get_rel, get_rels_path, get_dependents
from openpyxl2.drawing.spreadsheet_drawing import SpreadsheetDrawing
from openpyxl2.drawing.image import Image
from openpyxl2.chart.chartspace import ChartSpace
from openpyxl2.chart.reader import read_chart


def find_images(archive, path):
    """
    Given the path to a drawing file extract charts and images

    Ingore errors due to unsupported parts of DrawingML
    """

    src = archive.read(path)
    tree = fromstring(src)
    try:
        drawing = SpreadsheetDrawing.from_tree(tree)
    except TypeError:
        warn("DrawingML support is incomplete and limited to charts and images only. Shapes and drawings will be lost.")
        return [], []

    rels_path = get_rels_path(path)
    deps = []
    if rels_path in archive.namelist():
        deps = get_dependents(archive, rels_path)

    charts = []
    for rel in drawing._chart_rels:
        cs = get_rel(archive, deps, rel.id, ChartSpace)
        chart = read_chart(cs)
        chart.anchor = rel.anchor
        charts.append(chart)

    images = []
    for rel in drawing._blip_rels:
        dep = deps[rel.embed]
        if dep.Type == IMAGE_NS:
            image = Image(BytesIO(archive.read(dep.target)))
            image.anchor = rel.anchor
            images.append(image)
    return charts, images