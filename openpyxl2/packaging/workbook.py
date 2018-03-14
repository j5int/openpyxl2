from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""
OO-based reader
"""

import posixpath
from warnings import warn

from openpyxl2.xml.functions import fromstring

from openpyxl2.packaging.relationship import (
    get_dependents,
    get_rels_path,
    get_rel,
)
from openpyxl2.packaging.manifest import Manifest
from openpyxl2.workbook.workbook import Workbook
from openpyxl2.workbook.defined_name import (
    _unpack_print_area,
    _unpack_print_titles,
)
from openpyxl2.workbook.external_link.external import read_external_link
from openpyxl2.pivot.cache import CacheDefinition
from openpyxl2.pivot.record import RecordList

from openpyxl2.utils.datetime import CALENDAR_MAC_1904

