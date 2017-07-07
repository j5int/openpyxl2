from __future__ import absolute_import
# Copyright (c) 2010-2017 openpyxl

from collections import OrderedDict

from openpyxl2.descriptors import (
    Bool,
    String,
    Sequence,
    Alias,
)
from openpyxl2.descriptors.excel import ExtensionList
from openpyxl2.descriptors.serialisable import Serialisable

from .rule import Rule


class ConditionalFormatting(Serialisable):

    tagname = "conditionalFormatting"

    sqref = String(allow_none=True)
    pivot = Bool(allow_none=True)
    cfRule = Sequence(expected_type=Rule)
    rules = Alias("cfRule")


    def __init__(self, sqref=None, pivot=None, cfRule=(), extLst=None):
        self.sqref = sqref
        self.pivot = pivot
        self.cfRule = cfRule


    def __hash__(self):
        return hash(self.sqref)


    def __repr__(self):
        return repr(self.sqref)


class ConditionalFormattingList(object):
    """Conditional formatting rules."""


    def __init__(self):
        self._cf_rules = OrderedDict()
        self.max_priority = 0


    def add(self, range_string, cfRule):
        """Add a rule such as ColorScaleRule, FormulaRule or CellIsRule

         The priority will be added automatically.
        """
        cf = ConditionalFormatting(range_string)
        if not isinstance(cfRule, Rule):
            raise ValueError("Only instances of openpyxl2.formatting.rule.Rule may be added")
        rule = cfRule
        self.max_priority += 1
        if not rule.priority:
            rule.priority = self.max_priority

        self._cf_rules.setdefault(cf, []).append(rule)


    def __bool__(self):
        return bool(self._cf_rules)

    __nonzero = __bool__


    def __len__(self):
        return len(self._cf_rules)


    def __iter__(self):
        for cf, rules in self._cf_rules.items():
            cf.rules = rules
            yield cf


    def __getitem__(self, key):
        """
        Get the rules for a cell range
        """
        if isinstance(key, str):
            key = ConditionalFormatting(key)
        return self._cf_rules[key]


    def __setitem__(self, key, rule):
        """
        Add a rule for a cell range
        """
        self.add(key, rule)
