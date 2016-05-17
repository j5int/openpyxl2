from __future__ import absolute_import
# Copyright (c) 2010-2016 openpyxl

from abc import abstractproperty
from openpyxl2.compat.abc import ABC


class ISerialisableFile(ABC):

    """
    Interface for Serialisable classes that represent files in the archive
    """


    @abstractproperty
    def _id(self):
        """
        The id linking the object to its parent
        """
        pass


    @abstractproperty
    def _path(self):
        """
        File path in the archive
        """
        pass


    @abstractproperty
    def _namespace(self):
        """
        Qualified namespace when serialised
        """
        pass


    @abstractproperty
    def _type(self):
        """
        The content type for the manifest
        """


    @abstractproperty
    def _rel_type(self):
        """
        The content type for relationships
        """
