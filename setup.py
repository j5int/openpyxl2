#!/usr/bin/env python

"""Setup script for packaging openpyxl2.

Requires setuptools.

To build the setuptools egg use
    python setup.py bdist_egg
and either upload it to the PyPI with:
    python setup.py upload
or upload to your own server and register the release with PyPI:
    python setup.py register

A source distribution (.zip) can be built with
    python setup.py sdist --format=zip

That uses the manifest.in file for data files rather than searching for
them here.

"""

import sys
import os
import warnings
if sys.version_info < (2, 6):
    raise Exception("Python >= 2.6 is required.")
elif sys.version_info[:2] == (3, 2):
    warnings.warn("Python 3.2 is no longer officially supported")

from setuptools import setup, Extension, find_packages
import re

here = os.path.abspath(os.path.dirname(__file__))
try:
    with open(os.path.join(here, 'README.rst')) as f:
        README = f.read()
except IOError:
    README = ''


__author__ = 'See AUTHORS'
__license__ = 'MIT/Expat'
__author_email__ = 'eric.gazoni@gmail.com'
__maintainer_email__ = 'openpyxl2-users@googlegroups.com'
__url__ = 'http://openpyxl2.readthedocs.org'


def get_version():
    f = open(os.path.join(here, 'openpyxl2', '__init__.py'))
    version_file = f.read()
    f.close()
    version_match = re.search(r"^__version__ = ['\"]([^'\"]*)['\"]",
                              version_file, re.M)
    if version_match:
        return version_match.group(1)
    raise RuntimeError("Unable to find version string.")

setup(name='openpyxl2',
      packages=find_packages(),
      # metadata
      version=get_version(),
      description="A Python library to read/write Excel 2007 xlsx/xlsm files",
      long_description=README,
      author=__author__,
      author_email=__author_email__,
      url=__url__,
      license=__license__,
      requires=[
                'python (>=2.6.0)',
                ],
      install_requires=[
                        'jdcal',
                        ],
      classifiers=[
                   'Development Status :: 4 - Beta',
                   'Development Status :: 5 - Production/Stable'
                   'Operating System :: MacOS :: MacOS X',
                   'Operating System :: Microsoft :: Windows',
                   'Operating System :: POSIX',
                   'License :: OSI Approved :: MIT License',
                   'Programming Language :: Python',
                   'Programming Language :: Python :: 2.6',
                   'Programming Language :: Python :: 2.7',
                   'Programming Language :: Python :: 3.3',
                   'Programming Language :: Python :: 3.4',
                   ],
      )
