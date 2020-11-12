# -*- coding: utf-8 -*-
"""
__init__.py
~~~~~~~~~~~~~~
MathcadPy

Copyright 2020 Matt Woodhead

Requirements:

Mathcad Prime (https://www.mathcad.com)
PyWin32 (https://github.com/mhammond/pywin32)
"""

from ._application import *
from . import __version__

__author__ = __version__.__author__
__version__ = __version__.__version__
