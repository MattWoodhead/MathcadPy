# -*- coding: utf-8 -*-
"""
__init__.py
~~~~~~~~~~~~~~
MathcadPy
https://github.com/MattWoodhead/MathcadPy

Copyright 2023 Matt Woodhead

Requirements:

Mathcad Prime ( https://www.mathcad.com )
PyWin32 ( https://github.com/mhammond/pywin32 )
"""

from ._application import Mathcad, Worksheet
from ._application import _matrix_to_array, _array_check
from . import __version__

__author__ = __version__.__author__
__version__ = __version__.__version__
