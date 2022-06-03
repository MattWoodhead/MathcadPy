# -*- coding: utf-8 -*-
"""
examples.py
~~~~~~~~~~~~~~
MathcadPy
Copyright 2022 Matt Woodhead
"""

# Standard Library Imports
from pathlib import Path

# External library Imports
from MathcadPy import Mathcad

mathcad_app = Mathcad()  # creates an instance of the Mathcad class - this object represents the Mathcad application

mathcad_worksheet = mathcad_app.open(Path.cwd() / "simple_example_complete.mcdx")

print(mathcad_worksheet.inputs())

print(mathcad_worksheet.outputs())
