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

print(f"Old input value: {mathcad_worksheet.get_input('input_1')}")  # only here for debugging
mathcad_worksheet.set_real_input("input_1", 2)  # change the value in Mathcad programmatically
print(f"New input value: {mathcad_worksheet.get_input('input_1')}")  # only here for debugging

mathcad_worksheet.set_real_input("input_2", 4)  # change the value in Mathcad programmatically


value, units, error_code = mathcad_worksheet.get_real_output("output")
if error_code == 0:  # Good practice to check for errors when you retreive a value
    print(value, units)
else:
    raise ValueError
