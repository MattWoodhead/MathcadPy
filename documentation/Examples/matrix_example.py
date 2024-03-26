# -*- coding: utf-8 -*-
"""
matrix_example.py
~~~~~~~~~~~~~~
MathcadPy
Copyright 2024 Matt Woodhead
"""

# Standard Library Imports
from pathlib import Path

# External library Imports
from MathcadPy import Mathcad
import numpy as np

# create an instance of the Mathcad class - this object represents the Mathcad application
mathcad_app = Mathcad()
mathcad_app.Visible = True

mathcad_worksheet = mathcad_app.open(Path.cwd() / "matrix_example.mcdx")

print(mathcad_worksheet.inputs())

print(mathcad_worksheet.outputs())

# print the value of the input before we interact with it - for debugging purposes only
print(f"Old input 1 value: {mathcad_worksheet.get_matrix_input('input_1')}")
mathcad_worksheet.set_matrix_input(
    "input_1",
    [[1, 3],
     [2, 4],
     ],
)  # change the value in Mathcad programmatically
# print the value of the input after we interact with it - for debugging purposes only
print(f"New input 1 value: {mathcad_worksheet.get_matrix_input('input_1')}")

# MathcadPy can also use Numpy arrays
matrix_2 = np.array(
    [[1, 0],
     [0, 1],
     ],
)
mathcad_worksheet.set_matrix_input("input_2", matrix_2)

# fetch the output value now we have changed the inputs
value, units, error_code = mathcad_worksheet.get_matrix_output("output")
if error_code == 0:  # Good practice to check for errors when you retreive a value
    print(f"Output value: {value} {units}")
else:
    raise ValueError

mathcad_worksheet.save_as(Path.cwd() / "matrix_example_output.mcdx")
mathcad_app.quit()
