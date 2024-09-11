# -*- coding: utf-8 -*-
"""
examples.py
~~~~~~~~~~~~~~
MathcadPy
Copyright 2020 Matt Woodhead
"""

from pprint import pprint
from pathlib import Path
from MathcadPy import Mathcad

# Try and import numpy - not a prerequisite so fail gracefully if not installed
try:
    import numpy as np
    NUMPY_IMPORTED = True
except ModuleNotFoundError:
    NUMPY_IMPORTED = False

# First, create a new instance of the Mathcad class (the window can be set to hidden with visible=False)
mathcad_app = Mathcad(visible=True)

# Check the mathcad version and print to the console
print(f"Mathcad version: {mathcad_app.version}")

# Open a worksheet in Mathcad and set it as the active worksheet
test_ws = mathcad_app.open(Path.cwd() / "Test" / "test.mcdx")
test_ws.activate()
print(f"Worksheet input names: {test_ws.inputs()}")
print(f"Worksheet output names: {test_ws.outputs()}")


# Set some input values
test_ws.set_real_input("real_input_test", 11)
test_ws.set_real_input("real_input_with_units_test", 3, "mm", preserve_worksheet_units=False)
test_ws.set_string_input("string_input_test", "string from python script!")

if NUMPY_IMPORTED:
    matrix_to_send = np.array([1, 2, 3, 4]).reshape((2, 2))
else:
    matrix_to_send = [[1, 2],
                      [3, 4],
                      ]
test_ws.set_matrix_input("matrix_input_test", matrix_to_send, "s", preserve_worksheet_units=False)


# Fetch some output values
value, units, error_code = test_ws.get_real_output("real_output_test")
print(value, units)
value, units, error_code = test_ws.get_real_output("real_output_test", units="in")  # get the previous result, but this time in inches
if error_code == 0:  # Good practice to check for errors when you request specific units
    print(value, units)

matrix, units, error = test_ws.get_matrix_output("matrix_output_test")
print(matrix, units)

# Save the worksheet under new filenames, and then close it
try:
    print("Trying to save MCDX")
    test_ws.save_as(Path.cwd() / "Test" / "test_output.mcdx")

    print("Trying to save RTF")
    test_ws.save_as(Path.cwd() / "Test" / "test_output.rtf")

    print("Trying to save XPS")
    test_ws.save_as(Path.cwd() / "Test" / "test_output.xps")

    if mathcad_app.version_major_int > 4:
        print("Trying to save PDF")
        test_ws.save_as(Path.cwd() / "Test" / "test_output.xps")
except Exception as exc:
    print(exc)
    pass

test_ws.close()
