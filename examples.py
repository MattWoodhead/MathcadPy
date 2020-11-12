# -*- coding: utf-8 -*-
"""
Created on Wed Oct 21 21:17:05 2020

@author: matth
"""

from pathlib import Path
from MathcadPy import Mathcad, Matrix


mathcad_app = Mathcad(visible=True)

print(mathcad_app.version)

print(mathcad_app.open_worksheets)

test_ws = mathcad_app.open_worksheet(Path.cwd() / "Test" / "test.mcdx")

print(mathcad_app.open_worksheets)

test_ws.set_real_input("in1", 10, "mm")
out2 = test_ws.get_real_output("out2")

print(out2)


matrix, units, error = test_ws.get_matrix_output("out999")
print(matrix)

#matrix.
