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


# First, create a new instance of the Mathcad class (the window can be set to hidden with visible=False)
mathcad_app = Mathcad(visible=True)

# Open a worksheet in Mathcad and set it as the active worksheet
test_ws = mathcad_app.open(Path.cwd() / "test_units.mcdx")
test_ws.activate()

print("real_output_test:")
# Fetch the output value
value_1_a, units_1_a, error_code = test_ws.get_real_output("real_output_test")

# get the same value as the previous result, but this time in foot-pounds torque
value_1_b, units_1_b, error_code = test_ws.get_real_output("real_output_test", units="ft*kip")

value_1_c, units_1_c, error_code = test_ws._get_output("real_output_test")
print(f"{value_1_a} {units_1_a}, {value_1_b} {units_1_b}, {value_1_c} {units_1_c}")

print("real_output_test2:")
# Fetch the output value
value_2_a, units_2_a, error_code = test_ws.get_real_output("real_output_test2")

# get the same value as the previous result, but this time in foot-pounds torque
value_2_b, units_2_b, error_code = test_ws.get_real_output("real_output_test2", units="ft*kip")

value_2_c, units_2_c, error_code = test_ws._get_output("real_output_test2")
print(f"{value_2_a} {units_2_a}, {value_2_b} {units_2_b}, {value_2_c} {units_2_c}")

print("real_output_test3:")
# Fetch the output value
value_3_a, units_3_a, error_code = test_ws.get_real_output("real_output_test3")

# get the same value as the previous result, but this time in foot-pounds torque
value_3_b, units_3_b, error_code = test_ws.get_real_output("real_output_test3", units="ft*kip")

value_3_c, units_3_c, error_code = test_ws._get_output("real_output_test3")
print(f"{value_3_a} {units_3_a}, {value_3_b} {units_3_b}, {value_3_c} {units_3_c}")

#test_ws.close()
