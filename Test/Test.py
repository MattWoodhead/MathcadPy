# -*- coding: utf-8 -*-
"""
MathcadPy.py

Author: MattWoodhead
"""
import numpy as np

array = np.array([[1, 2], [3, 4]])

height, width = array.shape

print(isinstance(array, np.ndarray))

print(f"Height = {height}\nWidth = {width}")

