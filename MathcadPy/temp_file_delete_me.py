# -*- coding: utf-8 -*-
"""
Created on Sat Apr 23 22:09:27 2022

@author: matth
"""

import subprocess

#Data = subprocess.check_output('wmic product get name')
Data = subprocess.check_output('wmic product where "Name like \'%mathcad%\'" get Name, Version, installlocation')
print(Data)
Data = [l.strip() for l in str(Data).split("\\r\\r\\n") if "Mathcad" in l]
print(len(Data))

print("\n".join(Data))
