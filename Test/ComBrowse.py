# -*- coding: utf-8 -*-
"""
ComBrowse.py

Author: MattWoodhead
"""
import win32com.client.combrowse as cb

try:
    print("opening COM object browser")
    cb.main()
except:
    print("error opening COM object browser")
