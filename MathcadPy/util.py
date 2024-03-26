# -*- coding: utf-8 -*-
"""
util.py
~~~~~~~~~~~~~~
MathcadPy
Copyright 2024 Matt Woodhead
"""

import subprocess


def list_installations():
    """returns the install location, program name, and version of all Mathcad instances"""
    # TODO - tidy format
    data = subprocess.check_output("wmic product where \"Name like '%mathcad%'\" get Name, Version, InstallLocation")
    data = [dt.strip() for dt in str(data).split("\\r\\r\\n") if "Mathcad" in dt]
    print("Warning: The output format of mathcadpy.util.list_installations() is subject to change")
    return data


if __name__ == "__main__":
    print(list_installations())
