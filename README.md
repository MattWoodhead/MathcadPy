# MathcadPy
A Python wrapper for the Mathcad Prime automation API

### Requirements
- [Mathcad Prime 3+](https://www.mathcad.com/)
- [win32com](https://github.com/mhammond/pywin32)

### Features
MathcadPy is a python wrapper for the Mathcad Prime COM automation API. This allows a python script to interact with a Mathcad session, select worksheets that are already open or open worksheets from files, and send values to designated inputs and fetch them from designated outputs. The python script also has access to the worksheet units, and can even demand results in different units to those specified in a worksheet (e.g. calculation may output a value in meters, but the script wants the value in inches).

### Usage
Install using

    pip install MathcadPy
 
See [documentation/Getting_Started.md](https://github.com/MattWoodhead/MathcadPy/blob/master/documentation/Getting_Started.md) to get up and running with MathcadPy.
See examples.py for some simple function examples including matrix operations.

### Todo
- [ ] Wiki documenting the functions

### licensing and credits
Author: Matt Woodhead

MathcadPy is licensed under the GPLv3
```
This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <http://www.gnu.org/licenses/>.

For the full license, see the LICENSE file.
```
