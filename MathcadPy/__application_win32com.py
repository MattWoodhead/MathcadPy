# -*- coding: utf-8 -*-
"""
MathcadPy
|
|- Application.py

Author: MattWoodhead

Requirements:

Mathcad Prime
comtypes (https://github.com/enthought/comtypes)

"""

import os
import win32com.client as w32c
import numpy as np


class Mathcad():
    """ Mathcad application object """

    def __init__(self, visible=False):
        print("Loading Mathcad")
        try:
            self.__mcadapp = w32c.Dispatch("MathcadPrime.Application")
            self.version = self.__mcadapp.GetVersion()  # Fetches Mathcad version
            if visible is False:
                self.__mcadapp.Visible = False
            else:
                self.__mcadapp.Visible = True
        except:
            print("Could not locate the Mathcad Automation API")

    def activate(self):
        """ Activate the Mathcad window. If visible, this maximises Mathcad"""
        self.__mcadapp.Activate()

    def active_sheet(self):
        """ Returns the active worksheet name """
        name = self.__mcadapp.ActiveWorksheet.Name
        if name == "":
            return None  # Returns none if the current worksheet not saved
        return name

    def worksheet_names(self):
        """ lists worksheets open in the Mathcad instance """
        worksheets = []
        for i in range(self.__mcadapp.Worksheets.Count):  # no. of open sheets
            worksheets.append(self.__mcadapp.Worksheets.Item(i).Name)
        return worksheets  # Returns a list of open worksheet filenames

    def worksheet_paths(self):
        """ lists worksheets open in the Mathcad instance """
        worksheets = []
        for i in range(self.__mcadapp.Worksheets.Count):  # no. of open sheets
            worksheets.append(self.__mcadapp.Worksheets.Item(i).FullName)
        return worksheets  # Returns a list of open worksheet filenames

    def close_all(self, save_option="Discard"):
        """ Closes all worksheets. Can specify save options before closing """
        if save_option in ["Discard", 2]:
            self.__mcadapp.CloseAll(2)
        elif save_option in ["Prompt", 1]:
            self.__mcadapp.CloseAll(1)
        elif save_option in ["Save", 0]:
            self.__mcadapp.CloseAll(0)
        else:
            print("incorrect save argument specified")


class Worksheet():
    """ Mathcad Worksheet object

    Either a filepath for a mathcad file can be supplied, or the
    filepath can be set to None (or similar) and the optional
    open_sheet_name argument can be used
    """

    def __init__(self, filepath, open_sheet_name=None):
        self.__mcadapp = w32c.Dispatch("MathcadPrime.Application")
        self.__ws_at_init = {}
        for i in range(self.__mcadapp.Worksheets.Count):
            self.__ws_at_init[self.__mcadapp.Worksheets.Item(i).Name] = \
            (self.__mcadapp.Worksheets.Item(i).FullName,
             self.__mcadapp.Worksheets.Item(i))
        if open_sheet_name is not None:
            for n, (path, __mcobj) in self.__ws_at_init.items():
                if open_sheet_name == n:
                    self.__mcadapp.Open(path)
                    # Doesn't really open as it is already open
                    # @TODO - change to activate worksheet by same name
                    self.__obj = __mcobj
                    #self.__obj = self.__mcadapp.ActiveWorksheet.Name  # Fetches COM worksheet object
                    self.Name = self.__obj.Name
                    break
            else:
                print(f"open_sheet_name={open_sheet_name} does not match the name of any open worksheets")
        if filepath is not None:
            if os.path.isfile(filepath) and os.path.exists(filepath):
                try:
                    self.__mcadapp.Open(filepath)
                    # The below method has to be used because ActiveWorksheet
                    # only returns an IMathcadPrimeWorksheet object. This does
                    # Not have all of the required methods
                    for i in range(self.__mcadapp.Worksheets.Count):
                        if self.__mcadapp.Worksheets.Item(i).Name == self.__mcadapp.ActiveWorksheet.Name:
                            self.__obj = self.__mcadapp.Worksheets.Item(i)  # Returns IMathcadPrimeWorksheet2 object
                            break
                except:
                    print(f"Error opening {filepath}")

    # ~~~~~~~~~~~~~~~~~~~~~~~ File Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def activate(self):
        """ activates the worksheet object """
        self.__obj.Activate()

    def Close(self, save_option="Save"):
        """ Closes the worksheet """
        if save_option in ["Discard", 2]:
            self.__obj.Close(2)
        elif save_option in ["Prompt", 1]:
            self.__obj.Close(1)
        elif save_option in ["Save", 0]:
            self.__obj.Close(0)
        else:
            print("incorrect save argument specified")

    def save_as(self, new_filepath):
        """ Saves the worksheet under a new filename """
        try:
            path = str(new_filepath)
            if os.path.isdir() and os.path.exists():
                self.__obj.SaveAs(path)
                return True
            else:
                raise ValueError("the argument for new_filepath is invalid")
        except TypeError:
            raise TypeError("new_filepath must be a string")
        except:
            print("COM error saving new version")

    def name(self):
        """ Returns the filename of the Worksheet object """
        return self.__obj.Name

    def readonly(self):
        """ Returns the worksheets read only status """
        return self.__obj.IsReadOnly  # Always return state

    def modified(self, setmodfied=None):
        """ Returns (and can optionally set) the worksheets modified status """
        if setmodfied is True:  # If readonly has been set to True
            self.__obj.Modified = True
        elif setmodfied is False:  # If readonly has been set to False
            self.__obj.Modified = False
        return self.__obj.Modified  # Always return state

    # ~~~~~~~~~~~~~~~~~~~~~ Worksheet Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def pause_calculation(self):
        """ Pauses worksheet calculation """
        self.__obj.PauseCalculation()

    def resume_calculation(self):
        """ Resumes the worksheets calculation """
        self.__obj.ResumeCalculation()

    def inputs(self):
        """ returns a list of the designated input fields in the worksheet """
        _inputs = []
        for i in range(self.__obj.Inputs.Count):  # no. of open sheets
            _inputs.append(self.__obj.Inputs.GetAliasByIndex(i))
        return _inputs  # Returns a list of open worksheet filenames

    def get_input(self, input_alias):
        """ Fetches the curent value of a specific input """
        if input_alias in self.inputs():
            getinput = self.__obj.InputGetRealValue(input_alias)
            return getinput.RealResult, getinput.Units, getinput.ErrorCode
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")

    def outputs(self):
        """ returns a list of the designated output fields in the worksheet """
        _outputs = []
        for i in range(self.__obj.Outputs.Count):
            _outputs.append(self.__obj.Outputs.GetAliasByIndex(i))
        return _outputs  # Returns a list of open worksheet filenames

    def get_real_output(self, output_alias, units="Default"):
        try:
            output_alias = str(output_alias)
            units = str(units)
            if output_alias in self.outputs():
                try:
                    if units == "Default":
                        self.__obj.OutputGetRealValue(output_alias)
                    else:
                        self.__obj.OutputGetRealValueAs(output_alias, units)
                except:
                    print("COM Error fetching real_output")
            else:
                raise ValueError(f"{output_alias} is not a designated output field" +
                                 f"\n\nAvailable Output fields:\n{self.outputs()}")
        except TypeError:
            raise TypeError("Arguments must be strings")

    def get_matrix_output(self, output_alias, units="Default"):
        try:
            output_alias = str(output_alias)
            units = str(units)
            if output_alias in self.outputs():
                try:
                    if units == "Default":
                        result = self.__obj.OutputGetMatrixValue(output_alias)
                    else:
                        result = self.__obj.OutputGetMatrixValueAs(output_alias, units)
                    return result.MatrixResult, result.Units, result.ErrorCode
                except:
                    print("COM Error fetching real_output")
            else:
                raise ValueError(f"{output_alias} is not a designated output field" +
                                 f"\n\nAvailable Output fields:\n{self.outputs()}")
        except TypeError:
            raise TypeError("Arguments must be strings")


    def set_real_input(self, input_alias, value, units=""):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.__obj.SetRealValue(str(input_alias), value, str(units))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n" +
                  f"Check the '{self.__mcadapp.ActiveWorksheet.Name}' worksheet\n")
        return error

    def set_string_input(self, input_alias, string):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.__obj.SetStringValue(str(input_alias), str(string))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n" +
                  f"Check the '{self.__mcadapp.ActiveWorksheet.Name}' worksheet\n")
        return error

    def set_matrix_input(self, input_alias, matrix_object, units=""):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.__obj.SetRealValue(str(input_alias),
                                            matrix_object, str(units))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n" +
                  f"Check the '{self.__mcadapp.ActiveWorksheet.Name}' worksheet\n")
        return error

    def syncronize(self):
        self.__obj.Synchronize()


class Matrix(object):
    """ Mathcad Matrix object container """
    def __init__(self, python_name=""):
        self.__mcadapp = w32c.Dispatch("MathcadPrime.Application")
        for i in range(self.__mcadapp.Worksheets.Count):
            if self.__mcadapp.Worksheets.Item(i).Name == self.__mcadapp.ActiveWorksheet.Name:
                self.__ws = self.__mcadapp.Worksheets.Item(i)  # Returns IMathcadPrimeWorksheet2 object
                break
        self.python_name = python_name  # Just for organisation in scripts
        self.object = None
        self.shape = None

    def create_matrix(self, rows, columns):
        """ Creates a Mathcad matrix """
        try:
            rows, columns = int(rows), int(columns)
            self.shape = (rows, columns)
            self.object = self.__ws.CreateMatrix(rows, columns)
            return True
        except ValueError:
            raise ValueError("Matrix dimensions must be integers")
        except:
            raise Exception("COM Error creating Mathcad matrix")

    def set_element(self, row_index, column_index, value):
        """ Sets the value of an element in the Matrix """
        if self.object is not None:
            try:
                row, col = int(row_index), int(column_index)
                print(self.object)
                self.object.SetMatrixElement(row, col, value)  # @FIXME
            except ValueError:
                raise ValueError("Matrix maths can only use numerical values")
#            except:
#                raise Exception("COM Error setting element value")  # Hidden for above @FIXME
        else:
            raise TypeError("Matrix must first be created")

    def get_element(self, row_index, column_index):
        """ Fetches the value of an element in the Matrix """
        if self.object is not None:
            try:
                row, col = int(row_index), int(column_index)
                self.object.GetMatrixElement(row, col)
            except ValueError:
                raise ValueError("Matrix maths can only use numerical values")
            except:
                raise Exception("COM Error fetching element value")
        else:
            raise TypeError("Matrix must first be created")

    def numpy_array_as_matrix(self, numpy_array):
        """ Takes a numpy array, creates a matrix, and populates the values """
        if isinstance(numpy_array, np.ndarray):
            height, width = numpy_array.shape  # Get array dimensions
            matrix = self.create_matrix(width, height)
            if matrix is True:
                for r, row in enumerate(numpy_array):
                    for c, value in enumerate(row):
                        self.set_element(r, c, value)
        else:
            raise TypeError("Argument is not a Numpy array")


if __name__ == "__main__":
    TEST = os.path.join(os.getcwd(), "Test", "test.mcdx")
    MC = Mathcad(visible=True) # Open Mathcad with no GUI
    WS = Worksheet(TEST)
    WS = Worksheet(None, "test")
    a = WS.set_real_input("in_test", 9, "mm")
    print(a)
    matrix, units, error = WS.get_matrix_output("out_999")
    print(error)
    #matrix.
