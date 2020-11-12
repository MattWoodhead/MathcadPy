# -*- coding: utf-8 -*-
"""
MathcadPy __application_win32com.py

Author: MattWoodhead

"""

from pathlib import Path
import win32com.client as w32c
import numpy as np


class Mathcad():
    """ Mathcad application object """

    def __init__(self, visible=True):
        print("Loading Mathcad")
        try:
            self.__mcadapp = w32c.Dispatch("MathcadPrime.Application")
            self.version = self.__mcadapp.GetVersion()  # Fetches Mathcad version
            self.open_worksheets = {}
            if visible is False:
                self.__mcadapp.Visible = False
            else:
                self.__mcadapp.Visible = True
            self._list_worksheets()
        except:  # TODO - improve error handling - specific COM exceptions
            print("Could not locate the Mathcad Automation API")

    def _list_worksheets(self):
        """ lists worksheets open in the Mathcad instance """
        ws_list = {}
        for i in range(self.__mcadapp.Worksheets.Count):
            ws_list[self.__mcadapp.Worksheets.Item(i).Name] = Worksheet(self.__mcadapp.Worksheets.Item(i))  # {name: ws_object}
        self.open_worksheets = ws_list

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

    def open(self, filepath: Path):
        """ Opens the filepath (if valid) in Mathcad """
        try:
            if not isinstance(filepath, Path):
                filepath = Path(filepath)
            if filepath.exists() and (filepath.suffix.lower() == ".mcdx"):
                local_obj = self.__mcadapp.Open(str(filepath))
                # now we have opened a new worksheet, generate the list of open worksheets from the COM API
                local_worksheets = {}
                for i in range(self.__mcadapp.Worksheets.Count):  # a for loop because the Mathcad API is shit
                    sheet_object = self.__mcadapp.Worksheets.item(i)
                    local_worksheets[sheet_object.Name] = sheet_object# this is necessary because the open method only returns a basic IMathcadPrimeWorksheet object
                self.open_worksheets[local_obj.Name] = Worksheet(local_worksheets[local_obj.Name])  # add the worksheet into the open worksheets dictionary
                return self.open_worksheets[local_obj.Name]  # return the worksheet object
            else:
                raise ValueError("The provided path is not a Mathcad Prime file")
        except TypeError:
            raise TypeError("filepath expects a string or pathlib object")

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
        self._list_worksheets()


class Worksheet():
    """ Mathcad Worksheet object

    Either a filepath for a mathcad file can be supplied, or the
    filepath can be set to None (or similar) and the optional
    open_sheet_name argument can be used
    """

    def __init__(self, _worksheet_COM_object=None):
        self.ws_object = _worksheet_COM_object
        # try:
        self.__repr__ = self.ws_object.FullName
        # except:
        #     self.__repr__ = "__repr__error"

    def activate(self):
        """ activates the worksheet object """
        self.ws_object.Activate()

    def Close(self, save_option="Save"):
        """ Closes the worksheet """
        if save_option in ["Discard", 2]:
            self.ws_object.Close(2)
        elif save_option in ["Prompt", 1]:
            self.ws_object.Close(1)
        elif save_option in ["Save", 0]:
            self.ws_object.Close(0)
        else:
            print("incorrect save argument specified")

    def save_as(self, new_filepath: Path):
        """ Saves the worksheet under a new filename """
        try:
            new_filepath = Path(new_filepath)
            if new_filepath.is_file():
                self.ws_object.SaveAs(new_filepath)
                return True
            else:
                raise ValueError("new_filepath must be a file name, not a directory")
        except TypeError:
            raise TypeError("new_filepath must be a String or Pathlib object")
        except:
            print("COM error saving new version")

    def name(self):
        """ Returns the filename of the Worksheet object """
        return self.ws_object.Name

    def is_readonly(self):
        """ Returns the worksheets read only status """
        return self.ws_object.IsReadOnly  # Always return state

    def modified(self, setmodfied=None):
        """ Returns (and can optionally set) the worksheets modified status """
        if setmodfied is True:  # If readonly has been set to True
            self.ws_object.Modified = True
        elif setmodfied is False:  # If readonly has been set to False
            self.ws_object.Modified = False
        return self.ws_object.Modified  # Always return state

    # ~~~~~~~~~~~~~~~~~~~~~ Worksheet Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def pause_calculation(self):
        """ Pauses worksheet calculation """
        self.ws_object.PauseCalculation()

    def resume_calculation(self):
        """ Resumes the worksheets calculation """
        self.ws_object.ResumeCalculation()

    def inputs(self):
        """ returns a list of the designated input fields in the worksheet """
        _inputs = []
        for i in range(self.ws_object.Inputs.Count):  # no. of open sheets
            _inputs.append(self.ws_object.Inputs.GetAliasByIndex(i))
        return _inputs  # Returns a list of open worksheet filenames

    def get_input(self, input_alias):
        """ Fetches the curent value of a specific input """
        if input_alias in self.inputs():
            getinput = self.ws_object.InputGetRealValue(input_alias)
            return getinput.RealResult, getinput.Units, getinput.ErrorCode
        else:
            raise ValueError(f"{input_alias} is not a designated input field")

    def outputs(self):
        """ returns a list of the designated output fields in the worksheet """
        _outputs = []
        for i in range(self.ws_object.Outputs.Count):
            _outputs.append(self.ws_object.Outputs.GetAliasByIndex(i))
        return _outputs  # Returns a list of open worksheet filenames

    def get_real_output(self, output_alias, units="Default"):
        """  """
        try:
            output_alias = str(output_alias)
            units = str(units)
            if output_alias in self.outputs():
                try:
                    if units == "Default":
                        result = self.ws_object.OutputGetRealValue(output_alias)
                    else:
                        result = self.ws_object.OutputGetRealValueAs(output_alias, units)
                    return result.RealResult, result.Units, result.ErrorCode
                except:
                    print("COM Error fetching real_output")
            else:
                raise ValueError(f"{output_alias} is not a designated output field")
        except TypeError:
            raise TypeError("Arguments must be strings")

    def get_matrix_output(self, output_alias, units="Default"):
        try:
            output_alias = str(output_alias)
            units = str(units)
            if output_alias in self.outputs():
                try:
                    if units == "Default":
                        result = self.ws_object.OutputGetMatrixValue(output_alias)
                    else:
                        result = self.ws_object.OutputGetMatrixValueAs(output_alias, units)
                    return result.MatrixResult, result.Units, result.ErrorCode
                except:
                    print("COM Error fetching real_output")
            else:
                raise ValueError(f"{output_alias} is not a designated output field")
        except TypeError:
            raise TypeError("Arguments must be strings")


    def set_real_input(self, input_alias, value, units=""):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.ws_object.SetRealValue(str(input_alias), value, str(units))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def set_string_input(self, input_alias, string):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.ws_object.SetStringValue(str(input_alias), str(string))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def set_matrix_input(self, input_alias, matrix_object, units=""):
        """ Set the value of a numerical input range in the worksheet """
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.ws_object.SetRealValue(str(input_alias),
                                                matrix_object, str(units))
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field" +
                             f"\n\nAvailable Input fields:\n{self.inputs()}")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def syncronize(self):
        self.ws_object.Synchronize()


class Matrix():
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
