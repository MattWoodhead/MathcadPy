# -*- coding: utf-8 -*-
"""
_application.py
~~~~~~~~~~~~~~
MathcadPy
https://github.com/MattWoodhead/MathcadPy
Copyright 2023 Matt Woodhead
"""

from pathlib import Path
import pythoncom
import win32com.client as w32c


class Mathcad:
    """Mathcad application object"""

    _version_int = 0  # class variable for the Mathcad version

    def __init__(self, visible=True):
        # print("Loading Mathcad")
        try:
            self.__mcadapp = w32c.Dispatch("MathcadPrime.Application")

            self.version = self.__mcadapp.GetVersion()  # Fetches Mathcad version
            self.open_worksheets = {}
            if visible is False:
                self.__mcadapp.Visible = False
            else:
                self.__mcadapp.Visible = True
            self._list_worksheets()
        except pythoncom.com_error as pcoe:
            try:
                if pcoe.args[1] == "Invalid class string":
                    raise MathcadComError("Could not locate the Mathcad Automation API") from pcoe
            except:
                raise pythoncom.com_error from pcoe

    def __getattribute__(*args):
        """ Used to allow access to hidden attributes of class instances """
        # https://docs.python.org/3/reference/datamodel.html#special-method-lookup
        return object.__getattribute__(*args)

    def _list_worksheets(self):
        """lists worksheets open in the Mathcad instance"""
        ws_list = {}
        for i in range(self.__mcadapp.Worksheets.Count):
            ws_list[self.__mcadapp.Worksheets.Item(i).Name] = Worksheet(
                self.__mcadapp.Worksheets.Item(i)
            )  # {name: ws_object}
        self.open_worksheets = ws_list

    def activate(self):
        """Activate the Mathcad window. If visible, this maximises Mathcad"""
        self.__mcadapp.Activate()

    def get_version(self):
        """Fetches the version string from the attached MathCAD instance"""
        self.version = self.__mcadapp.GetVersion()  # update the class variables
        Mathcad._version_int = int(self.version[0])
        return self.version  # return the version string to the function caller

    def active_sheet(self):
        """Returns the active worksheet name"""
        # TODO - should this be changed to return a sheet object?
        name = self.__mcadapp.ActiveWorksheet.Name
        if name == "":
            return None  # Returns none if the current worksheet not saved
        return name

    def worksheet_names(self):
        """lists worksheets open in the Mathcad instance"""
        worksheets = []
        for i in range(self.__mcadapp.Worksheets.Count):  # no. of open sheets
            worksheets.append(self.__mcadapp.Worksheets.Item(i).Name)
        return worksheets  # Returns a list of open worksheet filenames

    def worksheet_paths(self):
        """lists worksheets open in the Mathcad instance"""
        worksheets = []
        for i in range(self.__mcadapp.Worksheets.Count):  # no. of open sheets
            worksheets.append(self.__mcadapp.Worksheets.Item(i).FullName)
        return worksheets  # Returns a list of open worksheet filenames

    def open(self, filepath: Path):
        """Opens the filepath (if valid) in Mathcad"""
        try:
            filepath = Path(filepath)
            if not filepath.exists():
                raise FileNotFoundError()
            if filepath.suffix.lower() != ".mcdx":
                raise ValueError()

            local_obj = self.__mcadapp.Open(str(filepath))
            # now we have opened a new worksheet, generate the list of open sheets from the API
            local_worksheets = {}
            # need to use a for-loop because the Mathcad API is shit
            for i in range(self.__mcadapp.Worksheets.Count):
                sheet_object = self.__mcadapp.Worksheets.item(i)
                # the api open method only returns a basic IMathcadPrimeWorksheet object
                local_worksheets[sheet_object.Name] = sheet_object

            # add the worksheet into the open worksheets dictionary
            self.open_worksheets[local_obj.Name] = Worksheet(local_worksheets[local_obj.Name])
            return self.open_worksheets[local_obj.Name]  # return the worksheet object

        except TypeError as exc:
            raise TypeError(
                f"filepath expects a string or pathlib.Path object. Got {type(filepath)}"
            ) from exc
        except ValueError as exc:
            raise ValueError(f"The provided path is not a Mathcad Prime file: {filepath}") from exc
        except FileNotFoundError as exc:
            raise FileNotFoundError(f"The provided path does not exist: {filepath}") from exc

    def close_all(self, save_option="Discard"):
        """Closes all worksheets. Can specify save options before closing"""
        if save_option in ["Discard", 2]:  # check for both "Discard" and its COM equivalent enum
            self.__mcadapp.CloseAll(2)
        elif save_option in ["Prompt", 1]:
            self.__mcadapp.CloseAll(1)
        elif save_option in ["Save", 0]:
            self.__mcadapp.CloseAll(0)
        else:
            print("incorrect save argument specified")
        self._list_worksheets()

    def quit(self, save_option="Discard"):
        """
        Closes all worksheets and closes the MathCAD instance.
        Can specify save options before closing
        """
        if save_option in ["Discard", 2]:  # check for both "Discard" and its COM equivalent enum
            self.__mcadapp.Quit(2)
        elif save_option in ["Prompt", 1]:
            self.__mcadapp.Quit(1)
        elif save_option in ["Save", 0]:
            self.__mcadapp.Quit(0)


class Worksheet:
    """Mathcad Worksheet object

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
        """activates the worksheet object"""
        self.ws_object.Activate()

    def close(self, save_option="Save"):
        """Closes the worksheet"""
        if save_option in ["Discard", 2]:
            self.ws_object.Close(2)
        elif save_option in ["Prompt", 1]:
            self.ws_object.Close(1)
        elif save_option in ["Save", 0]:
            self.ws_object.Close(0)
        else:
            print("incorrect save argument specified")

    def save(self):
        """Saves the worksheet"""
        self.ws_object.Save()

    def save_as(self, new_filepath: Path):
        """Saves the worksheet under a new filename"""
        new_filepath = Path(new_filepath)  # Cast to Path object incase they have used a string
        if new_filepath.suffix.lower() == ".pdf":
            if Mathcad._version_int > 7:
                # if _get_mathcad_version() > 7:
                self.ws_object.SaveAs(new_filepath)
            else:
                raise ValueError("Mathcad Prime 8 or newer is required to export as PDF")
        elif new_filepath.suffix.lower() == ".mcdx":
            self.ws_object.SaveAs(new_filepath)
        else:
            raise ValueError("Filename must include file extension '.mcdx' or '.pdf'")

    def name(self):
        """Returns the filename of the Worksheet object"""
        return self.ws_object.Name

    def is_readonly(self):
        """Returns the worksheets read only status"""
        return self.ws_object.IsReadOnly  # Always return state

    def is_modified(self, setmodfied=None):
        """Returns (and can optionally set) the worksheets modified status"""
        if setmodfied is True:  # If readonly has been set to True
            self.ws_object.Modified = True
        return self.ws_object.Modified  # Always return state

    # ~~~~~~~~~~~~~~~~~~~~~ Worksheet Operations ~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def pause_calculation(self):
        """Pauses worksheet calculation"""
        self.ws_object.PauseCalculation()

    def resume_calculation(self):
        """Resumes the worksheets calculation"""
        self.ws_object.ResumeCalculation()

    def inputs(self):
        """returns a list of the designated input fields in the worksheet"""
        _inputs = []
        for i in range(self.ws_object.Inputs.Count):  # no. of open sheets
            _inputs.append(self.ws_object.Inputs.GetAliasByIndex(i))
        return _inputs

    def get_input(self, input_alias):
        """Fetches the curent value of a specific input"""
        if input_alias in self.inputs():
            try:
                result = self.ws_object.InputGetValue(input_alias)
                result_type = result.ResultType
                if result_type == 1:  # ValueResultTypes_Real
                    return result.RealResult, result.Units, result.ErrorCode
                if result_type == 2:  # ValueResultTypes_String
                    return result.StringResult, result.Units, result.ErrorCode
                if result_type == 3:  # ValueResultTypes_Matrix
                    return _matrix_to_array(result.MatrixResult), result.Units, result.ErrorCode
                # else
                return None, None, None
            except pythoncom.com_error as pcoe:
                raise MathcadComError("COM Error fetching real_output") from pcoe
            getinput = self.ws_object.InputGetRealValue(input_alias)
            return getinput.RealResult, getinput.Units, getinput.ErrorCode
        # else
        raise ValueError(f"{input_alias} is not a designated input field")

    def _get_real_input_units(self, input_alias):
        """
        Fetches the units of a specific matrix input. This is an internal function and no input
        sanitisation checks are performed
        """
        return self.ws_object.InputGetRealValue(input_alias).Units

    def get_matrix_input(self, input_alias):
        """Fetches the curent value of a specific input"""
        if input_alias in self.inputs():
            getinput = self.ws_object.InputGetMatrixValue(input_alias)
            return _matrix_to_array(getinput.MatrixResult), getinput.Units, getinput.ErrorCode
        # else
        raise ValueError(f"{input_alias} is not a designated input field")

    def _get_matrix_input_units(self, input_alias):
        """
        Fetches the units of a specific matrix input. This is an internal function and no input
        sanitisation checks are performed
        """
        return self.ws_object.InputGetMatrixValue(input_alias).Units

    def outputs(self):
        """returns a list of the designated output fields in the worksheet"""
        _outputs = []
        for i in range(self.ws_object.Outputs.Count):
            _outputs.append(self.ws_object.Outputs.GetAliasByIndex(i))
        return _outputs  # Returns a list of output aliases

    def get_real_output(self, output_alias, units="Default"):
        """Gets the numerical value from a designated output in the worksheet"""
        assert isinstance(output_alias, str)
        assert isinstance(units, str)
        if output_alias in self.outputs():
            try:
                if units == "Default":
                    result = self.ws_object.OutputGetRealValue(output_alias)
                    return result.RealResult, result.Units, result.ErrorCode
                # else
                result = self.ws_object.OutputGetRealValueAs(output_alias, units)
                return result.RealResult, units, result.ErrorCode
            except pythoncom.com_error as pcoe:
                raise MathcadComError("COM Error fetching real_output") from pcoe
        else:
            raise ValueError(f"{output_alias} is not a designated output field")

    def get_output(self, output_alias):
        """Gets the value from a designated output in the worksheet"""
        assert isinstance(output_alias, str)
        if output_alias in self.outputs():
            try:
                result = self.ws_object.OutputGetValue(output_alias)
                result_type = result.ResultType
                if result_type == 1:  # ValueResultTypes_Real
                    return result.RealResult, result.Units, result.ErrorCode
                if result_type == 2:  # ValueResultTypes_String
                    return result.StringResult, result.Units, result.ErrorCode
                if result_type == 3:  # ValueResultTypes_Matrix
                    return _matrix_to_array(result.MatrixResult), result.Units, result.ErrorCode
                # else
                return None, None, None
            except pythoncom.com_error as pcoe:
                raise MathcadComError("COM Error fetching real_output") from pcoe
        else:
            raise ValueError(f"'{output_alias}' is not a designated output field")

    def _get_output(self, output_alias):
        """DEPRECATED: Gets the value from a designated output in the worksheet"""
        # TODO - add deprecation notice
        return self.get_output(output_alias)

    def get_matrix_output(self, output_alias, units="Default"):
        """Gets the numerical value from a designated output in the worksheet"""
        assert isinstance(output_alias, str)
        assert isinstance(units, str)
        if output_alias in self.outputs():
            try:
                if units == "Default":
                    result = self.ws_object.OutputGetMatrixValue(output_alias)
                else:
                    result = self.ws_object.OutputGetMatrixValueAs(output_alias, units)
                    print(dir(result))
                return _matrix_to_array(result.MatrixResult), None, result.ErrorCode
            except pythoncom.com_error as pcoe:
                raise MathcadComError("COM Error fetching matrix output") from pcoe
        else:
            raise ValueError(f"{output_alias} is not a designated output field")

    def set_real_input(self, input_alias, value, units="", preserve_worksheet_units=True):
        """Set the value of a numerical input range in the worksheet"""
        assert isinstance(input_alias, str)
        assert isinstance(units, str)
        assert isinstance(preserve_worksheet_units, bool)
        if input_alias in self.inputs():  # Use inputs function to get list
            if preserve_worksheet_units:
                previous_units = self._get_real_input_units(input_alias)
                if units:  # If units is not equal to ""
                    try:
                        assert units == previous_units
                    except AssertionError as exc:
                        raise AssertionError(
                            "preserve_worksheet_units is True. The units argument "
                            "does not equate to the units present in the Worksheet"
                        ) from exc
                else:  # No units are specified, but preserve_worksheet_units is True
                    units = previous_units
            error = self.ws_object.SetRealValue(input_alias, value, units)
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def set_string_input(self, input_alias, string_value):
        """Set the value of a numerical input range in the worksheet"""
        assert isinstance(input_alias, str)
        assert isinstance(string_value, str)
        if input_alias in self.inputs():  # Use inputs function to get list
            error = self.ws_object.SetStringValue(input_alias, string_value)
            # COM command returns error count. 0 = everything set correctly
        else:
            raise ValueError(f"{input_alias} is not a designated input field")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def set_matrix_input(self, input_alias, matrix_array, units="", preserve_worksheet_units=True):
        """Set the value of a numerical input range in the worksheet"""
        assert isinstance(input_alias, str)
        assert isinstance(units, str)
        assert isinstance(preserve_worksheet_units, bool)
        if input_alias in self.inputs():  # Check that the alias specified exists in the worksheet
            if preserve_worksheet_units:
                previous_units = self._get_matrix_input_units(input_alias)
                if units:  # If units is not equal to ""
                    try:
                        assert units == previous_units
                    except AssertionError as exc:
                        raise AssertionError(
                            "preserve_worksheet_units is True. The units argument "
                            "does not equate to the units present in the Worksheet"
                        ) from exc
                else:  # No units are specified, but preserve_worksheet_units is True
                    units = previous_units

            rows, cols = _array_check(matrix_array)
            temp_matrix = self.ws_object.CreateMatrix(rows, cols)
            for row in range(rows):
                for col in range(cols):
                    value = matrix_array[row][col]
                    try:
                        temp_matrix.SetMatrixElement(row, col, value)
                    except Exception as exc:
                        raise ValueError(
                            f"Error setting matrix element {row},{col}: {value}"
                        ) from exc

            error = self.ws_object.SetMatrixValue(str(input_alias), temp_matrix, str(units))
            # error = self.ws_object.SetRealValue(str(input_alias),
            #                                     matrix_array, str(units))
            # COM command returns error count. error = 0 -> everything set correctly

        else:
            raise ValueError(f"{input_alias} is not a designated input field")
        if error > 0:
            print(f"\nWarning!\nerror setting '{input_alias}' value/units\n")
        return error

    def PauseCalculation(self):  # todo - duplicate of pause_calculation
        """DEPRECATED: Pauses worksheet calculation - may speed up routines the set many input values"""
        print(
            "Warning: the PauseCalculation method will be removed in a future version "
            "- use pause_calculation instead"
        )
        self.ws_object.PauseCalculation()

    def ResumeCalculation(self):  # todo - duplicate of resume_calculation
        """DEPRECATED: Pauses worksheet calculation"""
        print(
            "Warning: the ResumeCalculation method will be removed in a future version "
            "- use resume_calculation instead"
        )
        self.ws_object.ResumeCalculation()

    def syncronize(self):
        """Syncronises (i.e. re-calculates) worksheet"""
        self.ws_object.Synchronize()

    def calculate(self):
        """Syncronises (i.e. re-calculates) worksheet"""
        self.ws_object.Synchronize()


def _matrix_to_array(mathcad_matrix_obj) -> list:
    """converts a COM matrix object to a list of lists (row = sub list, column = value)"""

    rows = int(mathcad_matrix_obj.Rows)
    # print(f"rows: {rows}")
    cols = int(mathcad_matrix_obj.Columns)
    # print(f"cols: {cols}")
    matrix = []
    for row in range(rows):
        row_list = [mathcad_matrix_obj.GetMatrixElement(row, col) for col in range(cols)]
        matrix.append(row_list)
    return matrix


def _array_check(matrix_array: list):
    """A helper function to validate that the array input is suitable to be sent to Mathcad"""
    rows = len(matrix_array)
    previous_cols = 0
    for i, row_list in enumerate(matrix_array):
        cols = len(row_list)
        if i != 0:
            if cols != previous_cols:  # check that every row has the same number of columns
                raise ValueError("Inconsistent number of columns in input matrix")
        previous_cols = cols
    return rows, cols


class MathcadComError(Exception):
    """Base class for all COM related exceptions for the MathcadPy module.

    If specified, the error code is automatically appended to the message:

    >>> # With an error code (it also works with a specific error):
    >>> error = CanOperationError(message="Failed to do the thing", error_code=42)
    >>> str(error)
    'Failed to do the thing [Error Code 42]'
    >>>
    >>> # Missing the error code:
    >>> plain_error = CanError(message="Something went wrong ...")
    >>> str(plain_error)
    'Something went wrong ...'

    :param error_code:
        An optional error code to narrow down the cause of the fault

    :arg error_code:
        An optional error code to narrow down the cause of the fault
    """

    def __init__(self, message: str = "", error_code: int = None) -> None:
        self.error_code = error_code
        super().__init__(message if error_code is None else f"{message} [Error Code {error_code}]")


if __name__ == "__main__":
    mc = Mathcad()
    print(mc.get_version())
    print(Mathcad._version_int)
