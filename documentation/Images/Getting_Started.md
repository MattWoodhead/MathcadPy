# MathcadPy - Getting Started
How to get started with MathcadPy

### Requirements
- [Mathcad Prime 3+](https://www.mathcad.com/) Ensure that you have got PTC Mathcad Prime installed. This wrapper has been tested with Mathcad Prime versions 3, 4 and 5.
- [win32com](https://github.com/mhammond/pywin32) MathcadPy uses the pywin32 package to interact with the Mathcad COM api. This should automatically be installed when you use pip to install MathcadPy, but you can also install it before hand or use an existing installation.


### Installation
MathcadPy has been released on PyPI, so installation is simple as follows:

    pip install MathcadPy

Thats it!

### Testing the link to Mathcad
Start a fresh python script, and type the following:

    from MathcadPy import Mathcad
    
    mathcad_app = Mathcad()  # creates an instance of the Mathcad class - this object represents the Mathcad window
    
    print(f"Mathcad version: {mathcad_app.version}")  # Check the mathcad version and print to the console
    
Save and then run the script. You should see an output similar to the following in the terminal window:

![alt text](https://github.com/MattWoodhead/MathcadPy/blob/master/documentation/Example_000.PNG "Fetching the Mathcad version - terminal output")

Note that there may be a slight delay whilst Mathcad loads if you do not already have it open.


### Designating an input or an output in the Mathcad worksheet
Now we have verified that the link between your Python script and Mathcad is working, we are ready to start interacting with Mathcad. First, you must designate inputs and outputs in the Mathcad worksheet you want to interact with. Select the value you want to designate as an input:

![alt text](https://github.com/MattWoodhead/MathcadPy/blob/master/documentation/Example_001.PNG "Selecting a value to designate as an input")


Go to the ribbon interface, and navigate to the input/output tab. Then select "Assign Inputs"

![alt text](https://github.com/MattWoodhead/MathcadPy/blob/master/documentation/Example_002.PNG "Designating an input")

Assigning an output is completed in the same manner, but using the command "Assign Ouputs". You can view all designated inputs and outputs in the worksheet using the "Show As List" button in the ribbon. The list is formatted with the worksheet name in the left column, and the alias you will use in your python script on the right.

Note that input alias values deafualt to the worksheet name, and output alias values default to "out", "out1", "out2" etc. It is worth changing these to more meaningful names to make your python script simpler to understand.


### Sending and Fetching worksheet values from Python
TODO
