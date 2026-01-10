# Use-Python-in-Excel
This repository contains beginner-friendly tutorials, examples, and exercises for learning how to use Python directly in Microsoft Excel.

Function of XL() in Python in Excel

XL() is a special function used in Python in Excel to connect Python with Excel cells, tables, and ranges.

It lets Python read data from Excel and send results back into Excel.

What it does:
Task	How XL() helps
Get data from Excel	Reads values from selected cells or tables
Send results to Excel	Returns Python output into worksheet cells
Keep formulas dynamic	Recalculates automatically when Excel data changes
Work with tables easily	Handles Excel Tables as DataFrames


Simple example:
import pandas as pd

data = XL("A1:C10")      # Reads Excel range into Python
result = data.sum()     # Python calculation
result                  # Sends answer back to Excel

In short:

XL() is the bridge between Excel and Python.
It allows Excel users to run Python calculations directly inside Excel.


Getting information about a DataFrame

Once your Excel data is inside Python (as a DataFrame), you can inspect it using these useful properties:

Function	What it shows
x.shape	Number of rows and columns
x.size	Total number of values
x.columns	Column names
x.head()	First 5 rows
x.tail()	Last 5 rows
x.info()	Data types and missing values
x.describe()	Summary statistics
Example: Get DataFrame information in Excel
=PY(
x = XL("A1:C10")
x.shape
)


Output:
Shows how many rows and columns are in your Excel range.

=PY(
x = XL("A1:C10")
x.columns
)


Output:
Displays the column names.

In simple words:

XL() brings Excel data into Python.
pandas functions like .shape, .size, and .columns let you explore and understand your Excel data easily.
