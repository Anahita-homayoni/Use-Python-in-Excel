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


Sorting data with sort_values() in Python in Excel
sortedbands = bands.sort_values(by=["Genre", "Year"], ascending=[True, False])

What this formula does

This line sorts your bands DataFrame in two levels:

By Genre (A → Z)

By Year (Newest → Oldest inside each Genre)

Explanation of each part
Part	Meaning
bands	Your original Excel data loaded as a DataFrame
sort_values()	Pandas function used to sort data
by=["Genre", "Year"]	Columns used for sorting
ascending=[True, False]	Sort order for each column
Ascending vs Descending
Value	Order	Example
True	Ascending (A→Z, 0→9, Oldest→Newest)	
False	Descending (Z→A, 9→0, Newest→Oldest)	

So:

Genre = True → A to Z

Year = False → Newest to Oldest (inside each Genre)

Result

Your data is grouped by Genre, and within each genre, bands are ordered by latest year first.

Simple explanation for beginners

This formula organizes your table by Genre alphabetically, and inside each Genre it shows the newest bands first.


Filtering data with query() in Python in Excel
bandsq = bands.query("Genre == 'Alternative rock' and Members >= 4")

What this formula does

This line filters your bands DataFrame and keeps only the rows that match both conditions:

Genre is Alternative rock

Members are 4 or more

The filtered result is saved in a new DataFrame called bandsq.

Explanation of each part
Part	Meaning
bands	Your original Excel data as a DataFrame
query()	Pandas function used to filter rows
"Genre == 'Alternative rock'"	Keeps only Alternative rock bands
Members >= 4	Keeps bands with 4 or more members
bandsq	New DataFrame containing the filtered result
Operators used
Symbol	Meaning
==	Equal to
>=	Greater than or equal to
and	Both conditions must be true
Simple explanation

This formula finds all Alternative rock bands that have at least four members and stores them in a new table.



Filtering rows using startswith()
bands[bands["Group"].str.startswith("The")]

What this formula does

This formula filters your bands table and returns only the rows where the Group name starts with the word “The”.

Example matches:

The Beatles

The Doors

The Rolling Stones

Explanation of each part
Part	Meaning
bands	Your original DataFrame (Excel table)
bands["Group"]	Selects the Group column
.str	Enables text (string) functions
.startswith("The")	Checks if the name begins with “The”
bands[ ... ]	Returns only rows that match the condition
Simple explanation

This formula finds all band names that start with “The” and shows them in a new filtered table.

Tip for learners

To ignore capital letters (The, the, THE), use:

bands[bands["Group"].str.lower().str.startswith("the")]




Example: Convert Group names to uppercase and sort
# Convert all Group names to uppercase
sortedbands["Group"] = sortedbands["Group"].apply(str.upper)

# Sort the DataFrame by Genre and Year
sortedbands = sortedbands.sort_values(by=["Genre", "Year"], ascending=[True, False])

# Show the result
sortedbands

What this does
Line	Explanation
sortedbands["Group"].apply(str.upper)	Converts every band name in the Group column to uppercase letters
sortedbands.sort_values(...)	Sorts the DataFrame first by Genre (A→Z), then by Year (Newest→Oldest)
sortedbands	Displays the updated table

Example Input
Group	Genre	Year	Members
The Beatles	Rock	1965	4
nirvana	Alternative Rock	1991	3
Metallica	Metal	1983	4
Example Output after applying str.upper and sorting
Group	Genre	Year	Members
METALLICA	Metal	1983	4
NIRVANA	Alternative Rock	1991	3
THE BEATLES	Rock	1965	4
Why this is useful for learners
.apply(str.upper) — Standardizes text formatting. Makes filtering (like .startswith("THE")) consistent.

.sort_values() — Organizes your table by multiple criteria (Genre, Year).

Combines text manipulation and sorting, which is very common in data cleaning.
