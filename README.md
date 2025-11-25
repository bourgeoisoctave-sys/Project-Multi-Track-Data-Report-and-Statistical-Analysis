Multi-Track Student Performance Analysis

Python Data Cleaning, Merging, and Statistical Reporting

\---------------------------------------------------------

1. PROJECT OVERVIEW

\-------------------

This project analyzes student performance data from three academic tracks:

- Data
- Finance
- Business Management (BM)

The program:

- Imports data from three Excel sheets
- Cleans the data (removes "WAIVE", fixes missing values, converts to numeric)
- Merges all tracks into one unified dataset
- Computes descriptive statistics (subject averages, attendance, pass rate)
- Performs cross-track, cohort, and income-group comparisons
- Generates visualizations (boxplots and bar charts)
- Exports CSV summary files and cleaned data

The goal is to practice data cleaning, aggregation, statistical analysis, and visualization in Python.


1. FILES INCLUDED

\-----------------

Python Script:

- main.py  (main program performing import, cleaning, analysis, and export)

Input File:

- student\_grades\_2027-2028.xlsx

Sheets: Data, Finance, BM

Generated Output Files:

- cleaned\_file.csv
- summary\_statistics.csv
- boxplot\_history.png
- avg\_math\_per\_track.png

Statistical tables and comparisons are also printed to the terminal.


1. REQUIREMENTS

\---------------

Python 3.8 or newer

Required packages:

- pandas
- numpy
- matplotlib
- openpyxl

Install dependencies with:

pip install pandas numpy matplotlib openpyxl


1. HOW TO RUN THE PROGRAM

\-------------------------

1. Place main.py and student\_grades\_2027-2028.xlsx in the same folder.
1. Open a terminal in that folder:

cd /path/to/project

1. Run the script:

python main.py

Outputs:

- cleaned\_file.csv : cleaned and merged dataset
- summary\_statistics.csv : summary metrics
- PNG image files : plots
- Full analysis printed in console

No command-line arguments are required.


5\. FEATURES IMPLEMENTED

\-----------------------

- Data import from three Excel sheets
- Track identification via "Track" column
- Cleaning:
* Replace "WAIVE" with NaN
* Convert grades and attendance to numeric
* Normalize "Passed (Y/N)" values
* Strip spaces from column names
- Merging into a single DataFrame
- Statistical analysis:
* Average Math, English, Science, History, Attendance
* Student counts per track
* Overall pass rate
* Correlation between Attendance and ProjectScore (if available)
- Cross-track analysis:
* Average Math comparison
* History score boxplot
* Math score bar chart
- Cohort-level analysis (subject averages, pass rate)
- Income-level analysis (pass rate by income group)
- Export of CSV summaries and plot images


6\. ASSUMPTIONS

\--------------

- All three Excel sheets share the same structure.
- Subject columns exist: Math, English, Science, History.
- Attendance column is named "Attendance (%)".
- Cohort and IncomeStudent columns are present.
- Non-numeric values like "WAIVE" indicate missing data.


7\. AUTHORS

\---------

Octave Bourgeois Axelle Sebban Loris Raoul-Parakian

EDHEC Business School

Master in Management — Data Science & AI for Business
