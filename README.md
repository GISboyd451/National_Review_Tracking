# National_Review_Tracking
Used in the monthly national QAQC process.

## Overview


## Requirements
* Python 3.4+
* Works on Linux, Windows, and Mac OSX.

#### Packages Needed
- os (default)
- time (default)
- glob (default)
- sys (default)
- datetime (https://docs.python.org/3/library/datetime.html)
- pandas (http://pandas.pydata.org/)
- numpy (https://pypi.org/project/numpy/)
- calendar (https://docs.python.org/3/library/calendar.html)
- shutil (https://docs.python.org/3/library/shutil.html)
- openpyxl (https://pypi.org/project/openpyxl/)

## Documentation
This script currently operates off of several report tables (excel, csv, etc.) to retrieve the monthly QAQC report numbers and stats to be formatted/appened into the 'National_Review_Tracking.xlsx'. The 'National_Review_Tracking.xlsx' will contain the quality numbers and the percent change over time, going back to the year 2017. *Note: Changes will be monthly.

EXCEL CURRENT SETUP

Info1 | Info2 | Info3 | Info4 | 
------------ | -------------|-------------|-------------|
Text Field | Text Field | Text Field | Text Field | 

- Info1 = State ADMIN CODE 
- Info2 = Dataset 
- Info3 = Attribute Name
- Info4 = Feature Class 

*Monthly change values will be to the right of the above constants.These values include:

- Pass Counts = How many records are passing quality control
- Total Counts = Total number of records
- Accuracy % = Percent of records passing
- Percent Change = Percent Accuracy change from previous month

## Running The Script



## Release Notes
Version: v1 02/27/2020
