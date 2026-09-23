# bond_automation_script

Python script that automates part of the daily fixed income trading report process.

## What This Does
- Reads raw daily bond market data from Excel spreadsheet
- Cleans and standardizes the data (handles missing values, converts text fields to numeric)
- Categorizes bonds by type (Treasury, policy financial, local government) and groups them by maturity range
- Organizes institutional categories (banks, securities firms, insurers, fund companies) into a consistent order
- Outputs a formatted Excel report with colour-coded conditional formatting for faster review

## Why I Built It
During a fixed income internship, compiling this report manually was repetitive and time-consuming. My mentor suggested writing a Python script to ameliorate this process. This script automates the categorization and formatting, so all I have to do is input the excel spreadsheet of the day.

## Tech Used
- Python
- pandas
- openpyxl

## Sample Output
See `formatted_report_20260722.xlsx` and `formatted_report_20260723.xlsx` for example outputs.
