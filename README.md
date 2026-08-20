# excel_automation
Excel Automation is a Python-based project using OpenPyXL to read Excel data, calculate row-wise totals, add computed columns, and save the processed workbook automatically. It streamlines repetitive spreadsheet tasks and improves efficiency.


## Excel Automation

A Python-based Excel automation project using OpenPyXL to generate realistic sales data, perform spreadsheet calculations, and automate workbook processing.

## Features
* Generates a sample Excel sales dataset with 2,000 records
* Creates customer, product, category, date, quantity, price, total, and city data
* Applies basic Excel formatting and column sizing
* Reads and processes existing Excel data
* Automatically calculates Total = Quantity × Price
* Adds calculated values to the workbook
* Saves the processed data as a new Excel file

## Technologies Used

* Python 3
* OpenPyXL – Excel workbook creation and automation
* Random – Sample data generation
* Datetime – Date generation and manipulation

## Installation 

Install the required dependency:

```
pip install -r requirements.txt
```

## Usage
1. Generate Sample Data

Run:

python create_data.py


This creates input.xlsx containing 2,000 sample sales records.

2. Automate the Excel File

Run:

```
python excel_automation.py
```


The script reads input.xlsx, processes the worksheet, adds a calculated Total column, and saves the result as output.xlsx.

## Workflow

create_data.py
      │
      ▼
  input.xlsx
      │
      ▼
excel_automation.py
      │
      ▼
  output.xlsx

## Example Calculation

The automation script creates the following Excel formula:

=B2*C2


The formula is applied row-by-row to calculate the total based on the quantity and price columns.

Output

After successful execution, the project generates:

input.xlsx – Sample sales dataset
output.xlsx – Automated and processed Excel workbook

The output workbook can be opened directly in Microsoft Excel or compatible spreadsheet applications.

## Requirements

* Python 3.8+

* OpenPyXL

Install dependencies with:

pip install -r requirements.txt

## Use Cases

This project demonstrates practical Excel automation concepts that can be extended to :

* Sales reporting
* Data processing
* Automated calculations
* Spreadsheet generation
* Business reporting workflows
* Bulk Excel data manipulation
* License

This project is available for educational and personal use.
