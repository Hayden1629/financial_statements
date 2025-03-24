# Financial Statement Consolidator

This program consolidates quarterly financial statements by combining quarterly data (Q1-Q3) with annual data to extract Q4 values and align all quarters for each year in a single consolidated view.

## Features

- Consolidates balance sheets, income statements, and cash flow statements into a single Excel worksheet
- Automatically extracts Q4 figures from annual statements by comparing with Q1-Q3 data
- Displays all financial data in chronological order by quarter
- Detects company ticker from filenames and includes in the output file name
- Formats output with appropriate styling for better readability

## Requirements

- Python 3.6 or higher
- Dependencies listed in `requirements.txt`

## Installation

1. Install the required dependencies:
```
pip install -r requirements.txt
```

## Usage

1. Place your financial statement Excel files in the `statements` directory with the following naming convention:
   - Files for balance sheets should include 'balance_sheet' in the filename
   - Files for income statements should include 'income_statement' in the filename
   - Files for cash flow statements should include 'cash_flow' in the filename
   - Files for yearly statements should start with 'FY_'
   - Files for quarterly statements should start with 'QTR_'
   - Include company ticker at the end (e.g., FY_balance_sheet_EQ_CL.xlsx)

2. Run the consolidator script:
```
python consolidator.py
```

3. The consolidated statements will be saved to `consolidated_statements_[TICKER].xlsx` in the same directory.

## Output

The output Excel workbook will contain a single sheet with:
- Balance Sheet section (top)
- Income Statement section (middle)
- Cash Flow section (bottom)

All data is aligned by quarters, with columns labeled as "Q1 YYYY", "Q2 YYYY", etc. 