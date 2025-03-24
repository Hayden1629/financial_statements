

This repository contains tools for financial statement analysis, consolidation, and valuation. It helps process quarterly reports, create consolidated statements, and perform discounted cash flow (DCF) valuation. I designed it to use the excel files exported from Godel Terminal.

## Features

### Financial Statement Consolidator (V2)
- First download quarterly and yearly balance sheet, income statement, and cashflow statement from godel terminal
- Put all your statements into a folder with the name as your company ticker
- run consolidator2.py and pass the ticker in as the argument
- Run this program to consolidate to one sheet 
- Find number of shares outstanding for your company and put it at the bottom of the file in any column with the row header "Shares Outstanding" if your want the program to pick up that value
- Consolidates balance sheets, income statements, and cash flow statements into a single Excel worksheet
- Automatically extracts Q4 figures from annual statements by comparing with Q1-Q3 data
- Displays all financial data in chronological order by quarter

### DCF Valuation Calculator
- Load in consolidated excel file
- Calculate enterprise and equity value using DCF methodology
- Auto-populates financial metrics from historical data, but make sure to verify them off the most recent filing
- Customizable forecast parameters:
  - Revenue growth
  - Operating margins
  - Tax rates
  - Capital expenditures
  - Working capital requirements
- Interactive visualization of forecasted cash flows
- Reverse DCF functionality to calculate implied discount rate from current stock price
- Detailed output with valuation summary and calculation breakdown

## Requirements

- Python 3.6 or higher
- Dependencies listed in `requirements.txt`
