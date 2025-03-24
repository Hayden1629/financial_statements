import os
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import re

class FinancialStatementConsolidator:
    def __init__(self, statements_dir):
        self.statements_dir = statements_dir
        self.files = {
            'balance_sheet': {
                'FY': None,
                'QTR': None
            },
            'income_statement': {
                'FY': None,
                'QTR': None
            },
            'cash_flow': {
                'FY': None,
                'QTR': None
            }
        }
        self.company_ticker = None
        self._load_files()

    def _load_files(self):
        """Load all Excel files and categorize them."""
        for filename in os.listdir(self.statements_dir):
            if not filename.endswith('.xlsx') or filename.startswith('~'):
                continue
                
            period_type = filename.split('_')[0]  # FY or QTR
            
            # Extract company ticker from filename
            ticker_match = re.search(r'_([A-Z]+)\.xlsx$', filename, re.IGNORECASE)
            if ticker_match and not self.company_ticker:
                self.company_ticker = ticker_match.group(1)
            
            if 'balance_sheet' in filename.lower():
                statement_type = 'balance_sheet'
            elif 'income_statement' in filename.lower():
                statement_type = 'income_statement'
            elif 'cash_flow' in filename.lower():
                statement_type = 'cash_flow'
            else:
                continue
                
            file_path = os.path.join(self.statements_dir, filename)
            self.files[statement_type][period_type] = file_path
            
        for statement_type, periods in self.files.items():
            print(f"{statement_type}: {periods}")
        print(f"Company ticker: {self.company_ticker}\n")

    def _read_excel(self, file_path):
        """Read Excel file with proper handling of headers."""
        if file_path is None:
            return None
            
        if not os.path.exists(file_path):
            print(f"Warning: File does not exist: {file_path}")
            return None
            
        try:
            # Try reading the Excel file
            df = pd.read_excel(file_path)
            
            # Check if we have unnamed columns
            unnamed_cols = [col for col in df.columns if 'Unnamed' in str(col)]
            
            # If we have unnamed columns, try reading with header on row 2
            if len(unnamed_cols) > 0:
                print(f"Detected unnamed columns in {os.path.basename(file_path)}, trying with header row 2")
                df = pd.read_excel(file_path, header=1)
            
            # Debug info
            print(f"Successfully read {os.path.basename(file_path)}, shape: {df.shape}")
            print(f"Columns: {df.columns.tolist()}")
            
            return df
        except Exception as e:
            print(f"Error reading file {file_path}: {str(e)}")
            return None

    def _extract_years_from_cols(self, df):
        """Extract year information from dataframe columns."""
        years = []
        for col in df.columns[1:]:  # Skip the first column (account names)
            # Try to find a year in the column name
            year_match = re.search(r'20\d{2}', str(col))
            if year_match:
                years.append(year_match.group(0))
            else:
                # If no year found, just use a placeholder
                years.append('Unknown')
        
        return years

    def consolidate_statements(self):
        """Create a single consolidated dataframe with all statements."""
        # Step 1: Read all files
        data = {}
        for statement_type in ['balance_sheet', 'income_statement', 'cash_flow']:
            data[statement_type] = {
                'QTR': self._read_excel(self.files[statement_type]['QTR']),
                'FY': self._read_excel(self.files[statement_type]['FY'])
            }
        
        # Extract years from the data for reference
        all_years = set()
        for statement_type in data:
            for period_type in data[statement_type]:
                if data[statement_type][period_type] is not None:
                    years = self._extract_years_from_cols(data[statement_type][period_type])
                    all_years.update(years)
        
        all_years = sorted(list(all_years))
        print(f"Found years: {all_years}")
        
        # Placeholder for the consolidated dataframe
        consolidated_df = pd.DataFrame()
        
        # Placeholder for section ranges - will be needed for styling
        section_ranges = {}
        
        # List of quarterly dataframes to process
        quarterly_dataframes = [data['balance_sheet']['QTR'], data['income_statement']['QTR'], data['cash_flow']['QTR']]
        
        # TODO: Fill consolidated_df and section_ranges
        for i, df in enumerate(quarterly_dataframes):
            if df is None:
                continue
            
            # Create a dictionary to track where each quarter should be inserted
            account_col = df.columns[0]  # First column is account names
            data_cols = list(df.columns[1:])  # All other columns are data
            
            for year in all_years:
                # Create a list of quarters we need for this year
                quarters = ['Q1', 'Q2', 'Q3', 'Q4']
                
                # Check which quarters already exist
                existing_qtrs = [q for q in quarters if f"{q} {year}" in data_cols]
                
                # Find quarters that need to be added
                for quarter in quarters:
                    qtr_col = f"{quarter} {year}"
                    if qtr_col not in data_cols:
                        # Determine where to insert it
                        qtr_num = int(quarter[1])  # Extract the quarter number
                        
                        # Find position based on order of quarters (1, 2, 3, 4)
                        insert_pos = 1  # Start after the account column
                        for existing_col in data_cols:
                            # Check if the existing column is for a year/quarter before our target
                            if existing_col.endswith(year):
                                existing_qtr = existing_col.split()[0]
                                if existing_qtr in quarters:
                                    existing_qtr_num = int(existing_qtr[1])
                                    if existing_qtr_num < qtr_num:
                                        # This existing quarter comes before our target
                                        insert_pos = data_cols.index(existing_col) + 2  # +1 for index, +1 for account col
                        
                        # Insert the missing quarter column at the calculated position
                        print(f"Inserting {qtr_col} at position {insert_pos}")
                        df.insert(loc=insert_pos, column=qtr_col, value=None)
                        
                        # Update data_cols to reflect the new column
                        data_cols.insert(insert_pos-1, qtr_col)

        # For each row in the quarterly balance sheet
        for row_idx, row in data['balance_sheet']['QTR'].iterrows():
            account_name = row[data['balance_sheet']['QTR'].columns[0]]  # Get account name
            
            # Look for this account in the yearly data
            if data['balance_sheet']['FY'] is not None:
                yearly_df = data['balance_sheet']['FY']
                yearly_col = yearly_df.columns[0]  # Account column name
                yearly_matches = yearly_df[yearly_df[yearly_col] == account_name]
                
                if not yearly_matches.empty:
                    # Get the matching row from yearly data
                    yearly_row = yearly_matches.iloc[0]
                    
                    # For each year column in yearly data
                    for i, col in enumerate(yearly_df.columns[1:], 1):
                        if 'FY' in col:
                            year = re.search(r'20\d{2}', col).group(0) if re.search(r'20\d{2}', col) else None
                            if year:
                                q4_col = f"Q4 {year}"
                                
                                # If Q4 column exists in quarterly data
                                if q4_col in data['balance_sheet']['QTR'].columns:
                                    # Copy the yearly value to Q4
                                    data['balance_sheet']['QTR'].at[row_idx, q4_col] = yearly_row.iloc[i]

        # Now simply concatenate all dataframes to create the consolidated dataframe
        row_counter = 0
        dfs_to_concat = []
        modified_dfs = {}  # Store references to the modified dataframes
        
        # Make sure all dataframes have consistent column names first
        for statement_type in ['balance_sheet', 'income_statement', 'cash_flow']:
            if data[statement_type]['QTR'] is not None:
                # Rename first column to 'Account' for consistency
                first_col = data[statement_type]['QTR'].columns[0]
                data[statement_type]['QTR'].rename(columns={first_col: 'Account'}, inplace=True)
                modified_dfs[statement_type] = data[statement_type]['QTR']
        
        # Track section ranges for styling
        for statement_type in ['balance_sheet', 'income_statement', 'cash_flow']:
            df = modified_dfs.get(statement_type)
            if df is not None:
                # Record section range
                start_row = row_counter
                row_counter += len(df)
                end_row = row_counter - 1
                section_ranges[statement_type] = (start_row, end_row)
                dfs_to_concat.append(df)
                print(f"Added {statement_type} rows {start_row}-{end_row}")
        
        # Concatenate all dataframes
        if dfs_to_concat:
            consolidated_df = pd.concat(dfs_to_concat, ignore_index=True)
            print(f"Created consolidated dataframe with shape: {consolidated_df.shape}")
        else:
            print("No dataframes to concatenate")
        
        # Save the consolidated dataframe
        self._save_consolidated_workbook(consolidated_df, section_ranges)
        
        return consolidated_df

    def _save_consolidated_workbook(self, df, section_ranges):
        """Save the consolidated dataframe to an Excel workbook."""
        if df is None or df.empty:
            print("No data to save")
            return
        
        # Debug print section ranges    
        print("Section ranges for styling:")
        for section, (start, end) in section_ranges.items():
            print(f"  {section}: rows {start}-{end}")
            
        output_path = os.path.join(os.path.dirname(self.files['balance_sheet']['FY']), 
                               f"consolidated_statements_{self.company_ticker}.xlsx" if self.company_ticker else "consolidated_statements.xlsx")
        
        # Create workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Consolidated Statements"
        
        # Define styles
        header_style = Font(bold=True)
        section_colors = {
            'balance_sheet': "C6EFCE",  # Light green
            'income_statement': "FFEB9C",  # Light yellow
            'cash_flow': "FFC7CE"  # Light red
        }
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Start writing data
        current_row = 1
        
        # Write title in first column only
        title = f"Consolidated Financial Statements - {self.company_ticker}" if self.company_ticker else "Consolidated Financial Statements"
        ws.cell(row=current_row, column=1, value=title)
        
        # No merging cells - apply styling only to the first cell
        title_cell = ws.cell(row=current_row, column=1)
        title_cell.font = Font(size=14, bold=True)
        title_cell.alignment = Alignment(horizontal="left")
        
        current_row += 2  # Add space after title
        
        # Write column headers
        for col_idx, col_name in enumerate(df.columns, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=col_name)
            cell.font = header_style
            cell.border = thin_border
            cell.fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
        
        current_row += 1
        
        # Pre-determine which rows belong to which section
        row_section_map = {}
        for row_idx in range(len(df)):
            for section, (start, end) in section_ranges.items():
                if start <= row_idx <= end:
                    row_section_map[row_idx] = section
                    break
        
        # Track when we add section headers
        section_headers_added = set()
        
        # Write rows with section styling
        for row_idx, row in df.iterrows():
            # Get section type from our pre-computed map
            section_type = row_section_map.get(row_idx)
            
            # If this is the first row of a section, add a section header
            if section_type and section_type not in section_headers_added:
                # Insert a section header
                header_row = current_row
                section_name = section_type.replace('_', ' ').title()
                ws.cell(row=header_row, column=1, value=section_name)
                
                # Apply formatting to just the first cell (no cell merging)
                header_cell = ws.cell(row=header_row, column=1)
                header_cell.font = Font(bold=True)
                header_cell.fill = PatternFill(start_color=section_colors[section_type], 
                                              end_color=section_colors[section_type], 
                                              fill_type="solid")
                header_cell.alignment = Alignment(horizontal="left")
                header_cell.border = thin_border
                
                # Add empty cells for the rest of the row
                for col_idx in range(2, len(df.columns) + 1):
                    empty_cell = ws.cell(row=header_row, column=col_idx, value="")
                    empty_cell.border = thin_border
                
                current_row += 1
                section_headers_added.add(section_type)
                print(f"Added section header for {section_type} at Excel row {header_row}")
            
            # Write data row
            for col_idx, col_name in enumerate(df.columns, 1):
                value = row[col_name]
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.border = thin_border
                
                # Format numbers
                if col_idx > 1 and value is not None:  # Skip account name column
                    try:
                        float_val = float(value)
                        cell.number_format = '#,##0.00'
                    except (ValueError, TypeError):
                        pass
            
            current_row += 1
        
        # Adjust column widths
        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # Save the workbook
        wb.save(output_path)
        print(f"Consolidated statements saved to {output_path}")

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    statements_dir = os.path.join(script_dir, 'statements')
    
    consolidator = FinancialStatementConsolidator(statements_dir)
    consolidator.consolidate_statements()
    
    print("Financial statement consolidation complete!")

if __name__ == "__main__":
    main()