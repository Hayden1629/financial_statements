import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re
import openpyxl

class DCFValuationCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("DCF Valuation Calculator")
        self.root.geometry("1200x800")
        self.root.minsize(1200, 800)
        
        self.df = None
        self.latest_year_data = {}
        self.forecast_years = 5
        
        self.create_widgets()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Financial Statement Selection", padding=10)
        file_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(file_frame, text="Select Consolidated Statement", command=self.load_file).grid(row=0, column=0, padx=5, pady=5)
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Notebook for different sections
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Historical Data Tab
        self.hist_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.hist_frame, text="Historical Data")
        
        # Forecast Tab
        self.forecast_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.forecast_frame, text="Forecast Parameters")
        
        # Create forecast input frames
        self.create_forecast_inputs()
        
        # DCF Value Tab
        self.dcf_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(self.dcf_frame, text="DCF Valuation")
        
        # Calculate button - place it in a separate frame at the bottom
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, padx=5, pady=10)
        self.calculate_button = ttk.Button(button_frame, text="Calculate Valuation", command=self.calculate_valuation)
        self.calculate_button.pack(padx=5, pady=5)
    
    def create_forecast_inputs(self):
        # Left frame for inputs
        left_frame = ttk.Frame(self.forecast_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Add date range selector frame at the top
        date_range_frame = ttk.LabelFrame(left_frame, text="Data Range for Calculations", padding=10)
        date_range_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Quarters to use for calculations
        ttk.Label(date_range_frame, text="Use data from the last:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        
        # Dropdown for selecting number of quarters
        self.quarters_var = tk.StringVar(value="All available data")
        quarters_options = ["All available data", "4 quarters (1 year)", "8 quarters (2 years)", "12 quarters (3 years)"]
        quarters_dropdown = ttk.Combobox(date_range_frame, textvariable=self.quarters_var, values=quarters_options, width=20)
        quarters_dropdown.grid(row=0, column=1, padx=5, pady=5)
        
        # Bind the dropdown to update parameters when changed
        quarters_dropdown.bind("<<ComboboxSelected>>", lambda e: self.recalculate_stats())
        
        # Button to update parameters with selected date range
        refresh_button = ttk.Button(date_range_frame, text="Refresh Parameters", 
                                  command=lambda: self.recalculate_stats())
        refresh_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Growth assumptions
        growth_frame = ttk.LabelFrame(left_frame, text="Growth & Margin Assumptions", padding=10)
        growth_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create labels dictionary to store references for auto-calculated indicators
        self.auto_calc_labels = {}
        
        # Revenue Growth
        ttk.Label(growth_frame, text="Revenue Growth Rate (YoY %):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.revenue_growth = ttk.Entry(growth_frame)
        self.revenue_growth.grid(row=0, column=1, padx=5, pady=5)
        self.revenue_growth.insert(0, "5.0")
        self.auto_calc_labels["revenue_growth"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["revenue_growth"].grid(row=0, column=2, sticky="w", padx=5, pady=5)
        
        # Operating Margin (% of Revenue)
        ttk.Label(growth_frame, text="Operating Margin (% of Revenue):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.operating_margin = ttk.Entry(growth_frame)
        self.operating_margin.grid(row=1, column=1, padx=5, pady=5)
        self.operating_margin.insert(0, "20.0")
        self.auto_calc_labels["operating_margin"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["operating_margin"].grid(row=1, column=2, sticky="w", padx=5, pady=5)
        
        # Tax Rate
        ttk.Label(growth_frame, text="Tax Rate (%):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.tax_rate = ttk.Entry(growth_frame)
        self.tax_rate.grid(row=2, column=1, padx=5, pady=5)
        self.tax_rate.insert(0, "25.0")
        self.auto_calc_labels["tax_rate"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["tax_rate"].grid(row=2, column=2, sticky="w", padx=5, pady=5)
        
        # CapEx % of Revenue
        ttk.Label(growth_frame, text="CapEx (% of Revenue):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.capex_percent = ttk.Entry(growth_frame)
        self.capex_percent.grid(row=3, column=1, padx=5, pady=5)
        self.capex_percent.insert(0, "3.0")
        self.auto_calc_labels["capex_percent"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["capex_percent"].grid(row=3, column=2, sticky="w", padx=5, pady=5)
        
        # Working Capital % of Revenue
        ttk.Label(growth_frame, text="Working Capital (% of Revenue):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.wc_percent = ttk.Entry(growth_frame)
        self.wc_percent.grid(row=4, column=1, padx=5, pady=5)
        self.wc_percent.insert(0, "5.0")
        self.auto_calc_labels["wc_percent"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["wc_percent"].grid(row=4, column=2, sticky="w", padx=5, pady=5)
        
        # DCF parameters
        dcf_frame = ttk.LabelFrame(left_frame, text="DCF Parameters", padding=10)
        dcf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Discount Rate (WACC)
        ttk.Label(dcf_frame, text="Discount Rate (%):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.discount_rate = ttk.Entry(dcf_frame)
        self.discount_rate.grid(row=0, column=1, padx=5, pady=5)
        self.discount_rate.insert(0, "10.0")
        
        # Terminal Growth Rate
        ttk.Label(dcf_frame, text="Terminal Growth Rate (%):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.terminal_growth = ttk.Entry(dcf_frame)
        self.terminal_growth.grid(row=1, column=1, padx=5, pady=5)
        self.terminal_growth.insert(0, "2.0")
        
        # Forecast Years
        ttk.Label(dcf_frame, text="Forecast Years:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.forecast_years_entry = ttk.Entry(dcf_frame)
        self.forecast_years_entry.grid(row=2, column=1, padx=5, pady=5)
        self.forecast_years_entry.insert(0, "5")
        
        # Shares Outstanding
        ttk.Label(dcf_frame, text="Shares Outstanding (millions):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.shares_outstanding = ttk.Entry(dcf_frame)
        self.shares_outstanding.grid(row=3, column=1, padx=5, pady=5)
        self.shares_outstanding.insert(0, "100.0")
        self.auto_calc_labels["shares_outstanding"] = ttk.Label(dcf_frame, text="", foreground="green")
        self.auto_calc_labels["shares_outstanding"].grid(row=3, column=2, sticky="w", padx=5, pady=5)
        
        # Current Debt
        ttk.Label(dcf_frame, text="Current Debt (millions):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.current_debt = ttk.Entry(dcf_frame)
        self.current_debt.grid(row=4, column=1, padx=5, pady=5)
        self.current_debt.insert(0, "0.0")
        self.auto_calc_labels["current_debt"] = ttk.Label(dcf_frame, text="", foreground="green")
        self.auto_calc_labels["current_debt"].grid(row=4, column=2, sticky="w", padx=5, pady=5)
        
        # Cash & Equivalents
        ttk.Label(dcf_frame, text="Cash & Equivalents (millions):").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.cash_equivalents = ttk.Entry(dcf_frame)
        self.cash_equivalents.grid(row=5, column=1, padx=5, pady=5)
        self.cash_equivalents.insert(0, "0.0")
        self.auto_calc_labels["cash_equivalents"] = ttk.Label(dcf_frame, text="", foreground="green")
        self.auto_calc_labels["cash_equivalents"].grid(row=5, column=2, sticky="w", padx=5, pady=5)
        
        # Right frame for preview
        right_frame = ttk.Frame(self.forecast_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Historical stats frame
        self.stats_frame = ttk.LabelFrame(right_frame, text="Historical Statistics", padding=10)
        self.stats_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Historical stats
        self.hist_stats = tk.Text(self.stats_frame, height=20, width=40)
        self.hist_stats.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.hist_stats.configure(state='disabled')
    
    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if file_path:
            try:
                self.file_label.config(text=file_path)
                
                # Check file extension and load accordingly
                if file_path.endswith(('.xlsx', '.xls')):
                    # For Excel files, don't specify a header row initially
                    self.df = pd.read_excel(file_path, header=None)
                else:
                    self.df = pd.read_csv(file_path, skipinitialspace=True)
                
                # Clean the data
                self.clean_data()
                
                # Display historical data
                self.display_historical_data()
                
                # Calculate and display historical stats
                self.calculate_historical_stats()
                
                # Pre-fill forecast parameters from historical data using all available quarters
                self.prefill_forecast_parameters(self.quarter_cols)
                
                # Initialize quarters dropdown with proper options based on available data
                if hasattr(self, 'quarter_cols') and len(self.quarter_cols) > 0:
                    num_quarters = len(self.quarter_cols)
                    quarters_options = ["All available data"]
                    
                    if num_quarters >= 4:
                        quarters_options.append("4 quarters (1 year)")
                    if num_quarters >= 8:
                        quarters_options.append("8 quarters (2 years)")
                    if num_quarters >= 12:
                        quarters_options.append("12 quarters (3 years)")
                    
                    # Update dropdown options
                    quarters_dropdown = None
                    for child in self.forecast_frame.winfo_children():
                        if isinstance(child, ttk.Frame):
                            for frame_child in child.winfo_children():
                                if isinstance(frame_child, ttk.Labelframe) and frame_child.winfo_children():
                                    for combobox in frame_child.winfo_children():
                                        if isinstance(combobox, ttk.Combobox):
                                            quarters_dropdown = combobox
                                            break
                    
                    if quarters_dropdown:
                        quarters_dropdown['values'] = quarters_options
                
                # Switch to forecast tab
                self.notebook.select(1)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load the file: {str(e)}")
                import traceback
                traceback.print_exc()
    
    def clean_data(self):
        # Skip the first two rows which contain header information and set row 3 (index 2) as column headers
        if not self.df.empty:
            try:
                # Find where the quarter headers are (usually row 3, index 2)
                header_row = None
                for i in range(5):  # Check first 5 rows
                    row = self.df.iloc[i]
                    # Look for cells that might contain quarter formatting like "Q1 2023"
                    quarter_pattern = [str(cell).strip() for cell in row if isinstance(cell, str) and re.match(r'Q\d 20\d\d', str(cell).strip())]
                    if quarter_pattern:
                        header_row = i
                        break
                
                if header_row is None:
                    header_row = 2  # Default to row 3 (index 2) if no quarter headers found
                    
                # Use the identified row as column headers
                header_data = self.df.iloc[header_row]
                self.df.columns = [str(x).strip() if isinstance(x, str) else x for x in header_data]
                
                # Skip rows up to and including the header row
                self.df = self.df.iloc[header_row+1:].reset_index(drop=True)
                
                # Replace empty strings with NaN
                self.df = self.df.replace(['', 'nan', 'NaN', 'None'], np.nan)
                
                # Find the Account column (first column containing text data)
                account_col = None
                for col in self.df.columns:
                    if self.df[col].dtype == 'object':  # Look for string/object column
                        account_col = col
                        break
                
                if account_col is None:
                    account_col = self.df.columns[0]  # Default to first column
                    
                # Set the account column as index
                self.df.set_index(account_col, inplace=True)
                
                # Convert numeric columns to float
                for col in self.df.columns:
                    try:
                        self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
                    except:
                        pass
                
                # Convert quarter column headers to date format if they are in "Q# YYYY" format
                self.quarter_cols = []
                for col in self.df.columns:
                    if isinstance(col, str) and re.match(r'Q\d 20\d\d', col):
                        self.quarter_cols.append(col)
                
                # Sort quarter columns chronologically
                self.quarter_cols = sorted(self.quarter_cols, 
                                          key=lambda x: (int(re.search(r'20(\d\d)', x).group(1)), 
                                                         int(re.search(r'Q(\d)', x).group(1))))
                
                # If we found quarter columns, print diagnostics
                if self.quarter_cols:
                    print(f"Found {len(self.quarter_cols)} quarter columns: {self.quarter_cols}")
                else:
                    print("Warning: No quarter columns (Q# YYYY format) found in the data")
            
            except Exception as e:
                messagebox.showerror("Error", f"Error cleaning data: {str(e)}")
                import traceback
                traceback.print_exc()
    
    def display_historical_data(self):
        # Clear existing widgets
        for widget in self.hist_frame.winfo_children():
            widget.destroy()
        
        if self.df is not None:
            # Create a treeview for the data
            tree = ttk.Treeview(self.hist_frame)
            tree.pack(fill=tk.BOTH, expand=True)
            
            # Add vertical scrollbar
            vsb = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
            vsb.pack(side='right', fill='y')
            tree.configure(yscrollcommand=vsb.set)
            
            # Add horizontal scrollbar
            hsb = ttk.Scrollbar(self.hist_frame, orient="horizontal", command=tree.xview)
            hsb.pack(side='bottom', fill='x')
            tree.configure(xscrollcommand=hsb.set)
            
            # Define columns
            df_display = self.df.reset_index()
            tree["columns"] = list(df_display.columns)
            
            # Format columns
            tree.column("#0", width=0, stretch=tk.NO)
            for col in df_display.columns:
                tree.column(col, anchor=tk.W, width=100)
                tree.heading(col, text=str(col), anchor=tk.W)  # Convert to string in case column name is a date
            
            # Add data to treeview
            for i, row in df_display.iterrows():
                tree.insert("", i, text="", values=[str(val) if pd.isna(val) else val for val in row])
    
    def recalculate_stats(self, event=None):
        """Recalculate statistics and prefill parameters based on selected date range"""
        if not hasattr(self, 'quarter_cols') or not self.quarter_cols:
            messagebox.showwarning("Warning", "No quarterly data available")
            return
        
        # Get the selected number of quarters
        selection = self.quarters_var.get()
        
        if selection.startswith("4"):
            quarters_to_use = min(4, len(self.quarter_cols))
        elif selection.startswith("8"):
            quarters_to_use = min(8, len(self.quarter_cols))
        elif selection.startswith("12"):
            quarters_to_use = min(12, len(self.quarter_cols))
        elif selection.startswith("16"):
            quarters_to_use = min(16, len(self.quarter_cols))
        else:  # "All"
            quarters_to_use = len(self.quarter_cols)
        
        # Store the selected range
        self.selected_quarters = self.quarter_cols[-quarters_to_use:] if quarters_to_use > 0 else self.quarter_cols
        
        # Update historical stats display
        self.calculate_historical_stats(self.selected_quarters)
        
        # Prefill parameters based on the selected range
        self.prefill_forecast_parameters(self.selected_quarters)
    
    def calculate_historical_stats(self, selected_cols=None):
        if self.df is None:
            return
        
        # Enable text widget for editing
        self.hist_stats.configure(state='normal')
        self.hist_stats.delete(1.0, tk.END)
        
        try:
            # Check if we have quarter columns identified
            if not hasattr(self, 'quarter_cols') or not self.quarter_cols:
                self.hist_stats.insert(tk.END, "Error: No quarterly data columns found.\n")
                self.hist_stats.configure(state='disabled')
                return
            
            # Use selected columns or default to all
            if selected_cols is None:
                # Use the last 20 quarters or all available if less
                latest_cols = self.quarter_cols[-min(20, len(self.quarter_cols)):]
            else:
                latest_cols = selected_cols
            
            if not latest_cols:
                self.hist_stats.insert(tk.END, "Error: Could not identify quarterly data columns.\n")
                self.hist_stats.configure(state='disabled')
                return
            
            self.hist_stats.insert(tk.END, f"Using {len(latest_cols)} quarters of data: {latest_cols[0]} to {latest_cols[-1]}\n\n")
            
            # Store latest quarter
            self.latest_quarter = latest_cols[-1] if latest_cols else None
            
            # Check if we have revenue data
            if 'Revenue' in self.df.index:
                # Group revenue data by year
                revenue_data = {}
                for col in latest_cols:
                    try:
                        value = self.df.loc['Revenue', col]
                        if pd.notna(value):
                            # Extract year and quarter
                            year_match = re.search(r'20(\d\d)', col)
                            quarter_match = re.search(r'Q(\d)', col)
                            if year_match and quarter_match:
                                year = int("20" + year_match.group(1))
                                quarter = int(quarter_match.group(1))
                                
                                # Initialize year in dictionary if not exists
                                if year not in revenue_data:
                                    revenue_data[year] = {}
                                    
                                # Store revenue for this quarter
                                revenue_data[year][quarter] = value
                    except Exception as e:
                        print(f"Error processing {col}: {str(e)}")
                
                # Calculate annual revenue by summing quarters for each year
                annual_revenue = {}
                for year, quarters in revenue_data.items():
                    # Only include years with all 4 quarters
                    if len(quarters) == 4:
                        annual_revenue[year] = sum(quarters.values())
                
                # Calculate year-over-year growth rates
                years = sorted(annual_revenue.keys())
                if len(years) >= 2:
                    growth_rates = []
                    for i in range(1, len(years)):
                        prev_year = years[i-1]
                        curr_year = years[i]
                        if annual_revenue[prev_year] > 0:
                            growth_rate = (annual_revenue[curr_year] / annual_revenue[prev_year] - 1) * 100
                            growth_rates.append(growth_rate)
                    
                    if growth_rates:
                        # Display each year's revenue and growth
                        self.hist_stats.insert(tk.END, "Annual Revenue:\n")
                        for i, year in enumerate(years):
                            self.hist_stats.insert(tk.END, f"{year}: ${annual_revenue[year]:.2f}M")
                            if i > 0:
                                growth = (annual_revenue[year] / annual_revenue[years[i-1]] - 1) * 100
                                self.hist_stats.insert(tk.END, f" (YoY: {growth:.2f}%)")
                            self.hist_stats.insert(tk.END, "\n")
                        
                        # Display average annual growth rate
                        avg_growth = np.mean(growth_rates)
                        self.hist_stats.insert(tk.END, f"\nAverage Annual Revenue Growth: {avg_growth:.2f}%\n\n")
                else:
                    self.hist_stats.insert(tk.END, "Insufficient complete years for Revenue Growth calculation\n\n")
            else:
                self.hist_stats.insert(tk.END, "Revenue data not found\n\n")
            
            # Calculate average operating margin
            if 'Operating Income' in self.df.index and 'Revenue' in self.df.index:
                operating_income = self.df.loc['Operating Income', latest_cols].dropna()
                revenue = self.df.loc['Revenue', latest_cols].dropna()
                
                # Match indices
                common_cols = operating_income.index.intersection(revenue.index)
                operating_income = operating_income[common_cols]
                revenue = revenue[common_cols]
                
                if len(operating_income) > 0 and len(revenue) > 0:
                    margins = (operating_income / revenue) * 100
                    avg_margin = np.mean(margins)
                    self.hist_stats.insert(tk.END, f"Average Operating Margin: {avg_margin:.2f}%\n\n")
                else:
                    self.hist_stats.insert(tk.END, "Insufficient data for Operating Margin calculation\n\n")
            else:
                self.hist_stats.insert(tk.END, "Operating Income or Revenue data not found\n\n")
            
            # Calculate average tax rate
            if 'Income Taxes' in self.df.index and 'Pretax Income' in self.df.index:
                taxes = self.df.loc['Income Taxes', latest_cols].dropna()
                pretax = self.df.loc['Pretax Income', latest_cols].dropna()
                
                # Match indices
                common_cols = taxes.index.intersection(pretax.index)
                taxes = taxes[common_cols]
                pretax = pretax[common_cols]
                
                if len(taxes) > 0 and len(pretax) > 0:
                    rates = (taxes / pretax) * 100
                    avg_tax_rate = np.mean(rates)
                    self.hist_stats.insert(tk.END, f"Average Tax Rate: {avg_tax_rate:.2f}%\n\n")
                else:
                    self.hist_stats.insert(tk.END, "Insufficient data for Tax Rate calculation\n\n")
            else:
                self.hist_stats.insert(tk.END, "Income Taxes or Pretax Income data not found\n\n")
            
            # Store latest financial data
            self.latest_year_data = {}
            
            # Define financial keys with alternative names
            financial_keys = {
                'Revenue': ['Revenue'],
                'Operating Income': ['Operating Income'],
                'Income Taxes': ['Income Taxes'],
                'Pretax Income': ['Pretax Income'],
                'Current Assets': ['Current Assets'],
                'Current Liabilities': ['Current Liabilities'],
                'Cash & Equivalents': ['Cash & Equivalents', 'Cash and Equivalents', 'Cash and Cash Equivalents'],
                'Long Term Debt': ['Long Term Debt', 'Long-Term Debt'],
                'Short Term Debt': ['Short Term Debt', 'Short-Term Debt'],
                'Purchase of PP&E': ['Purchase of PP&E', 'CapEx', 'Capital Expenditure', 'Purchase of Investment', 'Acquisitions']
            }
            
            # Helper function to find most recent non-NaN value with alternative keys
            def find_most_recent_value(key_list):
                for key in key_list:
                    if key in self.df.index:
                        # Try to find a non-NaN value starting from the most recent quarter
                        for col in reversed(latest_cols):
                            try:
                                value = self.df.loc[key, col]
                                if pd.notna(value) and value != 'nan' and value != 'NaN':
                                    return value, key
                            except:
                                continue
                return None, None
            
            # Get most recent values for each financial key group
            for primary_key, alt_keys in financial_keys.items():
                value, found_key = find_most_recent_value(alt_keys)
                if value is not None:
                    self.latest_year_data[primary_key] = value
                    self.hist_stats.insert(tk.END, f"Latest {primary_key}: {value:.2f} (from '{found_key}')\n")
                else:
                    self.hist_stats.insert(tk.END, f"No valid data found for {primary_key}\n")
        
        except Exception as e:
            self.hist_stats.insert(tk.END, f"Error calculating statistics: {str(e)}\n")
            import traceback
            traceback.print_exc()
        
        # Disable text widget
        self.hist_stats.configure(state='disabled')
    
    def get_selected_quarters(self):
        """Get the list of quarters to use based on user selection"""
        if not hasattr(self, 'quarter_cols') or not self.quarter_cols:
            return []
        
        # Get the selected number of quarters
        selection = self.quarters_var.get()
        
        if selection.startswith("4"):
            quarters_to_use = min(4, len(self.quarter_cols))
        elif selection.startswith("8"):
            quarters_to_use = min(8, len(self.quarter_cols))
        elif selection.startswith("12"):
            quarters_to_use = min(12, len(self.quarter_cols))
        else:  # "All available data"
            quarters_to_use = len(self.quarter_cols)
        
        # Return the selected quarters
        return self.quarter_cols[-quarters_to_use:] if quarters_to_use > 0 else self.quarter_cols

    def prefill_forecast_parameters(self, selected_cols=None):
        # Prefill forecast parameters from historical data
        if self.df is not None:
            try:
                # Reset all auto-calculated indicators
                for label in self.auto_calc_labels.values():
                    label.config(text="")
                    
                # Check if we have quarter columns identified
                if not hasattr(self, 'quarter_cols') or not self.quarter_cols:
                    return
                    
                # Use provided columns or get all quarters
                if selected_cols is None or len(selected_cols) == 0:
                    selected_cols = self.quarter_cols
                    
                # Show what range is being used
                quarters_used = f"{selected_cols[0]} to {selected_cols[-1]}" if selected_cols else "No data"
                print(f"Using data range: {quarters_used} ({len(selected_cols)} quarters)")
                
                # Revenue Growth (properly calculated year-over-year)
                if 'Revenue' in self.df.index and len(selected_cols) >= 8:  # Need at least 8 quarters (2 years)
                    # Get revenue data for the selected quarters
                    revenue_data = {}
                    for col in selected_cols:
                        try:
                            value = self.df.loc['Revenue', col]
                            if pd.notna(value):
                                # Extract year and quarter
                                year_match = re.search(r'20(\d\d)', col)
                                quarter_match = re.search(r'Q(\d)', col)
                                if year_match and quarter_match:
                                    year = int("20" + year_match.group(1))
                                    quarter = int(quarter_match.group(1))
                                    
                                    # Initialize year in dictionary if not exists
                                    if year not in revenue_data:
                                        revenue_data[year] = {}
                                        
                                    # Store revenue for this quarter
                                    revenue_data[year][quarter] = value
                        except Exception as e:
                            print(f"Error processing {col}: {str(e)}")
                    
                    # Calculate annual revenue by summing quarters for each year
                    annual_revenue = {}
                    for year, quarters in revenue_data.items():
                        # Only use years with all 4 quarters of data
                        if len(quarters) == 4:
                            annual_revenue[year] = sum(quarters.values())
                    
                    # Sort years and calculate growth rates
                    years = sorted(annual_revenue.keys())
                    if len(years) >= 2:
                        # Calculate all year-over-year growth rates
                        growth_rates = []
                        for i in range(1, len(years)):
                            prev_year = years[i-1]
                            curr_year = years[i]
                            if annual_revenue[prev_year] > 0:
                                growth_rate = ((annual_revenue[curr_year] / annual_revenue[prev_year]) - 1) * 100
                                growth_rates.append(growth_rate)
                        
                        if growth_rates:
                            # Use the average growth rate across all available years
                            avg_growth = sum(growth_rates) / len(growth_rates)
                            self.revenue_growth.delete(0, tk.END)
                            self.revenue_growth.insert(0, f"{avg_growth:.2f}")
                            
                            # Format all growth rates for display
                            growth_text = ""
                            for i in range(1, len(years)):
                                prev_year = years[i-1]
                                curr_year = years[i]
                                growth = ((annual_revenue[curr_year] / annual_revenue[prev_year]) - 1) * 100
                                growth_text += f"{prev_year}-{curr_year}: {growth:.2f}%, "
                            
                            # Remove trailing comma and space
                            if growth_text:
                                growth_text = growth_text[:-2]
                                
                            self.auto_calc_labels["revenue_growth"].config(
                                text=f"(Avg: {avg_growth:.2f}%, {growth_text})"
                            )
                
                # Operating Margin
                if 'Operating Income' in self.df.index and 'Revenue' in self.df.index and len(selected_cols) >= 4:
                    # Calculate average operating margin
                    margins = []
                    for col in selected_cols[-4:]:  # Last 4 quarters from selection
                        try:
                            op_income = self.df.loc['Operating Income', col]
                            revenue = self.df.loc['Revenue', col]
                            if pd.notna(op_income) and pd.notna(revenue) and revenue > 0:
                                margin = (op_income / revenue) * 100
                                margins.append(margin)
                        except:
                            continue
                    
                    if margins:
                        avg_margin = sum(margins) / len(margins)
                        self.operating_margin.delete(0, tk.END)
                        self.operating_margin.insert(0, f"{avg_margin:.2f}")
                        self.auto_calc_labels["operating_margin"].config(text=f"(Avg from {quarters_used})")
                
                # Tax Rate
                if 'Income Taxes' in self.df.index and 'Pretax Income' in self.df.index and len(selected_cols) >= 4:
                    # Calculate average tax rate
                    rates = []
                    for col in selected_cols[-4:]:  # Last 4 quarters from selection
                        try:
                            tax = self.df.loc['Income Taxes', col]
                            pretax = self.df.loc['Pretax Income', col]
                            if pd.notna(tax) and pd.notna(pretax) and pretax > 0:
                                rate = (tax / pretax) * 100
                                rates.append(rate)
                        except:
                            continue
                    
                    if rates:
                        avg_tax_rate = sum(rates) / len(rates)
                        self.tax_rate.delete(0, tk.END)
                        self.tax_rate.insert(0, f"{avg_tax_rate:.2f}")
                        self.auto_calc_labels["tax_rate"].config(text=f"(Avg from {quarters_used})")
                
                # Update other fields using the last selected quarter's data
                self.update_latest_financial_data(selected_cols)
                
            except Exception as e:
                messagebox.showwarning("Warning", f"Error pre-filling parameters: {str(e)}")
                import traceback
                traceback.print_exc()
            
    def update_latest_financial_data(self, selected_cols):
        """Update latest financial data based on the selected columns"""
        if not selected_cols:
            return
        
        # Define financial keys with alternative names
        financial_keys = {
            'Revenue': ['Revenue'],
            'Operating Income': ['Operating Income'],
            'Income Taxes': ['Income Taxes'],
            'Pretax Income': ['Pretax Income'],
            'Current Assets': ['Current Assets'],
            'Current Liabilities': ['Current Liabilities'],
            'Cash & Equivalents': ['Cash & Equivalents', 'Cash and Equivalents', 'Cash and Cash Equivalents'],
            'Long Term Debt': ['Long Term Debt', 'Long-Term Debt'],
            'Short Term Debt': ['Short Term Debt', 'Short-Term Debt'],
            'Purchase of PP&E': ['Purchase of PP&E', 'CapEx', 'Capital Expenditure', 'Purchase of Investment', 'Acquisitions']
        }
        
        # Helper function to find most recent non-NaN value with alternative keys
        def find_most_recent_value(key_list):
            for key in key_list:
                if key in self.df.index:
                    # Try to find a non-NaN value starting from the most recent quarter
                    for col in reversed(selected_cols):
                        try:
                            value = self.df.loc[key, col]
                            if pd.notna(value) and value != 'nan' and value != 'NaN':
                                return value, key, col
                        except:
                            continue
            return None, None, None
        
        # Get most recent values for each financial key group
        for primary_key, alt_keys in financial_keys.items():
            value, found_key, found_col = find_most_recent_value(alt_keys)
            if value is not None:
                self.latest_year_data[primary_key] = value
        
        # Update the fields that use latest financial data
        quarters_used = f"{selected_cols[0]}-{selected_cols[-1]}"
        
        # CapEx
        if 'Purchase of PP&E' in self.latest_year_data and 'Revenue' in self.latest_year_data:
            if self.latest_year_data['Revenue'] > 0:
                capex_ratio = (abs(self.latest_year_data['Purchase of PP&E']) / self.latest_year_data['Revenue']) * 100
                self.capex_percent.delete(0, tk.END)
                self.capex_percent.insert(0, f"{capex_ratio:.2f}")
                self.auto_calc_labels["capex_percent"].config(text=f"(From {quarters_used})")
        
        # Working Capital
        if all(key in self.latest_year_data for key in ['Current Assets', 'Current Liabilities', 'Revenue']):
            if self.latest_year_data['Revenue'] > 0:
                wc = self.latest_year_data['Current Assets'] - self.latest_year_data['Current Liabilities']
                wc_ratio = (wc / self.latest_year_data['Revenue']) * 100
                self.wc_percent.delete(0, tk.END)
                self.wc_percent.insert(0, f"{wc_ratio:.2f}")
                self.auto_calc_labels["wc_percent"].config(text=f"(From {quarters_used})")
        
        # Shares Outstanding
        share_fields = ['Common Stock', 'Common Equity', 'Outstanding Stock']
        for field in share_fields:
            if field in self.df.index:
                # Find most recent value
                for col in reversed(selected_cols):
                    try:
                        shares = self.df.loc[field, col]
                        if pd.notna(shares):
                            self.shares_outstanding.delete(0, tk.END)
                            self.shares_outstanding.insert(0, f"{shares:.2f}")
                            self.auto_calc_labels["shares_outstanding"].config(text=f"(From {field}, {col})")
                            break
                    except:
                        continue
                break  # Stop after we find the first valid field
        
        # Debt
        if 'Long Term Debt' in self.latest_year_data:
            debt = self.latest_year_data['Long Term Debt']
            debt_key, debt_value, debt_col = find_most_recent_value(['Long Term Debt', 'Long-Term Debt'])
            self.current_debt.delete(0, tk.END)
            self.current_debt.insert(0, f"{debt:.2f}")
            self.auto_calc_labels["current_debt"].config(text=f"(From {debt_col})")
        
        # Cash
        if 'Cash & Equivalents' in self.latest_year_data:
            cash = self.latest_year_data['Cash & Equivalents']
            cash_key, cash_value, cash_col = find_most_recent_value(['Cash & Equivalents', 'Cash and Equivalents', 'Cash and Cash Equivalents'])
            self.cash_equivalents.delete(0, tk.END)
            self.cash_equivalents.insert(0, f"{cash:.2f}")
            self.auto_calc_labels["cash_equivalents"].config(text=f"(From {cash_col})")
    
    def calculate_valuation(self):
        try:
            # Get forecast parameters
            self.forecast_years = int(self.forecast_years_entry.get())
            revenue_growth = float(self.revenue_growth.get()) / 100
            operating_margin = float(self.operating_margin.get()) / 100
            tax_rate = float(self.tax_rate.get()) / 100
            capex_percent = float(self.capex_percent.get()) / 100
            wc_percent = float(self.wc_percent.get()) / 100
            discount_rate = float(self.discount_rate.get()) / 100
            terminal_growth = float(self.terminal_growth.get()) / 100
            shares_outstanding = float(self.shares_outstanding.get())
            debt = float(self.current_debt.get())
            cash = float(self.cash_equivalents.get())
            
            # Get latest revenue (annualized from quarterly)
            if 'Revenue' in self.latest_year_data:
                base_revenue = self.latest_year_data['Revenue'] * 4  # Annualize quarterly revenue
            else:
                messagebox.showerror("Error", "Could not find revenue data in the financial statement.")
                return
            
            # Create forecast model
            years = list(range(1, self.forecast_years + 1))
            revenue = [base_revenue * (1 + revenue_growth) ** year for year in years]
            ebit = [rev * operating_margin for rev in revenue]
            tax = [op * tax_rate for op in ebit]
            nopat = [op - tx for op, tx in zip(ebit, tax)]
            
            # Calculate CapEx and Working Capital changes
            capex = [rev * capex_percent for rev in revenue]
            
            # For working capital, we need to calculate the change year over year
            wc = [rev * wc_percent for rev in revenue]
            wc_change = [0] + [wc[i] - wc[i-1] for i in range(1, len(wc))]
            wc_change[0] = wc[0] - (base_revenue * wc_percent)
            
            # Free Cash Flow
            fcf = [nopat[i] - capex[i] - wc_change[i] for i in range(len(nopat))]
            
            # Calculate Terminal Value
            terminal_value = fcf[-1] * (1 + terminal_growth) / (discount_rate - terminal_growth)
            
            # Discounted Cash Flows
            dcf = [flow / (1 + discount_rate) ** year for year, flow in zip(years, fcf)]
            
            # Discounted Terminal Value
            discounted_tv = terminal_value / (1 + discount_rate) ** self.forecast_years
            
            # Enterprise Value
            ev = sum(dcf) + discounted_tv
            
            # Equity Value
            equity_value = ev - debt + cash
            
            # Price per Share
            price_per_share = equity_value / shares_outstanding
            
            # Clear the DCF frame and display results
            for widget in self.dcf_frame.winfo_children():
                widget.destroy()
            
            # Create main container for results and ensure fixed control panel at bottom
            main_container = ttk.Frame(self.dcf_frame)
            main_container.pack(fill=tk.BOTH, expand=True)
            
            # Control panel that will stay at bottom even when resized
            control_panel = ttk.Frame(self.dcf_frame)
            control_panel.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=5)
            
            # Add buttons to the fixed control panel
            recalc_button = ttk.Button(
                control_panel, 
                text="Edit Parameters",
                command=lambda: self.notebook.select(1)
            )
            recalc_button.pack(side=tk.LEFT, padx=5, pady=5)
            
            # Add a direct recalculate button that stays on the current tab
            calc_again_button = ttk.Button(
                control_panel, 
                text="Recalculate Valuation",
                command=self.calculate_valuation
            )
            calc_again_button.pack(side=tk.RIGHT, padx=5, pady=5)
            
            # Create results frame within main container
            results_frame = ttk.Frame(main_container)
            results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Left panel for DCF summary
            left_panel = ttk.Frame(results_frame)
            left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # DCF Summary
            summary_frame = ttk.LabelFrame(left_panel, text="DCF Valuation Summary", padding=10)
            summary_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            ttk.Label(summary_frame, text=f"Enterprise Value: ${ev:.2f} million").grid(row=0, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(summary_frame, text=f"- Debt: ${debt:.2f} million").grid(row=1, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(summary_frame, text=f"+ Cash: ${cash:.2f} million").grid(row=2, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(summary_frame, text=f"= Equity Value: ${equity_value:.2f} million").grid(row=3, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(summary_frame, text=f"÷ Shares Outstanding: {shares_outstanding:.2f} million").grid(row=4, column=0, sticky="w", padx=5, pady=5)
            ttk.Label(summary_frame, text=f"= Price per Share: ${price_per_share:.2f}").grid(row=5, column=0, sticky="w", padx=5, pady=5)
            
            # Right panel for visualization
            right_panel = ttk.Frame(results_frame)
            right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Create a detailed table of the DCF calculation
            table_frame = ttk.LabelFrame(right_panel, text="DCF Calculation Details", padding=10)
            table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Create treeview for DCF details
            tree = ttk.Treeview(table_frame)
            tree.pack(fill=tk.BOTH, expand=True)
            
            # Add scrollbar
            vsb = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
            vsb.pack(side='right', fill='y')
            tree.configure(yscrollcommand=vsb.set)
            
            # Define columns
            tree["columns"] = ["Year", "Revenue", "EBIT", "Tax", "NOPAT", "CapEx", "WC Change", "FCF", "DCF"]
            
            # Format columns
            tree.column("#0", width=0, stretch=tk.NO)
            for col in tree["columns"]:
                tree.column(col, anchor=tk.CENTER, width=100)
                tree.heading(col, text=col, anchor=tk.CENTER)
            
            # Add DCF data to treeview
            for i in range(self.forecast_years):
                tree.insert("", i, text="", values=(
                    f"Year {i+1}",
                    f"${revenue[i]:.2f}",
                    f"${ebit[i]:.2f}",
                    f"${tax[i]:.2f}",
                    f"${nopat[i]:.2f}",
                    f"${capex[i]:.2f}",
                    f"${wc_change[i]:.2f}",
                    f"${fcf[i]:.2f}",
                    f"${dcf[i]:.2f}"
                ))
            
            # Add terminal value
            tree.insert("", self.forecast_years, text="", values=(
                "Terminal Value",
                "-",
                "-",
                "-",
                "-",
                "-",
                "-",
                f"${terminal_value:.2f}",
                f"${discounted_tv:.2f}"
            ))
            
            # Add visualization
            fig_frame = ttk.LabelFrame(left_panel, text="Cash Flow Visualization", padding=10)
            fig_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            fig, ax = plt.subplots(figsize=(8, 4))
            years_labels = [f"Year {i+1}" for i in range(self.forecast_years)]
            
            # Bar plot of FCF
            ax.bar(years_labels, fcf, color='skyblue', label='FCF')
            
            # Add TV as a separate bar
            ax.bar("Terminal Value", terminal_value, color='orange', label='Terminal Value')
            
            ax.set_ylabel('Value (millions)')
            ax.set_title('Forecasted Free Cash Flows')
            ax.legend()
            
            # Embed plot in tkinter
            canvas = FigureCanvasTkAgg(fig, master=fig_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Switch to DCF tab
            self.notebook.select(2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate valuation: {str(e)}")
            import traceback
            traceback.print_exc()

def main():
    root = tk.Tk()
    app = DCFValuationCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
