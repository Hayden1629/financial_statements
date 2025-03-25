import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, StringVar
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re
import openpyxl

class DCFValuationCalculator:
    def __init__(self, root):
        self.root = root
        self.root.title("DCF Valuation Calculator")
        self.root.geometry("1200x900")
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
        
        # Add Calculate Valuation button to the file frame for better visibility
        self.calculate_button = ttk.Button(file_frame, text="Calculate Valuation", command=self.calculate_valuation)
        self.calculate_button.grid(row=0, column=2, padx=20, pady=5)
        
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
        
        # Button to update parameters with selected date range - make it more prominent
        refresh_button = ttk.Button(date_range_frame, text="Refresh Parameters", 
                                  command=lambda: self.recalculate_stats(),
                                  style="Accent.TButton")  # Use an accent style for visibility
        refresh_button.grid(row=0, column=2, padx=5, pady=5)
        
        # Create an accent button style
        style = ttk.Style()
        if 'Accent.TButton' not in style.theme_names():  # Check if style already exists
            style.configure('Accent.TButton', background='#0078D7', font=('Arial', 10, 'bold'))
        
        # Add info box for revenue growth year-over-year
        self.revenue_growth_info = tk.Text(left_frame, height=4, width=40, wrap=tk.WORD)
        self.revenue_growth_info.pack(fill=tk.X, padx=5, pady=5)
        self.revenue_growth_info.insert(tk.END, "Historical Revenue Growth Y/Y will display here after loading data")
        self.revenue_growth_info.config(state=tk.DISABLED)
        
        # Growth assumptions
        growth_frame = ttk.LabelFrame(left_frame, text="Growth & Margin Assumptions", padding=10)
        growth_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Create labels dictionary to store references for auto-calculated indicators
        self.auto_calc_labels = {}
        
        # Base Revenue
        ttk.Label(growth_frame, text="Base Revenue (millions):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.base_revenue_var = StringVar()
        self.base_revenue_entry = ttk.Entry(growth_frame, width=10, textvariable=self.base_revenue_var)
        self.base_revenue_entry.grid(row=0, column=1, padx=5, pady=5)
        self.base_revenue_entry.insert(0, "")
        self.auto_calc_labels["base_revenue"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["base_revenue"].grid(row=0, column=2, sticky="w", padx=5, pady=5)
        
        # Revenue Growth
        ttk.Label(growth_frame, text="Revenue Growth Rate (YoY %):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.revenue_growth_var = StringVar()
        self.revenue_growth = ttk.Entry(growth_frame, width=10, textvariable=self.revenue_growth_var)
        self.revenue_growth.grid(row=1, column=1, padx=5, pady=5)
        self.revenue_growth.insert(0, "5.0")
        self.auto_calc_labels["revenue_growth"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["revenue_growth"].grid(row=1, column=2, sticky="w", padx=5, pady=5)
        
        # Operating Margin (% of Revenue)
        ttk.Label(growth_frame, text="Operating Margin (% of Revenue):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.operating_margin = ttk.Entry(growth_frame)
        self.operating_margin.grid(row=2, column=1, padx=5, pady=5)
        self.operating_margin.insert(0, "20.0")
        self.auto_calc_labels["operating_margin"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["operating_margin"].grid(row=2, column=2, sticky="w", padx=5, pady=5)
        
        # Tax Rate
        ttk.Label(growth_frame, text="Tax Rate (%):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.tax_rate = ttk.Entry(growth_frame)
        self.tax_rate.grid(row=3, column=1, padx=5, pady=5)
        self.tax_rate.insert(0, "25.0")
        self.auto_calc_labels["tax_rate"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["tax_rate"].grid(row=3, column=2, sticky="w", padx=5, pady=5)
        
        # CapEx % of Revenue
        ttk.Label(growth_frame, text="CapEx (% of Revenue):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.capex_percent = ttk.Entry(growth_frame)
        self.capex_percent.grid(row=4, column=1, padx=5, pady=5)
        self.capex_percent.insert(0, "3.0")
        self.auto_calc_labels["capex_percent"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["capex_percent"].grid(row=4, column=2, sticky="w", padx=5, pady=5)
        
        # Working Capital % of Revenue
        ttk.Label(growth_frame, text="Working Capital (% of Revenue):").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.wc_percent = ttk.Entry(growth_frame)
        self.wc_percent.grid(row=5, column=1, padx=5, pady=5)
        self.wc_percent.insert(0, "5.0")
        self.auto_calc_labels["wc_percent"] = ttk.Label(growth_frame, text="", foreground="green")
        self.auto_calc_labels["wc_percent"].grid(row=5, column=2, sticky="w", padx=5, pady=5)
        
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
        # Leave shares outstanding blank instead of setting a default value
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
        
        # Add Reverse DCF section
        reverse_dcf_frame = ttk.LabelFrame(left_frame, text="Reverse DCF Calculator", padding=10)
        reverse_dcf_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Current market price
        ttk.Label(reverse_dcf_frame, text="Current Share Price ($):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.current_share_price = ttk.Entry(reverse_dcf_frame)
        self.current_share_price.grid(row=0, column=1, padx=5, pady=5)
        
        # Button to calculate implied discount rate
        reverse_calc_button = ttk.Button(
            reverse_dcf_frame, 
            text="Calculate Implied Discount Rate", 
            command=self.calculate_implied_discount_rate
        )
        reverse_calc_button.grid(row=1, column=0, columnspan=2, padx=5, pady=10, sticky="ew")
        
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
        
        print(f"Selected date range: {self.selected_quarters[0]} to {self.selected_quarters[-1]}")
        
        # Update historical stats display
        self.calculate_historical_stats(self.selected_quarters)
        
        # Reset all parameter fields to ensure they're recalculated
        self.reset_parameter_fields()
        
        # Prefill parameters based on the selected range
        self.prefill_forecast_parameters(self.selected_quarters)
    
    def reset_parameter_fields(self):
        """Reset all parameter fields to ensure they're recalculated from scratch"""
        print("Resetting all parameter fields for recalculation")
        
        # Reset growth parameters
        self.revenue_growth_var.set("")
        self.operating_margin.delete(0, tk.END)
        self.tax_rate.delete(0, tk.END)
        self.capex_percent.delete(0, tk.END)
        self.wc_percent.delete(0, tk.END)
        
        # Reset all auto-calculated indicators
        for label in self.auto_calc_labels.values():
            label.config(text="")
    
    def prefill_forecast_parameters(self, selected_cols=None):
        """Prefill forecast parameters from historical data"""
        if self.df is not None:
            try:
                # Check if we have quarter columns identified
                if not hasattr(self, 'quarter_cols') or not self.quarter_cols:
                    return
                    
                # Use provided columns or get all quarters
                if selected_cols is None or len(selected_cols) == 0:
                    selected_cols = self.quarter_cols
                    
                # Show what range is being used
                quarters_used = f"{selected_cols[0]} to {selected_cols[-1]}" if selected_cols else "No data"
                print(f"Using data range: {quarters_used} ({len(selected_cols)} quarters)")
                
                # Calculate base revenue (annual) from quarterly data
                if 'Revenue' in self.df.index:
                    # Calculate average quarterly revenue from selected quarters
                    quarterly_revenue = []
                    for col in selected_cols:
                        try:
                            value = self.df.loc['Revenue', col]
                            if pd.notna(value) and value > 0:
                                quarterly_revenue.append(value)
                        except Exception as e:
                            print(f"Error getting revenue for {col}: {str(e)}")
                            continue
                    
                    if quarterly_revenue:
                        # Get average quarterly revenue and annualize
                        avg_quarterly_revenue = sum(quarterly_revenue) / len(quarterly_revenue)
                        annual_revenue = avg_quarterly_revenue * 4
                        
                        # Update the base revenue field
                        self.base_revenue_entry.delete(0, tk.END)
                        self.base_revenue_entry.insert(0, f"{annual_revenue:.2f}")
                        self.auto_calc_labels["base_revenue"].config(
                            text=f"(Calculated from quarterly data, annualized)"
                        )
                        print(f"Set base revenue to {annual_revenue:.2f}M (annualized from {len(quarterly_revenue)} quarters)")
                
                # Look for "Revenue Y/Y Growth" row in the data
                growth_rows = ['Revenue Y/Y Growth', 'Revenue Growth Y/Y', 'Revenue Growth YoY', 'Revenue YoY Growth']
                growth_values = []
                growth_row_found = False
                
                print("Searching for Revenue Growth rows")
                for growth_row in growth_rows:
                    if growth_row in self.df.index:
                        growth_row_found = True
                        print(f"Found '{growth_row}' in the spreadsheet, using these values")
                        for col in selected_cols:
                            try:
                                growth_value = self.df.loc[growth_row, col]
                                # Check if it's a percentage (could be formatted as decimal or percentage)
                                if pd.notna(growth_value):
                                    # If it's likely a decimal (e.g. 0.05 for 5%), convert to percentage
                                    if -1 < growth_value < 1:
                                        growth_value = growth_value * 100
                                    growth_values.append(growth_value)
                                    print(f"Found growth value for {col}: {growth_value:.2f}%")
                            except Exception as e:
                                print(f"Error getting growth for {col}: {str(e)}")
                                continue
                        break
                
                if growth_values:
                    # Use the average of existing growth values 
                    avg_growth = sum(growth_values) / len(growth_values)
                    self.revenue_growth_var.set(f"{avg_growth:.2f}")
                    
                    # Format growth values for display
                    growth_text = f"(Avg from {len(growth_values)} values in selected range)"
                    self.auto_calc_labels["revenue_growth"].config(text=growth_text)
                    print(f"Updated revenue growth to {avg_growth:.2f}% from {len(growth_values)} existing growth values")
                    
                elif not growth_row_found:
                    # Fall back to calculated growth if no growth row found
                    print("No 'Revenue Y/Y Growth' row found in the data, calculating from revenue values")
                    self.calculate_revenue_growth(selected_cols)
                else:
                    print("Growth row found but no valid values in selected range")
                    self.calculate_revenue_growth(selected_cols)
                
                # Operating Margin - Use all selected columns, not just the last 4
                if 'Operating Income' in self.df.index and 'Revenue' in self.df.index and len(selected_cols) >= 1:
                    # Calculate average operating margin across all selected quarters
                    margins = []
                    for col in selected_cols:  # Use all selected columns
                        try:
                            op_income = self.df.loc['Operating Income', col]
                            revenue = self.df.loc['Revenue', col]
                            if pd.notna(op_income) and pd.notna(revenue) and revenue > 0:
                                margin = (op_income / revenue) * 100
                                margins.append(margin)
                        except Exception as e:
                            print(f"Error calculating margin for {col}: {str(e)}")
                            continue
                    
                    if margins:
                        avg_margin = sum(margins) / len(margins)
                        self.operating_margin.delete(0, tk.END)
                        self.operating_margin.insert(0, f"{avg_margin:.2f}")
                        self.auto_calc_labels["operating_margin"].config(
                            text=f"(Avg from {len(margins)} quarters in range {quarters_used})"
                        )
                
                # Tax Rate - Use all selected columns, not just the last 4
                if 'Income Taxes' in self.df.index and 'Pretax Income' in self.df.index and len(selected_cols) >= 1:
                    # Calculate average tax rate across all selected quarters
                    rates = []
                    for col in selected_cols:  # Use all selected columns
                        try:
                            tax = self.df.loc['Income Taxes', col]
                            pretax = self.df.loc['Pretax Income', col]
                            if pd.notna(tax) and pd.notna(pretax) and pretax > 0:
                                rate = (tax / pretax) * 100
                                rates.append(rate)
                        except Exception as e:
                            print(f"Error calculating tax rate for {col}: {str(e)}")
                            continue
                    
                    if rates:
                        avg_tax_rate = sum(rates) / len(rates)
                        self.tax_rate.delete(0, tk.END)
                        self.tax_rate.insert(0, f"{avg_tax_rate:.2f}")
                        self.auto_calc_labels["tax_rate"].config(
                            text=f"(Avg from {len(rates)} quarters in range {quarters_used})"
                        )
                
                # First, populate the latest_year_data dictionary with yearly data
                try:
                    self.update_latest_financial_data(selected_cols)
                    print("Successfully updated latest financial data from yearly columns")
                except Exception as e:
                    print(f"Error updating latest financial data: {str(e)}")
                
                # Then update CapEx and Working Capital calculations
                try:
                    self.calculate_capex_wc_from_selected_quarters(selected_cols)
                    print("Successfully calculated CapEx and Working Capital values")
                except Exception as e:
                    print(f"Error calculating CapEx and Working Capital: {str(e)}")
                
                # Set the base revenue from the most recent yearly value if available
                if self.latest_year_data.get('Revenue'):
                    base_revenue = self.latest_year_data['Revenue']
                    self.base_revenue_entry.delete(0, tk.END)
                    self.base_revenue_entry.insert(0, f"{base_revenue:.1f}")
                    self.auto_calc_labels["base_revenue"].config(
                        text=f"(From latest yearly data)"
                    )
                    print(f"Set base revenue to {base_revenue:.1f}M from yearly data")
                else:
                    print("No yearly revenue data found to prefill base revenue")
            
            except Exception as e:
                messagebox.showwarning("Warning", f"Error pre-filling parameters: {str(e)}")
                import traceback
                traceback.print_exc()
    
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
                # Group revenue data by year - ALWAYS use ALL available quarters for annual revenue data
                revenue_data = {}
                all_quarters = self.quarter_cols  # Use all quarters regardless of selection
                
                for col in all_quarters:
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
                        
                        # Always update the revenue growth info box with ALL years' data
                        self.update_revenue_growth_info(years, annual_revenue)
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
    
    def update_revenue_growth_info(self, years, annual_revenue):
        """Update the revenue growth info box with year-over-year growth rates"""
        if not hasattr(self, 'revenue_growth_info'):
            return
            
        # Enable text widget for editing
        self.revenue_growth_info.config(state=tk.NORMAL)
        self.revenue_growth_info.delete(1.0, tk.END)
        
        self.revenue_growth_info.insert(tk.END, "Historical Revenue Growth Y/Y (All Years):\n")
        
        # Display each year-over-year growth rate
        for i in range(1, len(years)):
            prev_year = years[i-1]
            curr_year = years[i]
            growth = (annual_revenue[curr_year] / annual_revenue[prev_year] - 1) * 100
            self.revenue_growth_info.insert(tk.END, f"{prev_year}-{curr_year}: {growth:.2f}%\n")
        
        # Disable text widget
        self.revenue_growth_info.config(state=tk.DISABLED)
    
    def calculate_revenue_growth(self, selected_cols):
        """Calculate revenue growth from revenue values when Revenue Y/Y Growth is not available"""
        print("Calculating revenue growth from raw revenue values...")
        
        # Try to get direct quarterly growth first
        if 'Revenue' in self.df.index and len(selected_cols) >= 2:  # Need at least 2 quarters for QoQ
            # Get revenue data for the selected quarters
            quarterly_growth_rates = []
            
            # Calculate quarter-over-quarter growth
            for i in range(1, len(selected_cols)):
                try:
                    current_revenue = self.df.loc['Revenue', selected_cols[i]]
                    prev_revenue = self.df.loc['Revenue', selected_cols[i-1]]
                    
                    if pd.notna(current_revenue) and pd.notna(prev_revenue) and prev_revenue > 0:
                        # Get the quarter numbers to check if we're comparing same quarters from different years
                        current_quarter_match = re.search(r'Q(\d)', selected_cols[i])
                        prev_quarter_match = re.search(r'Q(\d)', selected_cols[i-1])
                        
                        if current_quarter_match and prev_quarter_match:
                            current_quarter = int(current_quarter_match.group(1))
                            prev_quarter = int(prev_quarter_match.group(1))
                            
                            current_year_match = re.search(r'20(\d\d)', selected_cols[i])
                            prev_year_match = re.search(r'20(\d\d)', selected_cols[i-1])
                            
                            if current_year_match and prev_year_match:
                                current_year = int("20" + current_year_match.group(1))
                                prev_year = int("20" + prev_year_match.group(1))
                                
                                # If same quarter of different years (e.g., Q1 2022 vs Q1 2023)
                                if current_quarter == prev_quarter and current_year > prev_year:
                                    growth_rate = ((current_revenue / prev_revenue) - 1) * 100
                                    quarterly_growth_rates.append((growth_rate, f"Q{current_quarter} {prev_year}-{current_year}"))
                                    print(f"Found YoY growth for {selected_cols[i]} vs {selected_cols[i-1]}: {growth_rate:.2f}%")
                    
                except Exception as e:
                    print(f"Error calculating YoY growth for {selected_cols[i]}: {str(e)}")
            
            # If we have at least one quarter-over-quarter growth rate from the same quarter in different years
            if quarterly_growth_rates:
                # Calculate average growth rate
                avg_growth = sum(rate for rate, _ in quarterly_growth_rates) / len(quarterly_growth_rates)
                self.revenue_growth_var.set(f"{avg_growth:.2f}")
                
                # Format for display
                growth_text = ", ".join([f"{period}: {rate:.2f}%" for rate, period in quarterly_growth_rates])
                self.auto_calc_labels["revenue_growth"].config(
                    text=f"(Avg from quarterly YoY: {avg_growth:.2f}%)"
                )
                print(f"Set revenue growth to {avg_growth:.2f}% from quarterly YoY comparisons")
                return
        
        # If quarterly comparison didn't work, fall back to annual comparison
        if 'Revenue' in self.df.index:
            # Get revenue data for ALL quarters (not just selected) to build complete years
            revenue_data = {}
            for col in self.quarter_cols:  # Use all available quarters to build complete years
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
                        growth_rates.append((growth_rate, f"{prev_year}-{curr_year}"))
                
                if growth_rates:
                    # Use the most recent growth rate as default
                    most_recent_growth = growth_rates[-1][0]
                    self.revenue_growth_var.set(f"{most_recent_growth:.2f}")
                    
                    # Average of all growth rates as fallback
                    if len(growth_rates) > 1:
                        avg_growth = sum(rate for rate, _ in growth_rates) / len(growth_rates)
                    else:
                        avg_growth = most_recent_growth
                    
                    # Format for display
                    growth_text = ", ".join([f"{period}: {rate:.2f}%" for rate, period in growth_rates])
                    self.auto_calc_labels["revenue_growth"].config(
                        text=f"(Using most recent: {most_recent_growth:.2f}%)"
                    )
                    print(f"Set revenue growth to {most_recent_growth:.2f}% (most recent annual)")
                    return
            
            # If we still don't have anything, use a default growth rate
            print("Warning: No complete years for annual growth calculation, using default growth rate")
            self.revenue_growth_var.set("5.00")
            self.auto_calc_labels["revenue_growth"].config(
                text="(Default value - no historical data available)"
            )
    
    def calculate_capex_wc_from_selected_quarters(self, selected_cols):
        """Calculate CapEx and Working Capital percentages from selected quarters"""
        if not selected_cols or len(selected_cols) < 1:
            return
        
        # Extract years from the selected date range
        selected_years = set()
        for col in selected_cols:
            if isinstance(col, str) and re.match(r'Q\d 20\d\d', col):
                year_match = re.search(r'20(\d\d)', col)
                if year_match:
                    year = int("20" + year_match.group(1))
                    selected_years.add(year)
        
        print(f"Extracted years from selected date range: {sorted(selected_years)}")
        
        # Identify all yearly columns
        all_yearly_cols = []
        for col in self.df.columns:
            if isinstance(col, str):
                # Try different yearly formats: "FY20XX", "FY XX", "FY XX-XX", etc.
                if col.startswith('FY'):
                    all_yearly_cols.append(col)
                    year_str = col.replace('FY', '').strip()
                    # Try to extract a year if available
                    if re.match(r'20\d\d', year_str):
                        print(f"Found yearly column: {col} (format: FY20XX)")
                    elif re.match(r'\d\d', year_str):
                        print(f"Found yearly column: {col} (format: FY XX)")
                    else:
                        print(f"Found yearly column: {col} (other format)")
        
        print(f"All yearly columns found: {all_yearly_cols}")
        
        # Filter yearly columns to match selected years
        yearly_cols = []
        for col in all_yearly_cols:
            # Try to extract a year from the column name
            if isinstance(col, str):
                # Try different formats
                fy_year = None
                
                # Format: "FY20XX"
                fy_year_match = re.search(r'FY\s*20(\d\d)', col)
                if fy_year_match:
                    fy_year = int("20" + fy_year_match.group(1))
                
                # Format: "FY XX"
                if not fy_year:
                    fy_year_match = re.search(r'FY\s*(\d\d)', col)
                    if fy_year_match:
                        fy_year = int("20" + fy_year_match.group(1))
                
                # If we found a year and it's in our selected years, add it
                if fy_year and fy_year in selected_years:
                    yearly_cols.append(col)
                    print(f"Including yearly column {col} (year: {fy_year})")
        
        # If no columns matched our specific years, try to use the most recent years
        if not yearly_cols and selected_years and all_yearly_cols:
            # If we're looking at a specific year, prioritize that year
            target_year = max(selected_years)
            print(f"No exact year matches, looking for FY {target_year}")
            
            # Try to find exact matches first
            for col in all_yearly_cols:
                if f"{target_year}" in col:
                    yearly_cols.append(col)
                    print(f"Adding column {col} as it contains the target year {target_year}")
            
            # If still no matches, use the most recent years (up to 3)
            if not yearly_cols:
                print(f"No columns found for target year {target_year}, using most recent years")
                # Sort yearly columns by year if possible
                year_cols = []
                for col in all_yearly_cols:
                    year_match = re.search(r'20(\d\d)', col)
                    if year_match:
                        year = int("20" + year_match.group(1))
                        year_cols.append((year, col))
                if year_cols:
                    year_cols.sort(reverse=True)  # Sort by year descending
                    for i, (year, col) in enumerate(year_cols[:3]):  # Take up to 3 most recent
                        yearly_cols.append(col)
                        print(f"Using recent yearly column: {col} (year: {year})")
                else:
                    # If we can't extract years, just use the last 3 columns
                    yearly_cols = all_yearly_cols[-3:] if len(all_yearly_cols) > 3 else all_yearly_cols
                    print(f"Using last {len(yearly_cols)} yearly columns: {yearly_cols}")
        
        if not yearly_cols:
            print("No yearly columns found for Working Capital calculation in the selected date range")
            # Try to use the quarterly data if no yearly data is available
            print("Attempting to use quarterly data for Working Capital calculation")
            self.calculate_wc_from_quarterly_data(selected_cols)
        else:
            print(f"Found {len(yearly_cols)} yearly columns for WC calculation: {yearly_cols}")
            self.calculate_wc_from_yearly_data(yearly_cols)
        
        # CapEx calculation from all selected quarters
        self.calculate_capex(selected_cols)
    
    def calculate_capex(self, selected_cols):
        """Calculate CapEx from selected quarters"""
        quarters_used = f"{selected_cols[0]} to {selected_cols[-1]}" if selected_cols else "No data"
        print(f"Calculating CapEx from quarters: {quarters_used}")
        
        capex_ratios = []
        capex_keys = ['Purchase of PP&E', 'CapEx', 'Capital Expenditure', 'Capital Expenditures',
                     'Purchase of Investment', 'Acquisitions']
        
        # Try each possible CapEx key
        for capex_field in capex_keys:
            if capex_field in self.df.index:
                for col in selected_cols:
                    try:
                        capex = self.df.loc[capex_field, col]
                        revenue = self.df.loc['Revenue', col]
                        if pd.notna(capex) and pd.notna(revenue) and revenue > 0:
                            # Use absolute value since CapEx is often negative in cash flow statements
                            capex_ratio = (abs(capex) / revenue) * 100
                            capex_ratios.append(capex_ratio)
                            print(f"Found CapEx for {col}: {capex} / {revenue} = {capex_ratio:.2f}%")
                    except Exception as e:
                        print(f"Error calculating CapEx ratio for {col} with field {capex_field}: {str(e)}")
                        continue
                
                # If we found values with this field, no need to check other fields
                if capex_ratios:
                    break
                
        if capex_ratios:
            avg_capex_ratio = sum(capex_ratios) / len(capex_ratios)
            self.capex_percent.delete(0, tk.END)
            self.capex_percent.insert(0, f"{avg_capex_ratio:.2f}")
            self.auto_calc_labels["capex_percent"].config(
                text=f"(Avg from {len(capex_ratios)} quarters in range {quarters_used})"
            )
            print(f"Updated CapEx to {avg_capex_ratio:.2f}% from {len(capex_ratios)} quarters")
        else:
            print("No CapEx data found in the selected range")
    
    def calculate_wc_from_yearly_data(self, yearly_cols):
        """Calculate Working Capital from yearly data"""
        print(f"Calculating Working Capital from {len(yearly_cols)} yearly columns")
        wc_ratios = []
        
        # Try different combinations of asset/liability fields
        asset_fields = ['Current Assets', 'Total Current Assets']
        liability_fields = ['Current Liabilities', 'Total Current Liabilities']
        
        for asset_field in asset_fields:
            for liability_field in liability_fields:
                if asset_field in self.df.index and liability_field in self.df.index and 'Revenue' in self.df.index:
                    for col in yearly_cols:  # Only use yearly columns from the filtered list
                        try:
                            current_assets = self.df.loc[asset_field, col]
                            current_liabilities = self.df.loc[liability_field, col]
                            revenue = self.df.loc['Revenue', col]
                            
                            if (pd.notna(current_assets) and pd.notna(current_liabilities) and 
                                pd.notna(revenue) and revenue > 0):
                                wc = current_assets - current_liabilities
                                wc_ratio = (wc / revenue) * 100
                                wc_ratios.append(wc_ratio)
                                print(f"Found WC for {col}: ({current_assets} - {current_liabilities}) / {revenue} = {wc_ratio:.2f}%")
                        except Exception as e:
                            print(f"Error calculating WC ratio for {col}: {str(e)}")
                            continue
                    
                    # If we found values with this combination, no need to check others
                    if wc_ratios:
                        break
                
                # Break out of outer loop too if we found values
                if wc_ratios:
                    break
                
        if wc_ratios:
            avg_wc_ratio = sum(wc_ratios) / len(wc_ratios)
            self.wc_percent.delete(0, tk.END)
            self.wc_percent.insert(0, f"{avg_wc_ratio:.2f}")
            self.auto_calc_labels["wc_percent"].config(
                text=f"(Avg from {len(wc_ratios)} yearly periods)"
            )
            print(f"Updated Working Capital to {avg_wc_ratio:.2f}% from {len(wc_ratios)} yearly periods")
        else:
            print("Could not calculate Working Capital percentage - no valid yearly data found")
            
    def calculate_wc_from_quarterly_data(self, selected_cols):
        """Fall back to calculating Working Capital from quarterly data if no yearly data is available"""
        print("Calculating Working Capital from quarterly data as fallback")
        wc_ratios = []
        
        # Try different combinations of asset/liability fields
        asset_fields = ['Current Assets', 'Total Current Assets']
        liability_fields = ['Current Liabilities', 'Total Current Liabilities']
        
        for asset_field in asset_fields:
            for liability_field in liability_fields:
                if asset_field in self.df.index and liability_field in self.df.index and 'Revenue' in self.df.index:
                    for col in selected_cols:
                        try:
                            current_assets = self.df.loc[asset_field, col]
                            current_liabilities = self.df.loc[liability_field, col]
                            revenue = self.df.loc['Revenue', col]
                            
                            if (pd.notna(current_assets) and pd.notna(current_liabilities) and 
                                pd.notna(revenue) and revenue > 0):
                                wc = current_assets - current_liabilities
                                wc_ratio = (wc / revenue) * 100
                                wc_ratios.append(wc_ratio)
                                print(f"Found WC for {col}: ({current_assets} - {current_liabilities}) / {revenue} = {wc_ratio:.2f}%")
                        except Exception as e:
                            print(f"Error calculating WC ratio for {col}: {str(e)}")
                            continue
                    
                    # If we found values with this combination, no need to check others
                    if wc_ratios:
                        break
                
                # Break out of outer loop too if we found values
                if wc_ratios:
                    break
                
        if wc_ratios:
            avg_wc_ratio = sum(wc_ratios) / len(wc_ratios)
            self.wc_percent.delete(0, tk.END)
            self.wc_percent.insert(0, f"{avg_wc_ratio:.2f}")
            self.auto_calc_labels["wc_percent"].config(
                text=f"(Avg from {len(wc_ratios)} quarterly periods - no yearly data available)"
            )
            print(f"Updated Working Capital to {avg_wc_ratio:.2f}% from {len(wc_ratios)} quarterly periods (fallback)")
        else:
            print("Could not calculate Working Capital percentage - no valid data found")
    
    def update_latest_financial_data(self, selected_cols):
        """Update the latest year data dictionary with values from the dataframe"""
        self.latest_year_data = {}
        
        # Look for yearly columns (FY columns)
        yearly_cols = [col for col in self.df.columns if col.startswith('FY')]
        if yearly_cols:
            # Sort yearly columns to find most recent
            yearly_cols.sort(reverse=True)
            most_recent_yearly = yearly_cols[0]
            print(f"Using most recent yearly column: {most_recent_yearly}")
            
            # Define helper function to find most recent value - MOVED UP BEFORE USAGE
            def find_most_recent_value(key_list):
                for key in key_list:
                    if key in self.df.index:
                        if most_recent_yearly in self.df.columns:
                            value = self.df.loc[key, most_recent_yearly]
                            if pd.notna(value):
                                return value
                return None
            
            # Extract key financial metrics from most recent yearly data
            key_metrics = {
                'Revenue': ['Revenue', 'Total Revenue'],
                'Operating Income': ['Operating Income', 'Operating Profit'],
                'Income Taxes': ['Income Taxes', 'Tax Expense', 'Income Tax Expense'],
                'Pretax Income': ['Pretax Income', 'Income Before Tax', 'EBT'],
                'Current Assets': ['Current Assets', 'Total Current Assets'],
                'Current Liabilities': ['Current Liabilities', 'Total Current Liabilities'],
                'Cash & Equivalents': ['Cash & Equivalents', 'Cash and Cash Equivalents', 'Cash'],
                'Long Term Debt': ['Long Term Debt', 'Long-Term Debt'],
                'Short Term Debt': ['Short Term Debt', 'Short-Term Debt'],
                'Purchase of PP&E': ['Purchase of PP&E', 'CapEx', 'Capital Expenditure', 'Capital Expenditures', 'Purchase of Investment', 'Acquisitions']
            }
            
            for metric, keys in key_metrics.items():
                value = find_most_recent_value(keys)
                if value is not None:
                    print(f"Found {metric}: {value}")
                    self.latest_year_data[metric] = value
        
        else:
            print("No yearly (FY) columns found in the data")
        
        # Update shares outstanding - this separate process is kept to handle older code paths
        share_fields = ['Shares Outstanding', 'shares outstanding']
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
                            print(f"Updated shares outstanding to {shares:.2f} from {field}, {col}")
                            break
                    except Exception as e:
                        print(f"Error getting shares from {field}, {col}: {e}")
                        continue
                break  # Stop after we find the first valid field
        
        # Debt
        if 'Long Term Debt' in self.latest_year_data:
            debt = self.latest_year_data['Long Term Debt']
            self.current_debt.delete(0, tk.END)
            self.current_debt.insert(0, f"{debt:.2f}")
            self.auto_calc_labels["current_debt"].config(text=f"(From most recent yearly value)")
            print(f"Updated debt to {debt:.2f} from most recent yearly value")
        
        # Cash and Equivalents
        if self.latest_year_data.get('Cash & Equivalents'):
            cash = self.latest_year_data['Cash & Equivalents']
            self.cash_equivalents.delete(0, tk.END)
            self.cash_equivalents.insert(0, f"{cash:.1f}")
            self.auto_calc_labels["cash_equivalents"].config(
                text=f"(From most recent yearly value)"
            )
    
    def calculate_valuation(self):
        try:
            # Switch to DCF tab first to show it's calculating
            self.notebook.select(2)
            
            # Display a status message
            for widget in self.dcf_frame.winfo_children():
                widget.destroy()
            
            status_label = ttk.Label(self.dcf_frame, text="Calculating valuation... Please wait.", font=("Arial", 12, "bold"))
            status_label.pack(pady=20)
            self.root.update()  # Force GUI update to show the status message
            
            # Validate all inputs before proceeding
            required_fields = {
                'Forecast Years': self.forecast_years_entry,
                'Revenue Growth': self.revenue_growth_var,
                'Operating Margin': self.operating_margin,
                'Tax Rate': self.tax_rate,
                'CapEx %': self.capex_percent,
                'Working Capital %': self.wc_percent,
                'Discount Rate': self.discount_rate,
                'Terminal Growth': self.terminal_growth,
                'Shares Outstanding': self.shares_outstanding,
                'Debt': self.current_debt,
                'Cash': self.cash_equivalents,
            }
            
            # Check for empty fields
            empty_fields = [name for name, field in required_fields.items() 
                            if not field.get().strip()]
            
            if empty_fields:
                messagebox.showerror("Input Error", 
                                    f"Please fill in all required fields: {', '.join(empty_fields)}")
                return
            
            # Get forecast parameters with proper validation
            try:
                self.forecast_years = int(self.forecast_years_entry.get())
                if self.forecast_years <= 0:
                    raise ValueError("Forecast years must be a positive integer")
            except ValueError:
                messagebox.showerror("Input Error", "Forecast years must be a valid positive integer")
                return
            
            try:
                revenue_growth = float(self.revenue_growth_var.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Revenue growth rate must be a valid number")
                return
            
            try:
                operating_margin = float(self.operating_margin.get()) / 100
                if not (0 <= operating_margin <= 1):
                    messagebox.showwarning("Warning", 
                        f"Operating margin is {operating_margin*100:.2f}%, which is outside normal range (0-100%)")
            except ValueError:
                messagebox.showerror("Input Error", "Operating margin must be a valid number")
                return
            
            try:
                tax_rate = float(self.tax_rate.get()) / 100
                if not (0 <= tax_rate <= 1):
                    messagebox.showwarning("Warning", 
                        f"Tax rate is {tax_rate*100:.2f}%, which is outside normal range (0-100%)")
            except ValueError:
                messagebox.showerror("Input Error", "Tax rate must be a valid number")
                return
            
            try:
                capex_percent = float(self.capex_percent.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "CapEx percentage must be a valid number")
                return
            
            try:
                wc_percent = float(self.wc_percent.get()) / 100
                if wc_percent > 0.5:  # If WC % is over 50%, show a warning
                    messagebox.showwarning("Warning", 
                        f"Working capital percentage is {wc_percent*100:.2f}%, which is unusually high. "
                        f"This could lead to negative valuations.")
            except ValueError:
                messagebox.showerror("Input Error", "Working capital percentage must be a valid number")
                return
            
            try:
                discount_rate = float(self.discount_rate.get()) / 100
                if not (0 < discount_rate < 1):
                    messagebox.showwarning("Warning", 
                        f"Discount rate is {discount_rate*100:.2f}%, which is outside typical range (1-99%)")
            except ValueError:
                messagebox.showerror("Input Error", "Discount rate must be a valid number")
                return
            
            try:
                terminal_growth = float(self.terminal_growth.get()) / 100
                if terminal_growth >= discount_rate:
                    messagebox.showerror("Input Error", 
                        "Terminal growth rate must be less than discount rate for model validity")
                    return
            except ValueError:
                messagebox.showerror("Input Error", "Terminal growth rate must be a valid number")
                return
            
            try:
                shares_outstanding = float(self.shares_outstanding.get())
                if shares_outstanding <= 0:
                    raise ValueError("Shares must be positive")
            except ValueError:
                messagebox.showerror("Input Error", "Shares outstanding must be a valid positive number")
                return
            
            try:
                debt = float(self.current_debt.get())
            except ValueError:
                messagebox.showerror("Input Error", "Debt must be a valid number")
                return
            
            try:
                cash = float(self.cash_equivalents.get())
            except ValueError:
                messagebox.showerror("Input Error", "Cash must be a valid number")
                return
            
            # Check if a base revenue is provided
            base_revenue_provided = False
            if self.base_revenue_var.get().strip():
                try:
                    base_revenue = float(self.base_revenue_var.get())
                    base_revenue_provided = True
                except ValueError:
                    messagebox.showerror("Input Error", "Base revenue must be a valid number")
                    return
            
            # If base revenue is not provided, calculate it from historical data
            if not base_revenue_provided:
                if 'Revenue' in self.latest_year_data:
                    # Try to get the last 12 quarters of revenue data
                    revenue_values = []
                    if hasattr(self, 'quarter_cols') and len(self.quarter_cols) > 0:
                        # Get the most recent quarters (up to 12)
                        quarters_to_use = self.quarter_cols[-min(12, len(self.quarter_cols)):]
                        
                        # Collect non-NaN revenue values from these quarters
                        for col in quarters_to_use:
                            try:
                                value = self.df.loc['Revenue', col]
                                if pd.notna(value) and value > 0:
                                    revenue_values.append(value)
                            except Exception as e:
                                print(f"Warning: Could not get revenue for {col}: {e}")
                        
                        if revenue_values:
                            # Calculate average quarterly revenue and annualize
                            avg_quarterly_revenue = sum(revenue_values) / len(revenue_values)
                            base_revenue = avg_quarterly_revenue * 4
                            print(f"Using average of {len(revenue_values)} quarters for base revenue calculation")
                        else:
                            # Fallback to latest revenue value if no historical data found
                            base_revenue = self.latest_year_data['Revenue'] * 4
                            print("Warning: No historical quarterly data found, using latest quarter * 4")
                    else:
                        # Fallback to latest revenue value if no quarter columns defined
                        base_revenue = self.latest_year_data['Revenue'] * 4
                        print("Warning: No quarter columns defined, using latest quarter * 4")
                else:
                    messagebox.showerror("Error", "Could not find revenue data in the financial statement. Please enter base revenue manually.")
                    return
            
            # Print inputs for debugging
            print(f"DCF Model Inputs:")
            print(f"  Base Revenue: ${base_revenue:.2f}M")
            print(f"  Forecast Years: {self.forecast_years}")
            print(f"  Revenue Growth: {revenue_growth*100:.2f}%")
            print(f"  Operating Margin: {operating_margin*100:.2f}%")
            print(f"  Tax Rate: {tax_rate*100:.2f}%")
            print(f"  CapEx %: {capex_percent*100:.2f}%")
            print(f"  Working Capital %: {wc_percent*100:.2f}%")
            print(f"  Discount Rate: {discount_rate*100:.2f}%")
            print(f"  Terminal Growth: {terminal_growth*100:.2f}%")
            print(f"  Shares Outstanding: {shares_outstanding}M")
            print(f"  Debt: ${debt}M")
            print(f"  Cash: ${cash}M")
            
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
            wc_change = [0] * len(wc)  # Initialize with zeros
            
            # Calculate working capital changes with a more stable approach
            wc_initial = base_revenue * wc_percent  # Initial working capital
            wc_change[0] = wc[0] - wc_initial  # Change in first year
            
            # Calculate changes for remaining years
            for i in range(1, len(wc)):
                wc_change[i] = wc[i] - wc[i-1]
            
            # Print working capital info for debugging
            print("\nWorking Capital Calculations (using YEARLY data):")
            print(f"  Initial WC: ${wc_initial:.2f}M")
            for i in range(len(wc)):
                print(f"  Year {i+1}: WC ${wc[i]:.2f}M, Change ${wc_change[i]:.2f}M")
            
            # Free Cash Flow - Fixed calculation
            fcf = [nopat[i] - capex[i] - wc_change[i] for i in range(len(nopat))]
            
            # Print FCF values for debugging
            print("\nFree Cash Flow Calculations:")
            for i in range(len(fcf)):
                print(f"  Year {i+1}: NOPAT ${nopat[i]:.2f}M - CapEx ${capex[i]:.2f}M - WC Change ${wc_change[i]:.2f}M = FCF ${fcf[i]:.2f}M")
            
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
            
            # Add a new FCF Projection Table
            fcf_table_frame = ttk.LabelFrame(left_panel, text="Free Cash Flow Projections", padding=10)
            fcf_table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Create a simple table for FCF projections
            fcf_table = ttk.Treeview(fcf_table_frame, height=len(years))
            fcf_table.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
            
            # Define columns: Year and FCF
            fcf_table["columns"] = ["Year", "FCF", "Present Value"]
            
            # Format columns
            fcf_table.column("#0", width=0, stretch=tk.NO)
            fcf_table.column("Year", anchor=tk.W, width=100)
            fcf_table.column("FCF", anchor=tk.E, width=150)
            fcf_table.column("Present Value", anchor=tk.E, width=150)
            
            # Create column headings
            fcf_table.heading("#0", text="", anchor=tk.W)
            fcf_table.heading("Year", text="Year", anchor=tk.W)
            fcf_table.heading("FCF", text="Free Cash Flow (millions)", anchor=tk.CENTER)
            fcf_table.heading("Present Value", text="Present Value (millions)", anchor=tk.CENTER)
            
            # Add data to the table
            for i in range(self.forecast_years):
                fcf_table.insert("", i, text="", values=(
                    f"Year {i+1}",
                    f"${fcf[i]:.2f}",
                    f"${dcf[i]:.2f}"
                ))
            
            # Add terminal value as the last row
            fcf_table.insert("", self.forecast_years, text="", values=(
                "Terminal Value",
                f"${terminal_value:.2f}",
                f"${discounted_tv:.2f}"
            ))
            
            # Add row for sum of present values
            fcf_table.insert("", self.forecast_years + 1, text="", values=(
                "Total Enterprise Value",
                "",
                f"${ev:.2f}"
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
            
            # Switch to DCF tab
            self.notebook.select(2)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate valuation: {str(e)}")
            import traceback
            traceback.print_exc()

    def calculate_implied_discount_rate(self):
        """Calculate the discount rate implied by the current share price"""
        try:
            # Validate required inputs
            required_fields = {
                'Forecast Years': self.forecast_years_entry,
                'Revenue Growth': self.revenue_growth_var,
                'Operating Margin': self.operating_margin,
                'Tax Rate': self.tax_rate,
                'CapEx %': self.capex_percent,
                'Working Capital %': self.wc_percent,
                'Terminal Growth': self.terminal_growth,
                'Shares Outstanding': self.shares_outstanding,
                'Debt': self.current_debt,
                'Cash': self.cash_equivalents,
                'Current Share Price': self.current_share_price,
            }
            
            # Check for empty fields
            empty_fields = [name for name, field in required_fields.items() 
                           if not field.get().strip()]
            
            if empty_fields:
                messagebox.showerror("Input Error", 
                                    f"Please fill in all required fields: {', '.join(empty_fields)}")
                return
            
            # Get input values
            try:
                forecast_years = int(self.forecast_years_entry.get())
                if forecast_years <= 0:
                    raise ValueError("Forecast years must be a positive integer")
            except ValueError:
                messagebox.showerror("Input Error", "Forecast years must be a valid positive integer")
                return
                
            try:
                revenue_growth = float(self.revenue_growth_var.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Revenue growth rate must be a valid number")
                return
                
            try:
                operating_margin = float(self.operating_margin.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Operating margin must be a valid number")
                return
                
            try:
                tax_rate = float(self.tax_rate.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Tax rate must be a valid number")
                return
                
            try:
                capex_percent = float(self.capex_percent.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "CapEx percentage must be a valid number")
                return
                
            try:
                wc_percent = float(self.wc_percent.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Working capital percentage must be a valid number")
                return
                
            try:
                terminal_growth = float(self.terminal_growth.get()) / 100
            except ValueError:
                messagebox.showerror("Input Error", "Terminal growth rate must be a valid number")
                return
                
            try:
                shares_outstanding = float(self.shares_outstanding.get())
                if shares_outstanding <= 0:
                    raise ValueError("Shares must be positive")
            except ValueError:
                messagebox.showerror("Input Error", "Shares outstanding must be a valid positive number")
                return
                
            try:
                debt = float(self.current_debt.get())
            except ValueError:
                messagebox.showerror("Input Error", "Debt must be a valid number")
                return
                
            try:
                cash = float(self.cash_equivalents.get())
            except ValueError:
                messagebox.showerror("Input Error", "Cash must be a valid number")
                return
                
            try:
                current_price = float(self.current_share_price.get())
                if current_price <= 0:
                    raise ValueError("Share price must be positive")
            except ValueError:
                messagebox.showerror("Input Error", "Current share price must be a valid positive number")
                return
            
            # Calculate base revenue in the same way as the forward DCF model
            if 'Revenue' in self.latest_year_data:
                # Try to get the last 12 quarters of revenue data
                revenue_values = []
                if hasattr(self, 'quarter_cols') and len(self.quarter_cols) > 0:
                    # Get the most recent quarters (up to 12)
                    quarters_to_use = self.quarter_cols[-min(12, len(self.quarter_cols)):]
                    
                    # Collect non-NaN revenue values from these quarters
                    for col in quarters_to_use:
                        try:
                            value = self.df.loc['Revenue', col]
                            if pd.notna(value) and value > 0:
                                revenue_values.append(value)
                        except Exception as e:
                            print(f"Warning: Could not get revenue for {col}: {e}")
                    
                    if revenue_values:
                        # Calculate average quarterly revenue and annualize
                        avg_quarterly_revenue = sum(revenue_values) / len(revenue_values)
                        base_revenue = avg_quarterly_revenue * 4
                        print(f"Using average of {len(revenue_values)} quarters for base revenue calculation")
                    else:
                        # Fallback to latest revenue value if no historical data found
                        base_revenue = self.latest_year_data['Revenue'] * 4
                        print("Warning: No historical quarterly data found, using latest quarter * 4")
                else:
                    # Fallback to latest revenue value if no quarter columns defined
                    base_revenue = self.latest_year_data['Revenue'] * 4
                    print("Warning: No quarter columns defined, using latest quarter * 4")
            else:
                messagebox.showerror("Error", "Could not find revenue data in the financial statement")
                return
            
            # Calculate the target equity value from the current share price
            target_equity_value = current_price * shares_outstanding
            
            # Calculate the target enterprise value
            target_ev = target_equity_value + debt - cash
            
            # Implement a binary search algorithm to find the discount rate that matches the target EV
            def calculate_ev_with_discount_rate(discount_rate):
                # Create forecast model
                years = list(range(1, forecast_years + 1))
                revenue = [base_revenue * (1 + revenue_growth) ** year for year in years]
                ebit = [rev * operating_margin for rev in revenue]
                tax_amount = [op * tax_rate for op in ebit]
                nopat = [op - tx for op, tx in zip(ebit, tax_amount)]
                
                # Calculate CapEx and Working Capital changes
                capex = [rev * capex_percent for rev in revenue]
                
                # For working capital, we need to calculate the change year over year
                wc = [rev * wc_percent for rev in revenue]
                wc_change = [0] + [wc[i] - wc[i-1] for i in range(1, len(wc))]
                wc_change[0] = wc[0] - (base_revenue * wc_percent)
                
                # Free Cash Flow
                fcf = [nopat[i] - capex[i] - wc_change[i] for i in range(len(nopat))]
                
                # Terminal Value
                terminal_value = fcf[-1] * (1 + terminal_growth) / (discount_rate - terminal_growth)
                
                # Discounted Cash Flows
                dcf_values = [flow / (1 + discount_rate) ** year for year, flow in zip(years, fcf)]
                
                # Discounted Terminal Value
                discounted_tv = terminal_value / (1 + discount_rate) ** forecast_years
                
                # Enterprise Value
                ev = sum(dcf_values) + discounted_tv
                
                return ev
            
            # Binary search to find the implied discount rate
            low_rate = 0.01  # 1%
            high_rate = 0.50  # 50%
            mid_rate = (low_rate + high_rate) / 2
            tolerance = 0.0001  # Acceptable error in enterprise value
            max_iterations = 100
            
            # Check if terminal growth is less than the lowest discount rate we'll try
            if terminal_growth >= low_rate:
                low_rate = terminal_growth + 0.01  # Set low_rate just above terminal growth
                if low_rate >= high_rate:
                    messagebox.showerror("Error", 
                        "Terminal growth rate is too high to find a valid discount rate solution")
                    return
            
            # Ensure we can bracket the solution
            ev_at_low = calculate_ev_with_discount_rate(low_rate)
            ev_at_high = calculate_ev_with_discount_rate(high_rate)
            
            if (ev_at_low < target_ev and ev_at_high < target_ev) or (ev_at_low > target_ev and ev_at_high > target_ev):
                messagebox.showerror("Error", 
                    f"Cannot find a solution in the range {low_rate*100:.1f}% to {high_rate*100:.1f}%. "
                    f"The current price may be outside the model's realistic valuation range.")
                return
            
            # Perform binary search
            iteration = 0
            while iteration < max_iterations:
                mid_rate = (low_rate + high_rate) / 2
                
                # Check if mid_rate would lead to invalid terminal value calculation
                if mid_rate <= terminal_growth:
                    low_rate = mid_rate
                    continue
                
                ev_at_mid = calculate_ev_with_discount_rate(mid_rate)
                error = abs(ev_at_mid - target_ev)
                
                # If within tolerance, we found our solution
                if error < tolerance * target_ev:
                    break
                
                # Adjust search range
                if ev_at_mid > target_ev:
                    low_rate = mid_rate
                else:
                    high_rate = mid_rate
                
                iteration += 1
            
            # Get the resulting values for displaying
            implied_discount_rate = mid_rate
            ev_at_implied_rate = calculate_ev_with_discount_rate(implied_discount_rate)
            
            # Calculate implied price per share for validation
            implied_equity_value = ev_at_implied_rate - debt + cash
            implied_price = implied_equity_value / shares_outstanding
            
            # Display results
            result_window = tk.Toplevel(self.root)
            result_window.title("Reverse DCF Analysis Results")
            result_window.geometry("500x400")
            result_window.transient(self.root)
            result_window.grab_set()
            
            # Create a frame to hold the results
            result_frame = ttk.Frame(result_window, padding=20)
            result_frame.pack(fill=tk.BOTH, expand=True)
            
            # Add results text
            ttk.Label(result_frame, text="Reverse DCF Analysis Results", 
                    font=("Arial", 14, "bold")).pack(pady=(0, 20))
            
            # Display the implied discount rate
            ttk.Label(result_frame, text=f"Implied Discount Rate: {implied_discount_rate*100:.2f}%", 
                    font=("Arial", 12)).pack(anchor="w", pady=5)
            
            # Display the target values
            ttk.Label(result_frame, text=f"Target Share Price: ${current_price:.2f}", 
                    font=("Arial", 11)).pack(anchor="w", pady=5)
            ttk.Label(result_frame, text=f"Target Equity Value: ${target_equity_value:.2f} million", 
                    font=("Arial", 11)).pack(anchor="w", pady=5)
            ttk.Label(result_frame, text=f"Target Enterprise Value: ${target_ev:.2f} million", 
                    font=("Arial", 11)).pack(anchor="w", pady=5)
            
            # Display key assumptions used
            ttk.Label(result_frame, text="Key Assumptions:", 
                    font=("Arial", 11, "bold")).pack(anchor="w", pady=(15, 5))
            ttk.Label(result_frame, text=f"Base Revenue: ${base_revenue:.2f} million").pack(anchor="w", pady=2)
            ttk.Label(result_frame, text=f"Revenue Growth: {revenue_growth*100:.2f}%").pack(anchor="w", pady=2)
            ttk.Label(result_frame, text=f"Operating Margin: {operating_margin*100:.2f}%").pack(anchor="w", pady=2)
            ttk.Label(result_frame, text=f"Terminal Growth Rate: {terminal_growth*100:.2f}%").pack(anchor="w", pady=2)
            
            # Add interpretation
            ttk.Separator(result_frame, orient="horizontal").pack(fill="x", pady=15)
            interpretation = (
                f"Based on the current share price of ${current_price:.2f}, the market is implying "
                f"a {implied_discount_rate*100:.2f}% discount rate (cost of capital).\n\n"
                f"{'This is higher than typical discount rates, suggesting the market views this investment as risky.' if implied_discount_rate > 0.12 else 'This is within typical discount rate ranges, suggesting fair market pricing.'}"
            )
            
            interpretation_label = ttk.Label(result_frame, text=interpretation, wraplength=450, justify="left")
            interpretation_label.pack(pady=10)
            
            # Add a button to use this discount rate in the model
            ttk.Button(result_frame, text="Use This Discount Rate in DCF Model", 
                     command=lambda: self.apply_implied_discount_rate(implied_discount_rate, result_window)).pack(pady=15)
            
            # Apply this discount rate to the main model
            self.discount_rate.delete(0, tk.END)
            self.discount_rate.insert(0, f"{implied_discount_rate*100:.2f}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to calculate implied discount rate: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def apply_implied_discount_rate(self, discount_rate, window):
        """Apply the calculated discount rate to the main model and close the window"""
        # Update the discount rate in the main form
        self.discount_rate.delete(0, tk.END)
        self.discount_rate.insert(0, f"{discount_rate*100:.2f}")
        
        # Close the window
        window.destroy()
        
        # Show confirmation
        messagebox.showinfo("Discount Rate Applied", 
                           f"The implied discount rate of {discount_rate*100:.2f}% has been applied to your DCF model.")

def main():
    root = tk.Tk()
    app = DCFValuationCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
