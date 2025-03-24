import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import re

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
        
        # Calculate button
        ttk.Button(main_frame, text="Calculate Valuation", command=self.calculate_valuation).pack(padx=5, pady=10)
    
    def create_forecast_inputs(self):
        # Left frame for inputs
        left_frame = ttk.Frame(self.forecast_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Growth assumptions
        growth_frame = ttk.LabelFrame(left_frame, text="Growth & Margin Assumptions", padding=10)
        growth_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # Revenue Growth
        ttk.Label(growth_frame, text="Revenue Growth Rate (%):").grid(row=0, column=0, sticky="w", padx=5, pady=5)
        self.revenue_growth = ttk.Entry(growth_frame)
        self.revenue_growth.grid(row=0, column=1, padx=5, pady=5)
        self.revenue_growth.insert(0, "5.0")
        
        # Operating Margin
        ttk.Label(growth_frame, text="Operating Margin (%):").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        self.operating_margin = ttk.Entry(growth_frame)
        self.operating_margin.grid(row=1, column=1, padx=5, pady=5)
        self.operating_margin.insert(0, "20.0")
        
        # Tax Rate
        ttk.Label(growth_frame, text="Tax Rate (%):").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        self.tax_rate = ttk.Entry(growth_frame)
        self.tax_rate.grid(row=2, column=1, padx=5, pady=5)
        self.tax_rate.insert(0, "25.0")
        
        # CapEx % of Revenue
        ttk.Label(growth_frame, text="CapEx (% of Revenue):").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        self.capex_percent = ttk.Entry(growth_frame)
        self.capex_percent.grid(row=3, column=1, padx=5, pady=5)
        self.capex_percent.insert(0, "3.0")
        
        # Working Capital % of Revenue
        ttk.Label(growth_frame, text="Working Capital (% of Revenue):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.wc_percent = ttk.Entry(growth_frame)
        self.wc_percent.grid(row=4, column=1, padx=5, pady=5)
        self.wc_percent.insert(0, "5.0")
        
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
        
        # Current Debt
        ttk.Label(dcf_frame, text="Current Debt (millions):").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        self.current_debt = ttk.Entry(dcf_frame)
        self.current_debt.grid(row=4, column=1, padx=5, pady=5)
        self.current_debt.insert(0, "0.0")
        
        # Cash & Equivalents
        ttk.Label(dcf_frame, text="Cash & Equivalents (millions):").grid(row=5, column=0, sticky="w", padx=5, pady=5)
        self.cash_equivalents = ttk.Entry(dcf_frame)
        self.cash_equivalents.grid(row=5, column=1, padx=5, pady=5)
        self.cash_equivalents.insert(0, "0.0")
        
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
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        
        if file_path:
            try:
                self.file_label.config(text=file_path)
                self.df = pd.read_csv(file_path, skipinitialspace=True)
                
                # Clean the data
                self.clean_data()
                
                # Display historical data
                self.display_historical_data()
                
                # Calculate and display historical stats
                self.calculate_historical_stats()
                
                # Pre-fill forecast parameters from historical data
                self.prefill_forecast_parameters()
                
                # Switch to forecast tab
                self.notebook.select(1)
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load the file: {str(e)}")
    
    def clean_data(self):
        # Replace empty strings with NaN
        self.df = self.df.replace('', np.nan)
        
        # Remove the first two rows (they are header information)
        self.df = self.df.iloc[2:].reset_index(drop=True)
        
        # Set 'Account' as the index
        if 'Account' in self.df.columns:
            self.df.set_index('Account', inplace=True)
        
        # Convert numeric columns to float
        for col in self.df.columns:
            try:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
            except:
                pass
    
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
                tree.heading(col, text=col, anchor=tk.W)
            
            # Add data to treeview
            for i, row in df_display.iterrows():
                tree.insert("", i, text="", values=list(row))
    
    def calculate_historical_stats(self):
        if self.df is None:
            return
        
        # Enable text widget for editing
        self.hist_stats.configure(state='normal')
        self.hist_stats.delete(1.0, tk.END)
        
        try:
            # Get the last 5 years of data
            latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-20:]
            recent_data = self.df.loc[:, latest_cols]
            
            # Store latest quarter data
            self.latest_quarter = latest_cols[-1]
            
            # Calculate average revenue growth
            if 'Revenue' in self.df.index:
                revenue_data = self.df.loc['Revenue', latest_cols].dropna()
                if len(revenue_data) > 4:  # Need at least 5 quarters to calculate 4 growth rates
                    growth_rates = []
                    for i in range(1, len(revenue_data)):
                        if revenue_data.iloc[i-1] > 0:  # Avoid division by zero
                            growth_rate = (revenue_data.iloc[i] / revenue_data.iloc[i-1] - 1) * 100
                            growth_rates.append(growth_rate)
                    
                    avg_growth = np.mean(growth_rates)
                    self.hist_stats.insert(tk.END, f"Average Quarterly Revenue Growth: {avg_growth:.2f}%\n")
                    
                    # Annualized growth rate
                    annualized_growth = ((1 + avg_growth/100) ** 4 - 1) * 100
                    self.hist_stats.insert(tk.END, f"Annualized Revenue Growth: {annualized_growth:.2f}%\n\n")
            
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
            
            # Calculate average CapEx as % of revenue
            if 'Purchase of PP&E' in self.df.index and 'Revenue' in self.df.index:
                capex = self.df.loc['Purchase of PP&E', latest_cols].dropna().abs()  # CapEx is usually negative in cashflow
                revenue = self.df.loc['Revenue', latest_cols].dropna()
                
                # Match indices
                common_cols = capex.index.intersection(revenue.index)
                capex = capex[common_cols]
                revenue = revenue[common_cols]
                
                if len(capex) > 0 and len(revenue) > 0:
                    capex_ratio = (capex / revenue) * 100
                    avg_capex_ratio = np.mean(capex_ratio)
                    self.hist_stats.insert(tk.END, f"Average CapEx (% of Revenue): {avg_capex_ratio:.2f}%\n\n")
            
            # Calculate Working Capital as % of revenue
            if all(item in self.df.index for item in ['Current Assets', 'Current Liabilities', 'Revenue']):
                current_assets = self.df.loc['Current Assets', latest_cols].dropna()
                current_liabilities = self.df.loc['Current Liabilities', latest_cols].dropna()
                revenue = self.df.loc['Revenue', latest_cols].dropna()
                
                # Match indices
                common_cols = current_assets.index.intersection(current_liabilities.index).intersection(revenue.index)
                current_assets = current_assets[common_cols]
                current_liabilities = current_liabilities[common_cols]
                revenue = revenue[common_cols]
                
                if len(current_assets) > 0 and len(current_liabilities) > 0 and len(revenue) > 0:
                    working_capital = current_assets - current_liabilities
                    wc_ratio = (working_capital / revenue) * 100
                    avg_wc_ratio = np.mean(wc_ratio)
                    self.hist_stats.insert(tk.END, f"Average Working Capital (% of Revenue): {avg_wc_ratio:.2f}%\n\n")
            
            # Store latest financial data
            self.latest_year_data = {}
            for key in ['Revenue', 'Operating Income', 'Income Taxes', 'Pretax Income', 
                       'Current Assets', 'Current Liabilities', 'Cash & Equivalents',
                       'Long Term Debt', 'Short Term Debt']:
                if key in self.df.index:
                    self.latest_year_data[key] = self.df.loc[key, latest_cols[-1]]
            
        except Exception as e:
            self.hist_stats.insert(tk.END, f"Error calculating statistics: {str(e)}")
        
        # Disable text widget
        self.hist_stats.configure(state='disabled')
    
    def prefill_forecast_parameters(self):
        # Prefill forecast parameters from historical data
        if self.df is not None:
            try:
                # Revenue Growth
                if 'Revenue' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-8:]  # Last 2 years
                    revenue_data = self.df.loc['Revenue', latest_cols].dropna()
                    if len(revenue_data) > 4:
                        # Calculate year-over-year growth
                        q1_this_year = revenue_data.iloc[-4]
                        q1_last_year = revenue_data.iloc[-8]
                        yoy_growth = ((q1_this_year / q1_last_year) - 1) * 100
                        self.revenue_growth.delete(0, tk.END)
                        self.revenue_growth.insert(0, f"{yoy_growth:.2f}")
                
                # Operating Margin
                if 'Operating Income' in self.df.index and 'Revenue' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-4:]  # Last year
                    op_income_data = self.df.loc['Operating Income', latest_cols].dropna()
                    revenue_data = self.df.loc['Revenue', latest_cols].dropna()
                    
                    common_cols = op_income_data.index.intersection(revenue_data.index)
                    op_income_data = op_income_data[common_cols]
                    revenue_data = revenue_data[common_cols]
                    
                    if len(op_income_data) > 0 and len(revenue_data) > 0:
                        avg_margin = np.mean((op_income_data / revenue_data) * 100)
                        self.operating_margin.delete(0, tk.END)
                        self.operating_margin.insert(0, f"{avg_margin:.2f}")
                
                # Tax Rate
                if 'Income Taxes' in self.df.index and 'Pretax Income' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-4:]  # Last year
                    tax_data = self.df.loc['Income Taxes', latest_cols].dropna()
                    pretax_data = self.df.loc['Pretax Income', latest_cols].dropna()
                    
                    common_cols = tax_data.index.intersection(pretax_data.index)
                    tax_data = tax_data[common_cols]
                    pretax_data = pretax_data[common_cols]
                    
                    if len(tax_data) > 0 and len(pretax_data) > 0:
                        avg_tax_rate = np.mean((tax_data / pretax_data) * 100)
                        self.tax_rate.delete(0, tk.END)
                        self.tax_rate.insert(0, f"{avg_tax_rate:.2f}")
                
                # CapEx
                if 'Purchase of PP&E' in self.df.index and 'Revenue' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-4:]  # Last year
                    capex_data = self.df.loc['Purchase of PP&E', latest_cols].dropna().abs()
                    revenue_data = self.df.loc['Revenue', latest_cols].dropna()
                    
                    common_cols = capex_data.index.intersection(revenue_data.index)
                    capex_data = capex_data[common_cols]
                    revenue_data = revenue_data[common_cols]
                    
                    if len(capex_data) > 0 and len(revenue_data) > 0:
                        avg_capex_ratio = np.mean((capex_data / revenue_data) * 100)
                        self.capex_percent.delete(0, tk.END)
                        self.capex_percent.insert(0, f"{avg_capex_ratio:.2f}")
                
                # Working Capital
                if all(item in self.df.index for item in ['Current Assets', 'Current Liabilities', 'Revenue']):
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-4:]  # Last year
                    ca_data = self.df.loc['Current Assets', latest_cols].dropna()
                    cl_data = self.df.loc['Current Liabilities', latest_cols].dropna()
                    revenue_data = self.df.loc['Revenue', latest_cols].dropna()
                    
                    common_cols = ca_data.index.intersection(cl_data.index).intersection(revenue_data.index)
                    ca_data = ca_data[common_cols]
                    cl_data = cl_data[common_cols]
                    revenue_data = revenue_data[common_cols]
                    
                    if len(ca_data) > 0 and len(cl_data) > 0 and len(revenue_data) > 0:
                        wc_data = ca_data - cl_data
                        avg_wc_ratio = np.mean((wc_data / revenue_data) * 100)
                        self.wc_percent.delete(0, tk.END)
                        self.wc_percent.insert(0, f"{avg_wc_ratio:.2f}")
                
                # Shares Outstanding (estimate from financial data)
                if 'Common Stock' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-1:]
                    shares_data = self.df.loc['Common Stock', latest_cols].dropna()
                    if len(shares_data) > 0:
                        shares = shares_data.iloc[-1]
                        self.shares_outstanding.delete(0, tk.END)
                        self.shares_outstanding.insert(0, f"{shares:.2f}")
                
                # Debt and Cash
                if 'Long Term Debt' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-1:]
                    debt_data = self.df.loc['Long Term Debt', latest_cols].dropna()
                    if len(debt_data) > 0:
                        debt = debt_data.iloc[-1]
                        self.current_debt.delete(0, tk.END)
                        self.current_debt.insert(0, f"{debt:.2f}")
                
                if 'Cash & Equivalents' in self.df.index:
                    latest_cols = sorted([col for col in self.df.columns if re.match(r'Q\d 20\d\d', col)])[-1:]
                    cash_data = self.df.loc['Cash & Equivalents', latest_cols].dropna()
                    if len(cash_data) > 0:
                        cash = cash_data.iloc[-1]
                        self.cash_equivalents.delete(0, tk.END)
                        self.cash_equivalents.insert(0, f"{cash:.2f}")
                
            except Exception as e:
                messagebox.showwarning("Warning", f"Error pre-filling parameters: {str(e)}")
    
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
            
            # Create results frame
            results_frame = ttk.Frame(self.dcf_frame)
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

def main():
    root = tk.Tk()
    app = DCFValuationCalculator(root)
    root.mainloop()

if __name__ == "__main__":
    main()
