import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from datetime import datetime

class ChartViewer:
    def __init__(self, parent, comparison_results, original_df, updated_df):
        self.window = tk.Toplevel(parent)
        self.window.title("Promotion Analysis Charts")
        self.window.geometry("1200x800")
        
        self.comparison_results = comparison_results
        self.original_df = original_df
        self.updated_df = updated_df
        self.current_chart = None
        self.fig = None
        self.ax = None
        
        self.chart_types = [
            ("Distribution of Changes", "distribution"),
            ("Top Changes", "top_changes"),
            ("Promotion Type Analysis", "promo_type"),
            ("Customer Group Analysis", "customer_group"),
            ("Timeline Analysis", "timeline"),
            ("Quantity Range Analysis", "qty_range"),
            ("Change Patterns", "patterns")
        ]
        
        self.setup_gui()
    
    def setup_gui(self):
        control_frame = ttk.Frame(self.window, padding="5")
        control_frame.pack(fill="x", padx=5, pady=5)
        
        # Chart Selection
        ttk.Label(control_frame, text="Select Analysis Type:").pack(side="left", padx=5)
        self.chart_var = tk.StringVar(value="Distribution of Changes")
        
        # Create chart selection dropdown
        chart_combo = ttk.Combobox(control_frame, 
                                 textvariable=self.chart_var,
                                 values=[chart[0] for chart in self.chart_types],
                                 state="readonly",
                                 width=30)
        chart_combo.pack(side="left", padx=5)
        chart_combo.bind('<<ComboboxSelected>>', lambda e: self.update_chart())
        
        # Save Button
        ttk.Button(control_frame, text="Save Chart",
                  command=self.save_chart).pack(side="right", padx=5)
        
        # Chart Frame
        self.chart_frame = ttk.LabelFrame(self.window, text="Visualization", padding="5")
        self.chart_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.update_chart()
    
    def clear_chart(self):
        for widget in self.chart_frame.winfo_children():
            widget.destroy()
        if self.fig is not None:
            plt.close(self.fig)
    
    def update_chart(self):
        self.clear_chart()
        
        changes_df = pd.DataFrame(self.comparison_results)
        chart_type = self.chart_var.get()
        
        try:
            if chart_type == "Distribution of Changes":
                self.create_distribution_chart(changes_df)
            elif chart_type == "Top Changes":
                self.create_top_changes_chart(changes_df)
            elif chart_type == "Promotion Type Analysis":
                self.create_promo_type_analysis(changes_df)
            elif chart_type == "Customer Group Analysis":
                self.create_customer_group_analysis(changes_df)
            elif chart_type == "Timeline Analysis":
                self.create_timeline_analysis(changes_df)
            elif chart_type == "Quantity Range Analysis":
                self.create_quantity_range_analysis(changes_df)
            elif chart_type == "Change Patterns":
                self.create_change_patterns(changes_df)
            
            self.fig.tight_layout()
            canvas = FigureCanvasTkAgg(self.fig, self.chart_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
            
        except Exception as e:
            messagebox.showerror("Error", f"Error creating chart: {str(e)}")
    
    def create_distribution_chart(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        
        # Overall distribution
        sns.histplot(data=df, x='change', bins=30, ax=ax1)
        ax1.axvline(x=0, color='r', linestyle='--')
        ax1.set_title('Distribution of Changes')
        
        # Top increases
        top_inc = df[df['change'] > 0].nlargest(5, 'change')
        ax2.barh(range(len(top_inc)), top_inc['change'], color='g')
        ax2.set_yticks(range(len(top_inc)))
        ax2.set_yticklabels(top_inc['td_desc'].str[:20])
        ax2.set_title('Top 5 Increases')
        
        # Top decreases
        top_dec = df[df['change'] < 0].nsmallest(5, 'change')
        ax3.barh(range(len(top_dec)), top_dec['change'], color='r')
        ax3.set_yticks(range(len(top_dec)))
        ax3.set_yticklabels(top_dec['td_desc'].str[:20])
        ax3.set_title('Top 5 Decreases')
        
        # Change summary
        changes = [
            len(df[df['change'] > 0]),
            len(df[df['change'] < 0]),
            len(df[df['change'] == 0])
        ]
        ax4.pie(changes, labels=['Increases', 'Decreases', 'No Change'],
                autopct='%1.1f%%', colors=['g', 'r', 'gray'])
        ax4.set_title('Change Distribution')
    
    def create_top_changes_chart(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        
        # Get number of top changes from user
        try:
            n_changes = simpledialog.askinteger("Input", 
                                              "Enter number of top changes to show (e.g., 5, 10):",
                                              minvalue=1, maxvalue=50, initialvalue=5)
            if n_changes is None:
                n_changes = 5  # Default value if user cancels
        except:
            n_changes = 5  # Default value if there's an error
        
        # Absolute changes (fixing nlargest error)
        abs_changes = df.copy()
        abs_changes['abs_change'] = abs(abs_changes['change'])
        abs_changes = abs_changes.nlargest(n_changes, 'abs_change')
        colors = ['g' if x > 0 else 'r' for x in abs_changes['change']]
        
        ax1.barh(range(len(abs_changes)), abs_changes['change'], color=colors)
        ax1.set_yticks(range(len(abs_changes)))
        ax1.set_yticklabels(abs_changes['td_desc'].str[:20])
        ax1.set_title(f'Top {n_changes} Absolute Changes')
        
        # Percentage changes
        df['pct_change'] = (df['change'] / df['orig_qty'] * 100).replace([np.inf, -np.inf], 0)
        pct_changes = df.copy()
        pct_changes['abs_pct_change'] = abs(pct_changes['pct_change'])
        top_pct = pct_changes.nlargest(n_changes, 'abs_pct_change')
        colors = ['g' if x > 0 else 'r' for x in top_pct['pct_change']]
        
        ax2.barh(range(len(top_pct)), top_pct['pct_change'], color=colors)
        ax2.set_yticks(range(len(top_pct)))
        ax2.set_yticklabels(top_pct['td_desc'].str[:20])
        ax2.set_title(f'Top {n_changes} Percentage Changes')
        
        # Volume impact
        volume_impact = df.copy()
        volume_impact['volume_change'] = volume_impact['change'] * volume_impact['orig_qty']
        volume_impact = volume_impact.nlargest(n_changes, 'volume_change')
        
        ax3.barh(range(len(volume_impact)), volume_impact['volume_change'])
        ax3.set_yticks(range(len(volume_impact)))
        ax3.set_yticklabels(volume_impact['td_desc'].str[:20])
        ax3.set_title(f'Top {n_changes} Volume Impact')
        
        # Change over time
        sorted_changes = df.sort_values('change', ascending=True)
        cumsum = sorted_changes['change'].cumsum()
        ax4.plot(range(len(cumsum)), cumsum)
        ax4.set_title('Cumulative Change Impact')
        ax4.set_xlabel('Number of Promotions')
        ax4.set_ylabel('Cumulative Change')
        
        plt.tight_layout()
    
    def create_customer_group_analysis(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        
        df_with_customer = df.merge(self.original_df[['Td No', 'Customer Group']], 
                                  left_on='td_no', right_on='Td No')
        
        # Average change by customer
        avg_by_customer = df_with_customer.groupby('Customer Group')['change'].mean().sort_values()
        avg_by_customer.plot(kind='barh', ax=ax1)
        ax1.set_title('Average Change by Customer Group')
        
        # Volume by customer
        volume_by_customer = df_with_customer.groupby('Customer Group')['updated_qty'].sum()
        volume_by_customer.plot(kind='pie', ax=ax2, autopct='%1.1f%%')
        ax2.set_title('Volume Distribution by Customer')
        
        # Change distribution by customer
        sns.boxplot(data=df_with_customer, x='Customer Group', y='change', ax=ax3)
        ax3.set_xticklabels(ax3.get_xticklabels(), rotation=45)
        ax3.set_title('Change Distribution by Customer')
        
        # Customer performance metrics
        customer_perf = df_with_customer.groupby('Customer Group').agg({
            'change': ['mean', 'std', 'count']
        }).plot(kind='bar', ax=ax4)
        ax4.set_xticklabels(ax4.get_xticklabels(), rotation=45)
        ax4.set_title('Customer Group Performance Metrics')
    
    def create_timeline_analysis(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        
        df_with_dates = df.merge(
            self.original_df[['Td No', 'Start Date', 'End Date']], 
            left_on='td_no', right_on='Td No'
        )
        
        # Convert dates
        df_with_dates['Start Date'] = pd.to_datetime(df_with_dates['Start Date'])
        df_with_dates['End Date'] = pd.to_datetime(df_with_dates['End Date'])
        df_with_dates['Duration'] = (df_with_dates['End Date'] - df_with_dates['Start Date']).dt.days
        
        # Changes over time
        timeline = df_with_dates.sort_values('Start Date')
        ax1.plot(timeline['Start Date'], timeline['change'].cumsum())
        ax1.set_title('Cumulative Changes Over Time')
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
        
        # Duration impact
        ax2.scatter(df_with_dates['Duration'], df_with_dates['change'])
        ax2.set_title('Promotion Duration vs Change')
        ax2.set_xlabel('Duration (days)')
        ax2.set_ylabel('Change Amount')
        
        # Monthly pattern
        monthly_changes = df_with_dates.groupby(df_with_dates['Start Date'].dt.month)['change'].mean()
        monthly_changes.plot(kind='bar', ax=ax3)
        ax3.set_title('Average Change by Month')
        ax3.set_xlabel('Month')
        
        # Duration distribution
        sns.histplot(data=df_with_dates, x='Duration', ax=ax4)
        ax4.set_title('Distribution of Promotion Durations')
    
    def create_quantity_range_analysis(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        
        # Create quantity ranges
        df['qty_range'] = pd.qcut(df['orig_qty'], 
                                q=5, 
                                labels=['Very Low', 'Low', 'Medium', 'High', 'Very High'])
        
        # Change distribution by range
        sns.boxplot(data=df, x='qty_range', y='change', ax=ax1)
        ax1.set_title('Change Distribution by Quantity Range')
        ax1.set_xticklabels(ax1.get_xticklabels(), rotation=45)
        
        # Range composition
        df['qty_range'].value_counts().plot(kind='pie', ax=ax2, autopct='%1.1f%%')
        ax2.set_title('Distribution of Quantity Ranges')
        
        # Average change by range
        df.groupby('qty_range')['change'].mean().plot(kind='bar', ax=ax3)
        ax3.set_title('Average Change by Quantity Range')
        ax3.set_xticklabels(ax3.get_xticklabels(), rotation=45)
        
        # Quantity vs Percentage change
        pct_change = (df['change'] / df['orig_qty'] * 100).replace([np.inf, -np.inf], 0)
        ax4.scatter(df['orig_qty'], pct_change)
        ax4.set_title('Original Quantity vs Percentage Change')
        ax4.set_xlabel('Original Quantity')
        ax4.set_ylabel('Percentage Change')
    
    def create_change_patterns(self, df):
        self.fig, ((ax1, ax2), (ax3, ax4)) = plt.subplots(2, 2, figsize=(12, 10))
        # Change frequency distribution
        sns.histplot(data=df, x='change', bins=30, ax=ax1)
        ax1.axvline(x=0, color='r', linestyle='--')
        ax1.set_title('Distribution of Changes')
        
        # Original vs Updated correlation
        ax2.scatter(df['orig_qty'], df['updated_qty'])
        ax2.plot([df['orig_qty'].min(), df['orig_qty'].max()],
                 [df['orig_qty'].min(), df['orig_qty'].max()],
                 'r--')
        ax2.set_title('Original vs Updated Quantities')
        ax2.set_xlabel('Original Quantity')
        ax2.set_ylabel('Updated Quantity')
        
        # Change magnitude categories
        df['change_cat'] = pd.cut(df['change'], 
                                bins=5, 
                                labels=['Large Decrease', 'Small Decrease', 'Minimal Change',
                                       'Small Increase', 'Large Increase'])
        df['change_cat'].value_counts().plot(kind='pie', ax=ax3, autopct='%1.1f%%')
        ax3.set_title('Distribution of Change Categories')
        
        # Change impact vs Original Quantity
        ax4.scatter(df['orig_qty'], abs(df['change']))
        ax4.set_title('Change Magnitude vs Original Quantity')
        ax4.set_xlabel('Original Quantity')
        ax4.set_ylabel('Absolute Change')
    
    def save_chart(self):
        if self.fig is None:
            messagebox.showwarning("Warning", "No chart to save")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[
                ("PNG files", "*.png"),
                ("JPEG files", "*.jpg"),
                ("PDF files", "*.pdf"),
                ("All files", "*.*")
            ]
        )
        
        if filename:
            try:
                self.fig.savefig(filename, bbox_inches='tight', dpi=300)
                messagebox.showinfo("Success", f"Chart saved successfully to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving chart: {str(e)}")

class PromotionAnalysisGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Promotion Analysis Dashboard")
        self.root.geometry("1400x800")
        
        plt.style.use('default')
        sns.set_theme(style="whitegrid")
        
        self.comparison_results = []
        self.original_df = None
        self.updated_df = None
        
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.comparison_tab = ttk.Frame(self.notebook)
        self.analysis_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.comparison_tab, text="Comparison View")
        self.notebook.add(self.analysis_tab, text="Analysis")
        
        self.setup_comparison_tab()
        self.setup_analysis_tab()
    
    def setup_comparison_tab(self):
        input_frame = ttk.LabelFrame(self.comparison_tab, text="Data Input", padding=10)
        input_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(input_frame, text="Original Data:").pack(anchor="w")
        self.original_text = tk.Text(input_frame, height=6, width=100)
        self.original_text.pack(fill="x", pady=5)
        
        ttk.Label(input_frame, text="Updated Data:").pack(anchor="w")
        self.updated_text = tk.Text(input_frame, height=6, width=100)
        self.updated_text.pack(fill="x", pady=5)
        
        control_frame = ttk.Frame(input_frame)
        control_frame.pack(fill="x", pady=5)
        
        ttk.Button(control_frame, text="Compare Data", 
                  command=self.compare_data).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Clear Data", 
                  command=self.clear_data).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Export Results", 
                  command=self.export_to_excel).pack(side="left", padx=5)
        ttk.Button(control_frame, text="View Charts", 
                  command=self.open_chart_viewer).pack(side="left", padx=5)
        
        filter_frame = ttk.LabelFrame(self.comparison_tab, text="Filters", padding=10)
        filter_frame.pack(fill="x", padx=5, pady=5)
        
        self.filter_var = tk.StringVar(value="all")
        ttk.Radiobutton(filter_frame, text="Show All", variable=self.filter_var, 
                       value="all", command=self.apply_filter).pack(side="left", padx=5)
        ttk.Radiobutton(filter_frame, text="Show Changes Only", variable=self.filter_var, 
                       value="changes", command=self.apply_filter).pack(side="left", padx=5)
        
        results_frame = ttk.LabelFrame(self.comparison_tab, text="Comparison Results", padding=10)
        results_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.tree = ttk.Treeview(results_frame, show="headings")
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(column=0, row=0, sticky="nsew")
        vsb.grid(column=1, row=0, sticky="ns")
        hsb.grid(column=0, row=1, sticky="ew")
        
        results_frame.grid_columnconfigure(0, weight=1)
        results_frame.grid_rowconfigure(0, weight=1)
        
        columns = ("Td No", "Td Desc", "Balance Qty (Original)", "Balance Qty (Updated)", "Change")
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, minwidth=100, width=200)
        
        self.tree.tag_configure('increase', background='#90EE90')
        self.tree.tag_configure('decrease', background='#FFB6C1')
    
    def setup_analysis_tab(self):
        self.summary_frame = ttk.LabelFrame(self.analysis_tab, text="Summary Statistics", padding=10)
        self.summary_frame.pack(fill="x", padx=5, pady=5)
        
        analysis_frame = ttk.LabelFrame(self.analysis_tab, text="Detailed Analysis", padding=10)
        analysis_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        self.analysis_tree = ttk.Treeview(analysis_frame, show="headings")
        vsb = ttk.Scrollbar(analysis_frame, orient="vertical", command=self.analysis_tree.yview)
        hsb = ttk.Scrollbar(analysis_frame, orient="horizontal", command=self.analysis_tree.xview)
        
        self.analysis_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.analysis_tree.grid(column=0, row=0, sticky="nsew")
        vsb.grid(column=1, row=0, sticky="ns")
        hsb.grid(column=0, row=1, sticky="ew")
        
        analysis_frame.grid_columnconfigure(0, weight=1)
        analysis_frame.grid_rowconfigure(0, weight=1)
    
    def clear_data(self):
        self.original_text.delete(1.0, tk.END)
        self.updated_text.delete(1.0, tk.END)
        self.tree.delete(*self.tree.get_children())
        self.comparison_results.clear()
        self.original_df = None
        self.updated_df = None
        self.clear_analysis()
    
    def clear_analysis(self):
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        self.analysis_tree.delete(*self.analysis_tree.get_children())
    
    def parse_excel_data(self, text_data):
        try:
            lines = text_data.strip().split('\n')
            headers = lines[0].split('\t')
            data = [line.split('\t') for line in lines[1:]]
            
            df = pd.DataFrame(data, columns=headers)
            df['Balance Qty'] = pd.to_numeric(df['Balance Qty'], errors='coerce').fillna(0)
            return df
        except Exception as e:
            messagebox.showerror("Error", f"Error parsing data: {str(e)}")
            return None
    
    def compare_data(self):
        self.tree.delete(*self.tree.get_children())
        self.comparison_results.clear()
        
        original_text = self.original_text.get("1.0", tk.END).strip()
        updated_text = self.updated_text.get("1.0", tk.END).strip()
        
        if not original_text or not updated_text:
            messagebox.showwarning("Warning", "Please paste both original and updated data")
            return
        
        self.original_df = self.parse_excel_data(original_text)
        self.updated_df = self.parse_excel_data(updated_text)
        
        if self.original_df is None or self.updated_df is None:
            return
        
        try:
            for _, orig_row in self.original_df.iterrows():
                td_no = orig_row['Td No']
                updated_row = self.updated_df[self.updated_df['Td No'] == td_no]
                
                if not updated_row.empty:
                    orig_qty = float(orig_row['Balance Qty'])
                    updated_qty = float(updated_row.iloc[0]['Balance Qty'])
                    change = updated_qty - orig_qty
                    
                    self.comparison_results.append({
                        'td_no': td_no,
                        'td_desc': orig_row['Td Desc'],
                        'orig_qty': orig_qty,
                        'updated_qty': updated_qty,
                        'change': change
                    })
            
            self.apply_filter()
            self.update_analysis()
            
        except Exception as e:
            messagebox.showerror("Error", f"Error during comparison: {str(e)}")
    
    def apply_filter(self):
        self.tree.delete(*self.tree.get_children())
        
        for result in self.comparison_results:
            if self.filter_var.get() == "all" or (
                self.filter_var.get() == "changes" and result['change'] != 0):
                
                tag = 'increase' if result['change'] > 0 else (
                    'decrease' if result['change'] < 0 else '')
                
                self.tree.insert('', 'end', values=(
                    result['td_no'],
                    result['td_desc'],
                    f"{result['orig_qty']:.0f}",
                    f"{result['updated_qty']:.0f}",
                    f"{result['change']:+.0f}"
                ), tags=(tag,))
    
    def update_analysis(self):
        for widget in self.summary_frame.winfo_children():
            widget.destroy()
        
        if not self.comparison_results:
            ttk.Label(self.summary_frame, text="No data to analyze").pack(anchor="w")
            return
        
        changes_df = pd.DataFrame(self.comparison_results)
        total_promos = len(changes_df)
        changes_only = changes_df[changes_df['change'] != 0]
        
        stats = [
            f"Total Promotions Analyzed: {total_promos}",
            f"Promotions with Changes: {len(changes_only)} ({(len(changes_only)/total_promos*100):.1f}%)",
            f"Total Net Change in Balance: {changes_df['change'].sum():,.0f}",
            f"Average Change per Promotion: {changes_df['change'].mean():,.1f}",
            "",
            "Change Distribution:",
            f"Increases: {len(changes_df[changes_df['change'] > 0])} promos",
            f"Decreases: {len(changes_df[changes_df['change'] < 0])} promos",
            f"No Change: {len(changes_df[changes_df['change'] == 0])} promos",
            "",
            "Magnitude of Changes:",
            f"Largest Increase: {changes_df['change'].max():,.0f}",
            f"Largest Decrease: {changes_df['change'].min():,.0f}",
            f"Standard Deviation: {changes_df['change'].std():,.1f}"
        ]
        
        for stat in stats:
            ttk.Label(self.summary_frame, text=stat).pack(anchor="w", pady=2)
        
        self.analysis_tree["columns"] = ["Category", "Details", "Value"]
        for col in self.analysis_tree["columns"]:
            self.analysis_tree.heading(col, text=col)
            self.analysis_tree.column(col, width=200)
        
        self.analysis_tree.delete(*self.analysis_tree.get_children())
        
        significant_increases = len(changes_df[changes_df['change'] > changes_df['orig_qty'] * 0.5])
        significant_decreases = len(changes_df[changes_df['change'] < -changes_df['orig_qty'] * 0.5])
        total_positive_change = changes_df[changes_df['change'] > 0]['change'].sum()
        total_negative_change = abs(changes_df[changes_df['change'] < 0]['change'].sum())
        
        analysis_items = [
            ("Significant Changes", "Promotions with >50% increase", significant_increases),
            ("", "Promotions with >50% decrease", significant_decreases),
            ("Volume Analysis", "Total volume of increases", f"{total_positive_change:,.0f}"),
            ("", "Total volume of decreases", f"{total_negative_change:,.0f}"),
            ("", "Net volume change", f"{(total_positive_change - total_negative_change):,.0f}")
        ]
        
        for category, details, value in analysis_items:
            self.analysis_tree.insert("", "end", values=(category, details, value))
    
    def open_chart_viewer(self):
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No data to display charts")
            return
        ChartViewer(self.root, self.comparison_results, self.original_df, self.updated_df)
    
    def export_to_excel(self):
        if not self.comparison_results:
            messagebox.showwarning("Warning", "No data to export")
            return
        
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile=f"promotion_comparison_{timestamp}.xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
            if not filename:
                return
            
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                # Comparison Results sheet
                comparison_df = pd.DataFrame(self.comparison_results)
                comparison_df.columns = ['TD No', 'TD Description', 'Original Qty', 
                                      'Updated Qty', 'Change']
                comparison_df.to_excel(writer, sheet_name='Comparison Results', 
                                    index=False)
                
                # Summary Statistics sheet
                summary_data = {
                    'Metric': [
                        'Total Promotions',
                        'Total Changes',
                        'Increases',
                        'Decreases',
                        'No Change',
                        'Average Change',
                        'Largest Increase',
                        'Largest Decrease'
                    ],
                    'Value': [
                        len(self.comparison_results),
                        sum(1 for r in self.comparison_results if r['change'] != 0),
                        sum(1 for r in self.comparison_results if r['change'] > 0),
                        sum(1 for r in self.comparison_results if r['change'] < 0),
                        sum(1 for r in self.comparison_results if r['change'] == 0),
                        sum(r['change'] for r in self.comparison_results) / 
                        len(self.comparison_results) if self.comparison_results else 0,
                        max(r['change'] for r in self.comparison_results),
                        min(r['change'] for r in self.comparison_results)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                summary_df.to_excel(writer, sheet_name='Summary Statistics', 
                                 index=False)
            
            messagebox.showinfo("Success", f"Data exported successfully to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to Excel: {str(e)}")

def main():
    root = tk.Tk()
    app = PromotionAnalysisGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()