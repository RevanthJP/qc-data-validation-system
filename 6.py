import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkFont

# ----------------------------------------------------------------------
# Token preservation logic (incorporated as requested)
# ----------------------------------------------------------------------
DEFAULT_TOKENS = ["None", "N/A", "NA", "#N/A"]

def load_excel(path, sheet, preserve_tokens=True):
    """
    Load Excel with optional token preservation so strings like 'None'/'N/A'
    are not auto-parsed as NaN.
    """
    if preserve_tokens:
        return pd.read_excel(
            path,
            sheet_name=sheet,
            dtype=str,             # Force strings so tokens stay literal
            keep_default_na=False, # Do not convert blanks or tokens to NaN
            na_values=[]           # No additional NA parsing
        )
    else:
        return pd.read_excel(path, sheet_name=sheet)


def normalize_blanks(df):
    """
    Convert empty/whitespace-only strings to true blanks "" (not NaN),
    while leaving tokens like 'None'/'N/A' intact (since they are literal strings).
    """
    # Replace whitespace-only with empty string, keep tokens as-is
    return df.replace(r'^\s*$', '', regex=True)

# ----------------------------------------------------------------------
# Existing app logic (unchanged except value handling/reading)
# ----------------------------------------------------------------------

def get_excel_file_from_input(input_path):
    if os.path.isfile(input_path):
        if input_path.lower().endswith(('.xls', '.xlsx')):
            return input_path
        else:
            raise ValueError("Provided file path is not an Excel file")
    elif os.path.isdir(input_path):
        files = os.listdir(input_path)
        excel_files = [f for f in files if f.lower().endswith(('.xls', '.xlsx'))]
        if not excel_files:
            raise FileNotFoundError("No Excel file found in directory")
        return os.path.join(input_path, excel_files[0])
    else:
        raise FileNotFoundError("The path does not exist")


def qc_comparison_script_gui(input_key_path, input_agent_path, output_file_path, progress_var, status_label):
    try:
        status_label.config(text="Processing files...")
        progress_var.set(10)
        keysheet_file = get_excel_file_from_input(input_key_path)
        agent_file = get_excel_file_from_input(input_agent_path)
        progress_var.set(20)
    except Exception as e:
        messagebox.showerror("Error", f"Error resolving input files:\n{e}")
        status_label.config(text="Error: Failed to locate files")
        progress_var.set(0)
        return

    try:
        status_label.config(text="Reading Excel files...")
        key_xl = pd.ExcelFile(keysheet_file)
        agent_xl = pd.ExcelFile(agent_file)
        keysheet_sheets = key_xl.sheet_names
        agent_sheets = agent_xl.sheet_names

        target_sheet = None
        possible_sheet_names = ['Production & QC Report', 'Production &amp; QC Report', 'Production and QC Report']
        for sheet_name in possible_sheet_names:
            if sheet_name in keysheet_sheets and sheet_name in agent_sheets:
                target_sheet = sheet_name
                break
        if target_sheet is None:
            error_msg = (
                "Sheet 'Production & QC Report' not found.\n\n"
                f"Available sheets in keysheet: {keysheet_sheets}\n"
                f"Available sheets in agent file: {agent_sheets}"
            )
            messagebox.showerror("Sheet Not Found", error_msg)
            status_label.config(text="Error: Required sheet not found")
            progress_var.set(0)
            return

        # USE TOKEN-PRESERVING LOAD FOR BOTH FILES
        keysheet_df = load_excel(keysheet_file, target_sheet, preserve_tokens=True)
        agent_df   = load_excel(agent_file,    target_sheet, preserve_tokens=True)

        # NORMALIZE BLANKS ONLY (tokens remain literal strings)
        keysheet_df = normalize_blanks(keysheet_df)
        agent_df    = normalize_blanks(agent_df)

        progress_var.set(40)
    except Exception as e:
        messagebox.showerror("Error", f"Error reading Excel files:\n{str(e)}")
        status_label.config(text="Error: Failed to read Excel files")
        progress_var.set(0)
        return

    if 'Item ID' not in keysheet_df.columns:
        messagebox.showerror("Error", f"'Item ID' column not found in keysheet file.\nAvailable columns: {list(keysheet_df.columns)}")
        status_label.config(text="Error: Item ID column missing")
        progress_var.set(0)
        return
    if 'Item ID' not in agent_df.columns:
        messagebox.showerror("Error", f"'Item ID' column not found in agent file.\nAvailable columns: {list(agent_df.columns)}")
        status_label.config(text="Error: Item ID column missing")
        progress_var.set(0)
        return

    # Unchanged comparison constraints
    exclude_columns = [
        'Date (YYYY/MM/DD)', 'Work Type', 'Associate Walmart ID', 'Item Status',
        'Submission ID', 'Product ID Type', 'Product ID', 'Audit Template Version',
        'SOP Reference', 'Initial CQ Score', 'Current CQ Score', 'Supplier ID',
        'Item Created Date', 'Active Status', 'is Private Label?', 'Website Link',
        'View Images'
    ]
    product_type_columns = [
        'Correct Product Type', 'Is Product Type Validated',
        'Product Error Code', 'Product Type Validation Comment'
    ]
    product_name_columns = [
        'Correct Product Name', 'Is Product Name Validated',
        'Product Name Error Code', 'Product Name Validation Comment'
    ]

    status_label.config(text="Comparing data...")
    merged_df = pd.merge(keysheet_df, agent_df, on='Item ID', how='inner', suffixes=('_key', '_agent'))

    if len(merged_df) == 0:
        messagebox.showwarning("No Matches", "No matching Item IDs found between keysheet and agent files.")
        status_label.config(text="Warning: No matching Item IDs found")
        progress_var.set(0)
        return

    progress_var.set(60)

    discrepancy_records = []
    keysheet_columns = keysheet_df.columns.tolist()

    for index, row in merged_df.iterrows():
        item_id = row['Item ID']
        associate_id = row.get('Associate Walmart ID_agent', row.get('Associate Walmart ID_key', ''))

        for col in keysheet_columns:
            if col in exclude_columns or col == 'Item ID':
                continue
            if col.startswith('Pre - '):
                continue
            if 'Product Type' in col and col not in product_type_columns:
                continue
            if 'Product Name' in col and col not in product_name_columns:
                continue

            agent_col = col + '_agent'
            key_col   = col + '_key'
            if agent_col in merged_df.columns:
                # Values are strings (or ''), with tokens preserved
                key_value   = row.get(key_col)
                agent_value = row.get(agent_col)

                # Ensure None types (if any) become '' to avoid literal None in file
                if key_value is None:
                    key_value = ''
                if agent_value is None:
                    agent_value = ''

                if str(key_value) != str(agent_value):
                    attribute_name = col
                    error_column   = col

                    if any(k in col.lower() for k in ['is ', 'validated', 'error code', 'validation comment']):
                        if 'Is ' in col and 'Validated' in col:
                            attribute_name = col.replace('Is ', '').replace(' Validated', '').replace(' Validate', '')
                        elif 'Error Code' in col:
                            attribute_name = col.replace(' Error Code', '')
                        elif 'Validation Comment' in col:
                            attribute_name = col.replace(' Validation Comment', '')
                        error_column = col
                    else:
                        attribute_name = col
                        error_column   = col

                    discrepancy_records.append({
                        'Item ID': item_id,
                        'Associate Walmart ID': associate_id,
                        'Attribute ': attribute_name,
                        'Error column': error_column,
                        'Key sheet Value': str(key_value),
                        'Agent Value': str(agent_value)
                    })

    error_df = pd.DataFrame(discrepancy_records)
    progress_var.set(80)

    # Agent-level aggregation (unchanged)
    if len(error_df) > 0:
        agent_stats = []
        unique_agents = merged_df['Associate Walmart ID_agent'].dropna().unique() if 'Associate Walmart ID_agent' in merged_df.columns else []
        for agent in unique_agents:
            agent_work = merged_df[merged_df.get('Associate Walmart ID_agent') == agent]
            total_items = len(agent_work)
            comparable_columns = [c for c in keysheet_columns
                                  if c not in exclude_columns
                                  and c != 'Item ID'
                                  and not c.startswith('Pre - ')
                                  and ((('Product Type' not in c) or (c in product_type_columns))
                                       and (('Product Name' not in c) or (c in product_name_columns)))]
            total_attributes = total_items * len(comparable_columns)
            agent_errors = error_df[error_df['Associate Walmart ID'] == agent]
            attribute_errors = len(agent_errors)
            error_percentage = attribute_errors / total_attributes if total_attributes > 0 else 0
            agent_stats.append({
                'Associate Walmart ID': agent,
                'Total Item ID worked': total_items,
                'Total attribute Worked': total_attributes,
                'Attribute Error': attribute_errors,
                'Attribute Error Percentage': error_percentage
            })
        agent_level_df = pd.DataFrame(agent_stats)
    else:
        agent_level_df = pd.DataFrame(columns=[
            'Associate Walmart ID', 'Total Item ID worked', 'Total attribute Worked',
            'Attribute Error', 'Attribute Error Percentage'
        ])

    try:
        status_label.config(text="Saving output file...")
        output_dir = os.path.dirname(output_file_path)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
            error_df.to_excel(writer, sheet_name='Error sheet', index=False)
            agent_level_df.to_excel(writer, sheet_name='Agent level error sheet', index=False)
        progress_var.set(100)
        status_label.config(text=f"Success! Report saved: {os.path.basename(output_file_path)} ({len(error_df)} discrepancies)")
        messagebox.showinfo("Success", f"QC Report generated successfully!\n\nDiscrepancies found: {len(error_df)}\nAgent records: {len(agent_level_df)}\nOutput saved at: {output_file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Error writing output file:\n{e}")
        status_label.config(text="Error: Failed to save output file")
        progress_var.set(0)


class QCComparisonGUI:
    def __init__(self, root):
        self.root = root
        self.setup_gui()

    def setup_gui(self):
        self.root.title("QC Comparison Tool - Token Aware")
        self.root.geometry("700x350")
        self.root.resizable(False, False)
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        title_label = ttk.Label(main_frame, text="Quality Control Comparison Tool", font=("Arial", 14, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        ttk.Label(main_frame, text="Keysheet File/Folder:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.keysheet_entry = ttk.Entry(main_frame, width=50)
        self.keysheet_entry.grid(row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Browse", command=self.browse_keysheet_file).grid(row=1, column=2, padx=5)
        ttk.Label(main_frame, text="Agent Production File/Folder:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.agent_entry = ttk.Entry(main_frame, width=50)
        self.agent_entry.grid(row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Browse", command=self.browse_agent_file).grid(row=2, column=2, padx=5)
        ttk.Label(main_frame, text="Output Excel File:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(main_frame, width=50)
        self.output_entry.grid(row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Save As", command=self.browse_output_file).grid(row=3, column=2, padx=5)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(main_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(20, 5))
        self.status_label = ttk.Label(main_frame, text="Ready to process files", foreground="blue")
        self.status_label.grid(row=5, column=0, columnspan=3, pady=5)
        run_button = ttk.Button(main_frame, text="Run QC Comparison", command=self.run_comparison)
        run_button.grid(row=6, column=1, pady=20)
        main_frame.columnconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def browse_keysheet_file(self):
        filename = filedialog.askopenfilename(title="Select Keysheet File", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if filename:
            self.keysheet_entry.delete(0, tk.END)
            self.keysheet_entry.insert(0, filename)

    def browse_agent_file(self):
        filename = filedialog.askopenfilename(title="Select Agent Production File", filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")])
        if filename:
            self.agent_entry.delete(0, tk.END)
            self.agent_entry.insert(0, filename)

    def browse_output_file(self):
        filename = filedialog.asksaveasfilename(title="Save QC Report As", defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def run_comparison(self):
        input_key = self.keysheet_entry.get().strip()
        input_agent = self.agent_entry.get().strip()
        output_path = self.output_entry.get().strip()
        if not input_key or not input_agent or not output_path:
            messagebox.showerror("Error", "Please provide all file paths before running the comparison.")
            return
        self.progress_var.set(0)
        qc_comparison_script_gui(input_key, input_agent, output_path, self.progress_var, self.status_label)

def main():
    root = tk.Tk()
    app = QCComparisonGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
