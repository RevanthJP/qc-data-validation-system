import os
import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import threading
from pathlib import Path

class ModernQCComparisonGUI:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.setup_styles()
        self.create_widgets()
        
    def setup_window(self):
        """Configure the main window"""
        self.root.title("Training QC Tool - Normal Validation")
        self.root.geometry("800x600")
        self.root.minsize(700, 500)
        
        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (800 // 2)
        y = (self.root.winfo_screenheight() // 2) - (600 // 2)
        self.root.geometry(f"800x600+{x}+{y}")
        
        # Configure background color
        self.root.configure(bg='#f0f0f0')
        
    def setup_styles(self):
        """Configure modern styles"""
        self.style = ttk.Style()
        
        # Configure styles for modern look
        self.style.configure('Title.TLabel', 
                           font=('Segoe UI', 20, 'bold'),
                           background='#f0f0f0',
                           foreground='#2c3e50')
        
        self.style.configure('Subtitle.TLabel',
                           font=('Segoe UI', 10),
                           background='#f0f0f0',
                           foreground='#7f8c8d')
        
        self.style.configure('Modern.TLabel',
                           font=('Segoe UI', 10),
                           background='#ffffff',
                           foreground='#2c3e50')
        
        self.style.configure('Modern.TButton',
                           font=('Segoe UI', 10),
                           padding=(20, 10))
        
        self.style.configure('Browse.TButton',
                           font=('Segoe UI', 9),
                           padding=(10, 5))
        
        self.style.configure('Modern.TEntry',
                           font=('Segoe UI', 10),
                           fieldbackground='#ffffff')
        
    def create_widgets(self):
        """Create all GUI widgets"""
        # Main container
        main_container = ttk.Frame(self.root, padding="30")
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Title section
        title_frame = ttk.Frame(main_container)
        title_frame.pack(fill=tk.X, pady=(0, 30))
        
        title_label = ttk.Label(title_frame, text="Training QC Tool - Normal Validation", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(title_frame, 
                                 text="Compare keysheet and agent production files to generate quality control reports",
                                 style='Subtitle.TLabel')
        subtitle_label.pack(pady=(5, 0))
        
        # Input section
        input_frame = ttk.LabelFrame(main_container, text="File Selection", padding="20")
        input_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Keysheet file selection
        self.create_file_input(input_frame, "Keysheet File:", "keysheet", 0)
        
        # Agent file selection
        self.create_file_input(input_frame, "Agent Production File:", "agent", 1)
        
        # Output file selection
        self.create_output_input(input_frame, 2)
        
        # Action buttons
        button_frame = ttk.Frame(main_container)
        button_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Create a sub-frame to center the buttons
        center_frame = ttk.Frame(button_frame)
        center_frame.pack(expand=True)
        
        # Compare button
        self.compare_btn = ttk.Button(center_frame, 
                                    text="Generate QC Report",
                                    style='Modern.TButton',
                                    command=self.run_comparison)
        self.compare_btn.pack(side=tk.LEFT, padx=(0, 20))
        
        # Clear button
        clear_btn = ttk.Button(center_frame,
                             text="Clear All",
                             style='Modern.TButton',
                             command=self.clear_all)
        clear_btn.pack(side=tk.LEFT)
        
        # Progress bar
        self.progress = ttk.Progressbar(main_container, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # Output/Log section
        log_frame = ttk.LabelFrame(main_container, text="Output Log", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = ScrolledText(log_frame, 
                                   height=12,
                                   font=('Consolas', 9),
                                   bg='#2c3e50',
                                   fg='#ecf0f1',
                                   insertbackground='#ecf0f1')
        self.log_text.pack(fill=tk.BOTH, expand=True)
        
        # Initial log message
        self.log("Training QC Tool initialized. Select files to begin.")
        
    def create_file_input(self, parent, label_text, var_name, row):
        """Create a file input row with label, entry, and browse button"""
        # Label
        label = ttk.Label(parent, text=label_text, style='Modern.TLabel')
        label.grid(row=row, column=0, sticky=tk.W, pady=5)
        
        # Entry
        entry_var = tk.StringVar()
        setattr(self, f"{var_name}_var", entry_var)
        
        entry = ttk.Entry(parent, textvariable=entry_var, style='Modern.TEntry', width=50)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=(10, 5), pady=5)
        
        # Browse button
        browse_btn = ttk.Button(parent, text="Browse", style='Browse.TButton',
                              command=lambda: self.browse_file(var_name))
        browse_btn.grid(row=row, column=2, padx=(5, 0), pady=5)
        
        # Configure column weights
        parent.columnconfigure(1, weight=1)
        
    def create_output_input(self, parent, row):
        """Create output file selection"""
        label = ttk.Label(parent, text="Output File:", style='Modern.TLabel')
        label.grid(row=row, column=0, sticky=tk.W, pady=5)
        
        self.output_var = tk.StringVar()
        self.output_var.set("")
        
        entry = ttk.Entry(parent, textvariable=self.output_var, style='Modern.TEntry', width=50)
        entry.grid(row=row, column=1, sticky=tk.EW, padx=(10, 5), pady=5)
        
        browse_btn = ttk.Button(parent, text="Save As", style='Browse.TButton',
                              command=self.browse_output_file)
        browse_btn.grid(row=row, column=2, padx=(5, 0), pady=5)
        
    def browse_file(self, var_name):
        """Browse for input files"""
        var = getattr(self, f"{var_name}_var")
        
        filename = filedialog.askopenfilename(
            title=f"Select {var_name.title()} File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if filename:
            var.set(filename)
            self.log(f"{var_name.title()} path set: {filename}")
            
    def browse_output_file(self):
        """Browse for output file location"""
        filename = filedialog.asksaveasfilename(
            title="Save QC Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if filename:
            self.output_var.set(filename)
            self.log(f"Output path set: {filename}")
            
    def clear_all(self):
        """Clear all input fields"""
        self.keysheet_var.set("")
        self.agent_var.set("")
        self.output_var.set("qc_report.xlsx")
        self.log_text.delete(1.0, tk.END)
        self.log("All fields cleared.")
        
    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def validate_inputs(self):
        """Validate user inputs"""
        keysheet_path = self.keysheet_var.get().strip()
        agent_path = self.agent_var.get().strip()
        output_path = self.output_var.get().strip()
        
        if not keysheet_path:
            messagebox.showerror("Error", "Please select a keysheet file.")
            return False
            
        if not agent_path:
            messagebox.showerror("Error", "Please select an agent production file.")
            return False
            
        if not output_path:
            messagebox.showerror("Error", "Please specify an output file path.")
            return False
            
        # Ensure output file has .xlsx extension
        if not output_path.lower().endswith('.xlsx'):
            self.output_var.set(output_path + '.xlsx')
            
        return True
        
    def run_comparison(self):
        """Run the QC comparison in a separate thread"""
        if not self.validate_inputs():
            return
            
        # Disable buttons during processing
        self.compare_btn.configure(state='disabled')
        self.progress.start()
        
        # Run in separate thread to prevent GUI freezing
        thread = threading.Thread(target=self._run_comparison_thread)
        thread.daemon = True
        thread.start()
        
    def _run_comparison_thread(self):
        """Thread function for running comparison"""
        try:
            keysheet_path = self.keysheet_var.get().strip()
            agent_path = self.agent_var.get().strip()
            output_path = self.output_var.get().strip()
            
            self.log("=" * 50)
            self.log("Starting QC Comparison Process...")
            
            # Load requirement level information before running comparison
            requirement_level_info = self.get_requirement_level_info(keysheet_path)
            self.log(f"Loaded requirement level information for {len(requirement_level_info)} attributes")
            
            # Run the comparison using the original logic
            self.qc_comparison_script(keysheet_path, agent_path, output_path, requirement_level_info)
            
        except Exception as e:
            self.log(f"Error during comparison: {str(e)}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
        finally:
            # Re-enable buttons
            self.root.after(0, self._finish_comparison)
            
    def _finish_comparison(self):
        """Clean up after comparison is finished"""
        self.progress.stop()
        self.compare_btn.configure(state='normal')

    # Original functions with logging modifications
    def get_excel_file_from_input(self, input_path):
        """Original function with GUI logging"""
        if os.path.isfile(input_path):
            if input_path.lower().endswith(('.xls', '.xlsx')):
                return input_path
            else:
                raise ValueError(f"Provided file path is not an Excel file: {input_path}")
        elif os.path.isdir(input_path):
            files = os.listdir(input_path)
            excel_files = [f for f in files if f.lower().endswith(('.xls', '.xlsx'))]
            
            if not excel_files:
                raise FileNotFoundError(f"No Excel file found in directory: {input_path}")
            
            if len(excel_files) > 1:
                self.log(f"Multiple Excel files found in {input_path}:")
                for i, file in enumerate(excel_files, 1):
                    self.log(f"{i}. {file}")
                self.log(f"Using the first file: {excel_files[0]}")
            
            return os.path.join(input_path, excel_files[0])
        else:
            raise FileNotFoundError(f"The path does not exist: {input_path}")

    def get_requirement_level_info(self, keysheet_path):
        """Get Product Type and Requirement Level information from the keysheet"""
        try:
            # Read the Requirement level sheet from the keysheet
            req_level_df = pd.read_excel(keysheet_path, sheet_name='Requirement Level')
            
            # Log the column names for debugging
            self.log("Requirement Level sheet columns:")
            self.log(", ".join(req_level_df.columns.tolist()))
            
            # Create a dictionary for quick lookup, using lowercase keys for case-insensitive matching
            req_info = {}
            for _, row in req_level_df.iterrows():
                # Try different possible column names for Attribute
                for possible_attr_col in ['Attribute', 'Attribute ', 'Attributes']:
                    if possible_attr_col in row.index:
                        attribute = str(row[possible_attr_col]).strip()
                        if attribute and attribute.lower() != 'nan':
                            # Log each attribute being processed
                            self.log(f"Processing attribute: {attribute}")
                            
                            # Get Product Type and Requirement Level, checking different possible column names
                            product_type = ''
                            for pt_col in ['Product Type', 'ProductType', 'Product_Type']:
                                if pt_col in row.index:
                                    product_type = str(row[pt_col]).strip()
                                    break
                            
                            req_level = ''
                            for rl_col in ['Requirement Level', 'Requirement level', 'RequirementLevel']:
                                if rl_col in row.index:
                                    req_level = str(row[rl_col]).strip()
                                    break
                            
                            # Store both original and lowercase versions for accurate lookup
                            req_info[attribute.lower()] = {
                                'Product Type': product_type,
                                'Requirement Level': req_level,
                                'Original Attribute': attribute
                            }
                            # Log the stored values
                            self.log(f"Stored - Attribute: {attribute}, "
                                    f"Product Type: {product_type}, "
                                    f"Requirement Level: {req_level}")
                            break
            
            self.log(f"\nLoaded {len(req_info)} attributes from Requirement Level sheet")
            if len(req_info) == 0:
                self.log("WARNING: No attributes were loaded from Requirement Level sheet!")
            
            return req_info
        except Exception as e:
            self.log(f"Error reading Requirement level sheet: {str(e)}")
            import traceback
            self.log(traceback.format_exc())
            return {}

    def normalize_numeric_values(self, value):
        """Normalize numeric values to avoid float vs int mismatches"""
        if pd.isna(value):
            return value
        try:
            # Handle numeric types first
            if isinstance(value, (int, float)):
                # Convert scientific notation to regular number
                if abs(value) >= 1e4:
                    str_val = f"{value:.0f}"  # Convert to regular number without decimals
                else:
                    str_val = str(float(value))  # Convert to string with possible decimals
                
                # Remove .0 if it's a whole number
                if str_val.endswith('.0'):
                    return str_val[:-2]
                return str_val
            
            # If it's already a string, clean it up
            str_val = str(value).strip()
            
            # Try to convert string to float and back to handle scientific notation in strings
            try:
                num_val = float(str_val)
                if abs(num_val) >= 1e4:
                    str_val = f"{num_val:.0f}"
                else:
                    str_val = str(num_val)
                if str_val.endswith('.0'):
                    return str_val[:-2]
            except:
                pass
            
            return str_val
        except:
            return str(value)

    def qc_comparison_script(self, input_key_path, input_agent_path, output_file_path, requirement_level_info):
        """Original QC comparison function with GUI logging"""
        self.log("Starting QC Comparison Process...")
        self.log("=" * 50)
        
        # Resolve Excel files from inputs
        try:
            keysheet_file = self.get_excel_file_from_input(input_key_path)
            agent_file = self.get_excel_file_from_input(input_agent_path)
            
            self.log(f"Keysheet file: {keysheet_file}")
            self.log(f"Agent file: {agent_file}")
            
        except Exception as e:
            self.log(f"Error resolving input files: {e}")
            return

        def normalize_numeric_values(self, value):
            """Normalize numeric values to avoid float vs int mismatches"""
            if pd.isna(value):
                return value
            try:
                # Convert to string first to handle both numeric types
                str_val = str(value).strip()
                # Remove .0 from float strings
                if str_val.endswith('.0'):
                    return str_val[:-2]
                return str_val
            except:
                return str(value)

        # Read the Production & QC Report sheets from both files
        try:
            keysheet_df = pd.read_excel(keysheet_file, sheet_name='Production & QC Report')
            requirement_df = pd.read_excel(keysheet_file, sheet_name='Requirement Level')
            agent_df = pd.read_excel(agent_file, sheet_name='Production & QC Report')
            
            # Normalize numeric values in both dataframes and log some examples
            for df_name, df in [("Keysheet", keysheet_df), ("Agent", agent_df)]:
                for col in df.columns:
                    original_values = df[col].head()
                    df[col] = df[col].apply(self.normalize_numeric_values)
                    normalized_values = df[col].head()
                    
                    # Log some example conversions for numeric-looking columns
                    if any(str(x).replace('.', '').isdigit() for x in original_values if pd.notna(x)):
                        self.log(f"\nNormalization examples for {df_name} - {col}:")
                        for orig, norm in zip(original_values, normalized_values):
                            if pd.notna(orig):
                                self.log(f"Original: {orig} ({type(orig)}) -> Normalized: {norm} ({type(norm)})")
            
            self.log(f"Successfully loaded keysheet with {len(keysheet_df)} rows and agent file with {len(agent_df)} rows")
        except Exception as e:
            self.log(f"Error reading Excel files: {e}")
            return
        
        # Function to identify columns to exclude
        def is_validation_or_error_column(col):
            return ('Validation Comment' in col) or ('Error Code' in col)
        
        # Columns to exclude from comparison
        exclude_columns = [
            'Date (YYYY/MM/DD)', 'Work Type', 'Associate Walmart ID', 'Item Status',
            'Submission ID', 'Product ID Type', 'Product ID', 'Audit Template Version',
            'SOP Reference', 'Initial CQ Score', 'Current CQ Score', 'Supplier ID',
            'Item Created Date', 'Active Status', 'is Private Label?', 'Website Link',
            'View Images'
        ]
        
        # Add validation comment and error code columns to exclude list
        exclude_columns.extend([col for col in keysheet_df.columns if is_validation_or_error_column(col)])
        
        # Special handling for Product Type and Product Name
        product_type_columns = [
            'Correct Product Type', 'Is Product Type Validated', 
            'Product Error Code', 'Product Type Validation Comment'
        ]
        product_name_columns = [
            'Correct Product Name', 'Is Product Name Validated',
            'Product Name Error Code', 'Product Name Validation Comment'
        ]
        
        # Merge dataframes on Item ID
        merged_df = pd.merge(keysheet_df, agent_df, on='Item ID', how='inner', suffixes=('_key', '_agent'))
        self.log(f"Successfully matched {len(merged_df)} records on Item ID")
        
        # Initialize lists to store discrepancy records
        discrepancy_records = []
        
        # Get all columns from keysheet for comparison
        keysheet_columns = keysheet_df.columns.tolist()
        
        # Process each matched record
        for index, row in merged_df.iterrows():
            item_id = row['Item ID']
            associate_id = row.get('Associate Walmart ID_agent', row.get('Associate Walmart ID_key', ''))
            
            for col in keysheet_columns:
                # Skip excluded columns
                if col in exclude_columns or col == 'Item ID':
                    continue
                    
                # Skip columns with 'Pre - ' prefix in keysheet
                if col.startswith('Pre - '):
                    continue
                    
                # Special handling for Product Type and Product Name
                if 'Product Type' in col and col not in product_type_columns:
                    continue
                if 'Product Name' in col and col not in product_name_columns:
                    continue
                    
                # Get corresponding agent column
                agent_col = col + '_agent'
                key_col = col + '_key'
                
                if agent_col in merged_df.columns:
                    key_value = row.get(key_col, '')
                    agent_value = row.get(agent_col, '')
                    
                    # Convert to string for comparison and handle NaN values
                    key_value_str = str(key_value) if pd.notna(key_value) else ''
                    agent_value_str = str(agent_value) if pd.notna(agent_value) else ''
                    
                    # Check for discrepancy
                    if key_value_str != agent_value_str:
                        # Determine if this is an error column or attribute column
                        attribute_name = col
                        error_column = col
                        
                        import re
                        
                        # Extract the base attribute name and set error column
                        error_column = col
                        
                        # Define regex patterns for different column types
                        validation_patterns = [
                            (r'^Is\s+(.+?)\s+Validated?.*$', r'\1')  # Is X Validated or Is X Validate?
                        ]
                        
                        # Try each pattern
                        attribute_name = col
                        for pattern, replacement in validation_patterns:
                            match = re.match(pattern, col, re.IGNORECASE)
                            if match:
                                attribute_name = re.sub(pattern, replacement, col, flags=re.IGNORECASE)
                                self.log(f"Column: {col}")
                                self.log(f"Matched pattern: {pattern}")
                                self.log(f"Extracted attribute: {attribute_name}")
                                break
                        
                        # Remove any extra spaces
                        attribute_name = attribute_name.strip()
                            
                        # Log for debugging
                        self.log(f"Processing column: {col}")
                        self.log(f"Extracted attribute name: {attribute_name}")
                            
                        # Log the attribute name extraction
                        self.log(f"Column: {col} -> Extracted attribute name: {attribute_name}")
                        
                        # Get Product Type and Requirement Level for the attribute
                        import re
                        
                        # Define regex patterns for different column types
                        validation_patterns = [
                            (r'^Is\s+(.+?)\s+Validated?.*$', r'\1')  # Is X Validated or Is X Validate?
                        ]
                        
                        # Process the attribute name
                        base_attribute = attribute_name.strip()
                        for pattern, replacement in validation_patterns:
                            match = re.match(pattern, base_attribute, re.IGNORECASE)
                            if match:
                                base_attribute = re.sub(pattern, replacement, base_attribute, flags=re.IGNORECASE)
                                self.log(f"Base attribute processing:")
                                self.log(f"Original: {attribute_name}")
                                self.log(f"After regex: {base_attribute}")
                                break
                        
                        # Remove any extra spaces
                        base_attribute = base_attribute.strip()
                            
                        # Try to find the attribute in requirement_level_info using case-insensitive matching
                        attribute_key = base_attribute.lower()
                        req_info = requirement_level_info.get(attribute_key, {})
                        
                        product_type = req_info.get('Product Type', '')
                        requirement_level = req_info.get('Requirement Level', '')
                        
                        # Detailed logging for debugging
                        self.log(f"\nProcessing error record:")
                        self.log(f"Original attribute name: {attribute_name}")
                        self.log(f"Cleaned attribute name: {base_attribute}")
                        self.log(f"Lookup key: {attribute_key}")
                        self.log(f"Found in requirement info: {bool(req_info)}")
                        if req_info:
                            self.log(f"Found values - Product Type: {product_type}, Requirement Level: {requirement_level}")
                        elif base_attribute not in ['Item ID', 'Associate Walmart ID']:
                            self.log(f"No requirement level info found for attribute: {base_attribute}")
                        
                        discrepancy_records.append({
                            'Item ID': item_id,
                            'Associate Walmart ID': associate_id,
                            'Attribute ': attribute_name,  # Note the space after 'Attribute' to match sample
                            'Error column': error_column,
                            'Key sheet Value': key_value_str,
                            'Agent Value': agent_value_str,
                            'Product Type': product_type,
                            'Requirement Level': requirement_level
                        })
        
        # Create Error sheet DataFrame
        error_df = pd.DataFrame(discrepancy_records)
        self.log(f"Found {len(error_df)} discrepancies")
        
        # Create Agent level error sheet
        if len(error_df) > 0:
            agent_stats = []
            
            # Get unique agents from merged data
            unique_agents = merged_df['Associate Walmart ID_agent'].dropna().unique()
            
            for agent in unique_agents:
                # Get agent's work from merged data
                agent_work = merged_df[merged_df['Associate Walmart ID_agent'] == agent]
                total_items = len(agent_work)
                
                # Count total attributes worked (approximate based on comparable columns)
                comparable_columns = [col for col in keysheet_columns 
                                    if col not in exclude_columns 
                                    and col != 'Item ID'
                                    and not col.startswith('Pre - ')
                                    and (('Product Type' not in col) or (col in product_type_columns))
                                    and (('Product Name' not in col) or (col in product_name_columns))]
                
                total_attributes = total_items * len(comparable_columns)
                
                # Count errors for this agent
                agent_errors = error_df[error_df['Associate Walmart ID'] == agent]
                attribute_errors = len(agent_errors)
                
                # Calculate error percentage
                error_percentage = attribute_errors / total_attributes if total_attributes > 0 else 0
                
                # Get agent's data and errors
                agent_data = agent_df[agent_df['Associate Walmart ID'] == agent]
                agent_errors_df = error_df[error_df['Associate Walmart ID'] == agent]
                
                # Count failed line items (unique Item IDs in error_df for this agent)
                failed_line_items = len(agent_errors_df['Item ID'].unique())
                
                self.log(f"\nAgent {agent} statistics:")
                self.log(f"Total items worked: {len(agent_data)}")
                self.log(f"Failed line items (unique Item IDs with errors): {failed_line_items}")
            
            possible_names = [
                'Is Product Type Validated', 
                'Product Type Validation', 
                'Product Type Validation Comment', 
                'Product Type Validated'
            ]
            for col_name in possible_names:
                if col_name in agent_data.columns:
                    validation_col = col_name
                    self.log(f"Found Product Type validation column: {col_name}")
                    break
            
            if validation_col is None:
                self.log("WARNING: Could not find Product Type validation column. Available columns:")
                self.log("\n".join(agent_data.columns.tolist()))
                failed_line_items = 0
            else:
                try:
                    # Convert the validation column to string and handle NaN values
                    agent_data[validation_col] = agent_data[validation_col].fillna('').astype(str)
                    
                    # Get validation-related columns
                    validation_columns = agent_data.columns[agent_data.columns.str.contains('Validation|Validated', case=False)]
                    self.log(f"Validation columns found: {', '.join(validation_columns)}")
                    
                    # Count failed items (unique Item IDs in error_df for this agent)
                    agent_errors_df = error_df[error_df['Associate Walmart ID'] == agent]
                    failed_line_items = len(agent_errors_df['Item ID'].unique())
                    self.log(f"Failed line items for agent {agent}: {failed_line_items} (unique Item IDs in error sheet)")
                    
                    self.log(f"Number of failed line items found: {failed_line_items}")
                except Exception as e:
                    self.log(f"Error calculating failed line items: {str(e)}")
                    self.log("Using 0 as default value for failed line items")
                    failed_line_items = 0
            
            # Calculate line items audited and passed
            line_items_audited = len(agent_data)
            line_items_passed = line_items_audited - failed_line_items
            
            # Get counts of Recommended and Required attributes from requirement level info
            recommended_count = sum(1 for info in requirement_level_info.values() 
                                 if info['Requirement Level'].lower() == 'recommended')
            required_count = sum(1 for info in requirement_level_info.values() 
                               if info['Requirement Level'].lower() == 'required')
            
            # Count Recommended and Required attribute errors from error_df
            agent_errors_df = error_df[error_df['Associate Walmart ID'] == agent]
            recommended_errors = len(agent_errors_df[
                agent_errors_df['Requirement Level'].str.lower() == 'recommended'
            ]) if len(agent_errors_df) > 0 else 0
            required_errors = len(agent_errors_df[
                agent_errors_df['Requirement Level'].str.lower() == 'required'
            ]) if len(agent_errors_df) > 0 else 0
            
            # Get the count of total attributes (excluding system columns)
            total_attr_count = len([col for col in agent_data.columns 
                                  if col not in ['Item ID', 'Associate Walmart ID', 'Date (YYYY/MM/DD)',
                                               'Work Type', 'Item Status', 'Submission ID', 'Product ID Type',
                                               'Product ID', 'Audit Template Version', 'SOP Reference',
                                               'Initial CQ Score', 'Current CQ Score', 'Supplier ID',
                                               'Item Created Date', 'Active Status', 'is Private Label?',
                                               'Website Link', 'View Images']])
            
            # Calculate all metrics
            total_attributes = line_items_audited * (recommended_count + required_count) * 2
            recommended_total_attrs = line_items_audited * recommended_count * 2
            required_total_attrs = line_items_audited * required_count * 2
            total_errors = recommended_errors + required_errors
            
            self.log(f"Metrics calculation:")
            self.log(f"Total attributes count: {total_attr_count}")
            self.log(f"Total attributes: {total_attributes}")
            self.log(f"Recommended total attributes: {recommended_total_attrs}")
            self.log(f"Required total attributes: {required_total_attrs}")
            self.log(f"Total errors: {total_errors}")
            
            # Calculate percentages
            item_level_qc = 1 - (failed_line_items / line_items_audited) if line_items_audited > 0 else 0
            recommended_accuracy = 1 - (recommended_errors / recommended_total_attrs) if recommended_total_attrs > 0 else 0
            required_accuracy = 1 - (required_errors / required_total_attrs) if required_total_attrs > 0 else 0
            overall_accuracy = (recommended_accuracy + required_accuracy) / 2
            total_rec_req_attrs = recommended_count + required_count
            
            agent_stats.append({
                'Associate Walmart ID': agent,
                'Line item Audited': line_items_audited,
                'No. of line items Failed': failed_line_items,
                'No of line items Passed': line_items_passed,
                'Total Attributes': total_attributes,
                'Recommended Count': recommended_count,
                'Recommended Attributes Failed': recommended_errors,
                'Required Count': required_count,
                'Required Attributes Failed': required_errors,
                '(Recommended) Total Attributes': recommended_total_attrs,
                '(Recommended) No of Attributes Failed': recommended_errors,
                '(Required) Total Attributes': required_total_attrs,
                '(Required) No of Attributes Failed': required_errors,
                'Total No of Attributes Failed': total_errors,
                'Item Level QC %': f"{item_level_qc * 100:.2f}%",
                'Recommended Attribute level Accuracy %': f"{recommended_accuracy * 100:.2f}%",
                'Required Attribute level Accuracy %': f"{required_accuracy * 100:.2f}%",
                'Overall %': f"{overall_accuracy * 100:.2f}%",
                'Total Recommended and Required Attributes': total_rec_req_attrs
            })
            
            # Add logging for verification
            self.log(f"\nAgent Level Statistics for {agent}:")
            self.log(f"Line items Audited: {line_items_audited}")
            self.log(f"Failed Line Items: {failed_line_items}")
            self.log(f"Passed Line Items: {line_items_passed}")
            self.log(f"Recommended Count: {recommended_count}")
            self.log(f"Required Count: {required_count}")
            self.log(f"Recommended Errors: {recommended_errors}")
            self.log(f"Required Errors: {required_errors}")
            self.log(f"Item Level QC %: {item_level_qc * 100:.2f}%")
            self.log(f"Overall Accuracy: {overall_accuracy * 100:.2f}%")
            
            agent_level_df = pd.DataFrame(agent_stats)
        else:
            # Create empty agent level dataframe if no errors found
            agent_level_df = pd.DataFrame(columns=[
                'Associate Walmart ID', 'Line item Audited', 'No. of line items Failed',
                'No of line items Passed', 'Total Attributes', 
                'Recommended Count', 'Recommended Attributes Failed',
                'Required Count', 'Required Attributes Failed',
                '(Recommended) Total Attributes', '(Recommended) No of Attributes Failed',
                '(Required) Total Attributes', '(Required) No of Attributes Failed',
                'Total No of Attributes Failed', 'Item Level QC %',
                'Recommended Attribute level Accuracy %', 'Required Attribute level Accuracy %',
                'Overall %', 'Total Recommended and Required Attributes'
            ])
        
        # Write to Excel file with multiple sheets
        try:
            # Create output directory if it doesn't exist
            output_dir = os.path.dirname(output_file_path)
            if output_dir and not os.path.exists(output_dir):
                os.makedirs(output_dir)
                
            # Create formulas DataFrame
            formulas_data = {
                'Column Name': [
                    'No of line items Passed',
                    'Total Attributes',
                    '(Recommended) Total Attributes',
                    '(Required) Total Attributes',
                    'Total No of Attributes Failed',
                    'Item Level QC %',
                    'Recommended Attribute level Accuracy %',
                    'Required Attribute level Accuracy %',
                    'Overall %',
                    'Total Recommended and Required Attributes'
                ],
                'Formula': [
                    'Line items Audited - No of line items Failed',
                    'Line items Audited * (Recommended count + Required count) * 2',
                    'Line items Audited * Recommended Count * 2',
                    'Line items Audited * Required Count * 2',
                    '(Recommended) No of Attributes Failed + (Required)No of Attributes Failed',
                    '1 - (No of line items Failed)/(Line items Audited)',
                    '1 - ((Recommended) No of Attributes Failed)/((Recommended) Total Attributes)',
                    '1 - ((Required) No of Attributes Failed)/((Required) Total Attributes)',
                    'Average of Recommended Attribute level Accuracy % and Required Attribute level Accuracy %',
                    'Recommended Count + Required Count'
                ],
                'Description': [
                    'Number of line items that passed QC',
                    'Total number of attributes across all line items',
                    'Total number of recommended attributes across all line items',
                    'Total number of required attributes across all line items',
                    'Total number of failed attributes (both recommended and required)',
                    'Percentage of line items that passed QC',
                    'Accuracy percentage for recommended attributes',
                    'Accuracy percentage for required attributes',
                    'Average of recommended and required accuracy percentages',
                    'Total count of all recommended and required attributes'
                ]
            }
            formulas_df = pd.DataFrame(formulas_data)
            
            # Write all sheets
            with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
                error_df.to_excel(writer, sheet_name='Error sheet', index=False)
                agent_level_df.to_excel(writer, sheet_name='Agent level error summary', index=False)
                formulas_df.to_excel(writer, sheet_name='Formulas', index=False)
            
            self.log(f"\nQC Report Generation Completed Successfully!")
            self.log(f"Output file saved at: {output_file_path}")
            self.log(f"Error sheet contains {len(error_df)} discrepancy records")
            self.log(f"Agent level error summary sheet contains {len(agent_level_df)} agent records")
            
            # Show success message
            self.root.after(0, lambda: messagebox.showinfo("Success", 
                f"QC Report generated successfully!\n\nOutput saved to:\n{output_file_path}\n\nDiscrepancies found: {len(error_df)}\nAgents analyzed: {len(agent_level_df)}"))
            
        except Exception as e:
            self.log(f"Error writing output file: {e}")


def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    app = ModernQCComparisonGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
