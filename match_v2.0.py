import pandas as pd
from rapidfuzz import fuzz
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Text, StringVar, Frame, Scrollbar
import os
from datetime import datetime

class FuzzyMatcherGUI:
    def __init__(self, master):
        self.master = master
        master.title("Fuzzy Matcher Dinamis")
        
        # Configurable settings
        self.settings = {
            'output_filename': f"matches_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            'clean_replacements': {'.': '', ',': '', '-': ' '},
            'scoring_methods': {
                'exact': lambda x, y: 100.0 if x == y else 0.0,
                'fuzzy': [
                    fuzz.token_sort_ratio,
                    fuzz.token_set_ratio,
                    fuzz.partial_ratio
                ],
                'numeric': lambda x, y: 100.0 if str(x).strip() == str(y).strip() else 0.0,
                'date': lambda x, y: 100.0 if str(x).strip() == str(y).strip() else 0.0
            },
            'default_method': 'fuzzy',
            'field_methods': {
                'id': 'exact',
                'kode': 'exact',
                'nomor': 'exact',
                'number': 'exact',
                'date': 'date',
                'tanggal': 'date',
                'numeric': 'numeric'
            },
            'score_precision': 1,
            'window_geometry': "1000x600",  # Reduced height from 750 to 600
            'default_encoding': 'utf-8',
            'max_log_lines': 1000,
            'fonts': {
                'default': ('Segoe UI', 9),  # Windows default UI font
                'header': ('Segoe UI', 10, 'bold'),
                'monospace': ('Consolas', 10),  # For log output
                'alternative1': ('Calibri', 9),  # Clean, modern font
                'alternative2': ('Tahoma', 9),  # Compact, readable font
                'alternative3': ('Verdana', 9),  # Highly readable font
            }
        }
        
        master.geometry(self.settings['window_geometry'])
        
        # Apply a theme for better appearance
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure styles
        style.configure('TFrame', background='#f0f0f0')
        style.configure('TButton', font=self.settings['fonts']['default'])
        style.configure('Header.TLabel', font=self.settings['fonts']['header'], background='#f0f0f0')
        style.configure('TLabel', font=self.settings['fonts']['default'], background='#f0f0f0')
        style.configure('TCombobox', font=self.settings['fonts']['default'])
        style.configure('TCheckbutton', font=self.settings['fonts']['default'])
        style.configure('Remove.TButton', padding=0, font=('Arial', 8))
        
        # Configure LabelFrame style for section titles
        style.configure('TLabelframe', background='#f0f0f0')  # Original light gray background
        style.configure('TLabelframe.Label', 
                       font=('Arial', 10),  # Regular font, not bold
                       background='#f0f0f0',  # Original light gray background
                       foreground='black')  # Black text
        
        # Main container with padding
        main_frame = ttk.Frame(master, padding="10 8 10 8")
        main_frame.pack(fill="both", expand=True)
        
        # Create a horizontal split frame
        split_frame = ttk.Frame(main_frame)
        split_frame.pack(fill="both", expand=True)
        
        # Left frame for controls
        left_frame = ttk.Frame(split_frame)
        left_frame.pack(side="left", fill="both", expand=False, padx=(0, 10))
        left_frame.configure(width=400)  # Set fixed width after packing
        
        # ===== File Selection Section =====
        file_frame = ttk.LabelFrame(left_frame, text="File Selection", padding="6 2 6 2")
        file_frame.pack(fill="x", pady=(0, 10))
        
        # File 1 selection
        file1_frame = ttk.Frame(file_frame)
        file1_frame.pack(fill="x", pady=2)  # Reduced pady from 5 to 2
        ttk.Label(file1_frame, text="File 1:", width=8).pack(side="left", padx=(0, 5))  # Reduced width and padding
        self.file1_label = ttk.Label(file1_frame, text="No file selected", width=35)  # Reduced width from 50
        self.file1_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file1_frame, text="Browse", width=8, command=self.load_file1).pack(side="right")  # Added fixed width
        
        # File 2 selection
        file2_frame = ttk.Frame(file_frame)
        file2_frame.pack(fill="x", pady=2)  # Reduced pady from 5 to 2
        ttk.Label(file2_frame, text="File 2:", width=8).pack(side="left", padx=(0, 5))  # Reduced width and padding
        self.file2_label = ttk.Label(file2_frame, text="No file selected", width=35)  # Reduced width from 50
        self.file2_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file2_frame, text="Browse", width=8, command=self.load_file2).pack(side="right")  # Added fixed width
        
        # ===== Configuration Section =====
        config_frame = ttk.LabelFrame(left_frame, text="Configuration", padding="6 2 6 2")
        config_frame.pack(fill="x", pady=(0, 10))
        
        # IDSBR selector (only for File 1)
        idsbr_frame = ttk.Frame(config_frame)
        idsbr_frame.pack(fill="x", pady=5)
        ttk.Label(idsbr_frame, text="ID Column (File 1):", width=20).pack(side="left")
        self.idsbr_var = StringVar(master)
        self.idsbr_dropdown = ttk.Combobox(idsbr_frame, textvariable=self.idsbr_var, state="readonly")
        self.idsbr_dropdown.pack(side="left", fill="x", expand=True)
        
        # ===== Column Mapping Section =====
        mapping_frame_outer = ttk.LabelFrame(left_frame, text="Column Mapping (Include Region Columns Here)", padding="6 2 6 6")
        mapping_frame_outer.pack(fill="both", expand=True, pady=(0, 10))
        
        # Add headers for columns
        headers_frame = ttk.Frame(mapping_frame_outer)
        headers_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(headers_frame, text="File 1 Column", width=25).pack(side="left", padx=(10, 5))
        ttk.Label(headers_frame, text="").pack(side="left", padx=5, pady=5)
        ttk.Label(headers_frame, text="File 2 Column", width=25).pack(side="left", padx=5)
        ttk.Button(headers_frame, text="➕ Add Mapping", command=self.add_mapping_pair).pack(side="right", padx=5)
        
        # Create scrollable frame for mappings
        mapping_canvas_frame = ttk.Frame(mapping_frame_outer)
        mapping_canvas_frame.pack(fill="both", expand=True)
        
        self.mapping_canvas = tk.Canvas(mapping_canvas_frame, height=80, background="#f0f0f0", highlightthickness=0)  # Reduced from 120 to 80
        scrollbar = ttk.Scrollbar(mapping_canvas_frame, orient="vertical", command=self.mapping_canvas.yview)
        
        self.mapping_inner_frame = ttk.Frame(self.mapping_canvas)
        
        self.mapping_inner_frame.bind(
            "<Configure>",
            lambda e: self.mapping_canvas.configure(scrollregion=self.mapping_canvas.bbox("all"))
        )
        
        self.mapping_canvas.create_window((0, 0), window=self.mapping_inner_frame, anchor="nw")
        self.mapping_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.mapping_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # ===== Additional Output Columns Section =====
        additional_cols_frame = ttk.LabelFrame(left_frame, text="Additional Output Columns", padding="6 2 6 6")
        additional_cols_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Headers for additional columns
        add_cols_headers_frame = ttk.Frame(additional_cols_frame)
        add_cols_headers_frame.pack(fill="x", pady=(0, 5))
        ttk.Label(add_cols_headers_frame, text="Source", width=8).pack(side="left", padx=(10, 5))
        ttk.Label(add_cols_headers_frame, text="Column", width=25).pack(side="left", padx=(0, 5))
        ttk.Button(add_cols_headers_frame, text="➕ Add Column", command=self.add_output_column).pack(side="right", padx=5)
        
        # Create scrollable frame for additional columns
        add_cols_canvas_frame = ttk.Frame(additional_cols_frame)
        add_cols_canvas_frame.pack(fill="both", expand=True)
        
        self.add_cols_canvas = tk.Canvas(add_cols_canvas_frame, height=60, background="#f0f0f0", highlightthickness=0)
        add_cols_scrollbar = ttk.Scrollbar(add_cols_canvas_frame, orient="vertical", command=self.add_cols_canvas.yview)
        
        self.add_cols_inner_frame = ttk.Frame(self.add_cols_canvas)
        
        self.add_cols_inner_frame.bind(
            "<Configure>",
            lambda e: self.add_cols_canvas.configure(scrollregion=self.add_cols_canvas.bbox("all"))
        )
        
        self.add_cols_canvas.create_window((0, 0), window=self.add_cols_inner_frame, anchor="nw")
        self.add_cols_canvas.configure(yscrollcommand=add_cols_scrollbar.set)
        
        self.add_cols_canvas.pack(side="left", fill="both", expand=True)
        add_cols_scrollbar.pack(side="right", fill="y")
        
        # Output file selection
        output_frame = ttk.Frame(left_frame)
        output_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(output_frame, text="Output:", width=8).pack(side="left", padx=(0, 5))
        self.output_file_entry = ttk.Entry(output_frame)
        self.output_file_entry.pack(side="left", fill="x", expand=True, padx=5)
        self.output_file_entry.insert(0, self.settings['output_filename'])
        ttk.Button(output_frame, text="Browse", width=8, command=self.select_output_file).pack(side="right")
        
        # Format settings for output
        format_frame = ttk.Frame(left_frame)
        format_frame.pack(fill="x", pady=(0, 10))
        
        # ===== Log Output Section =====
        log_section = ttk.Frame(split_frame)
        log_section.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Run Matching button at the top
        run_button_frame = ttk.Frame(log_section)
        run_button_frame.pack(fill="x", pady=(0, 5))
        
        run_button = tk.Button(
            run_button_frame, 
            text="🚀 RUN MATCHING",  # All caps
            command=self.run_match_thread,
            bg='#4CAF50',  # Bright green background
            fg='white',    # White text
            font=('Arial', 9, 'bold'),  # Bold font
            relief='raised',
            bd=2
        )
        run_button.pack(fill="x", expand=True)
        
        # Log frame below the button
        log_frame = ttk.LabelFrame(log_section, text="Process Log", padding="10 5 10 10")
        log_frame.pack(fill="both", expand=True)
        
        # Text widget with scrollbar for log output
        log_container = ttk.Frame(log_frame)
        log_container.pack(fill="both", expand=True)
        
        self.text_output = Text(log_container, height=6, wrap="word",
                              font=self.settings['fonts']['monospace'])
        log_scrollbar = ttk.Scrollbar(log_container, orient="vertical", command=self.text_output.yview)
        self.text_output.configure(yscrollcommand=log_scrollbar.set)
        
        self.text_output.pack(side="left", fill="both", expand=True)
        log_scrollbar.pack(side="right", fill="y")
        
        # Initialize variables
        self.file1 = None
        self.file2 = None
        self.df1 = None
        self.df2 = None
        self.df1_cols = []
        self.df2_cols = []
        self.mapping_pairs = []
        self.additional_output_cols = []  # New: stores additional output column selections
        self.running = False

    def clean_text(self, text):
        if pd.isna(text): return ""
        # Convert to string and strip whitespace
        text = str(text).strip()
        # If it's a number, preserve it exactly
        if text.replace('.', '', 1).isdigit():
            return text
        # Otherwise apply normal cleaning
        text = text.lower()
        for old, new in self.settings['clean_replacements'].items():
            text = text.replace(old, new)
        return ' '.join(text.split())

    def get_field_matching_method(self, field_name):
        field_lower = str(field_name).lower()
        for key, method in self.settings['field_methods'].items():
            if key in field_lower:
                return method
        return self.settings['default_method']

    def get_match_score(self, val1, val2, field_name):
        # If values are exactly the same (including case), return 100
        if val1 == val2:
            return 100.0
            
        # If either value is empty after cleaning, return 0
        if not val1 or not val2:
            return 0.0
            
        # Get the appropriate matching method for this field
        method_type = self.get_field_matching_method(field_name)
        
        if method_type == 'exact':
            return self.settings['scoring_methods']['exact'](val1, val2)
        elif method_type == 'numeric':
            return self.settings['scoring_methods']['numeric'](val1, val2)
        elif method_type == 'date':
            return self.settings['scoring_methods']['date'](val1, val2)
        else:  # fuzzy
            scores = [
                method(val1, val2) 
                for method in self.settings['scoring_methods']['fuzzy']
            ]
            return sum(scores) / len(scores)

    def deduplicate_columns(self, df):
        new_columns = []
        seen = {}
        
        for col in df.columns:
            col_str = str(col).strip()
            if col_str in seen:
                seen[col_str] += 1
                new_columns.append(f"{col_str}_{seen[col_str]}")
            else:
                seen[col_str] = 0
                new_columns.append(col_str)
        
        return pd.DataFrame(df.values, columns=new_columns)

    def select_output_file(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile=self.settings['output_filename'])
        if filename:
            self.output_file_entry.delete(0, tk.END)
            self.output_file_entry.insert(0, filename)

    def log(self, msg):
        self.text_output.insert("end", msg + "\n")
        self.text_output.see("end")
        
        # Limit log size
        lines = int(self.text_output.index('end').split('.')[0])
        if lines > self.settings['max_log_lines']:
            self.text_output.delete(1.0, f"{lines-self.settings['max_log_lines']}.0")
            
        self.master.update_idletasks()

    def load_file1(self):
        self.file1 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not self.file1: return
        self.file1_label.config(text=os.path.basename(self.file1))
        self.log(f"✅ File 1 loaded: {self.file1}")
        
        try:
            # Read all columns as strings to preserve leading zeros
            temp_df = pd.read_excel(self.file1, dtype=str)
            self.df1 = self.deduplicate_columns(temp_df)
            
            if len(temp_df.columns) != len(set(temp_df.columns)):
                self.log("Note: Duplicate column names in File 1 have been renamed")
            
            self.df1_cols = list(self.df1.columns)
            self.idsbr_dropdown['values'] = self.df1_cols
            
            if self.df1_cols:
                self.idsbr_var.set(self.df1_cols[0])
                
        except Exception as e:
            self.log(f"❌ Failed to read File 1: {str(e)}")
            self.file1 = None
            self.df1 = None

    def load_file2(self):
        self.file2 = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if not self.file2: return
        self.file2_label.config(text=os.path.basename(self.file2))
        self.log(f"✅ File 2 loaded: {self.file2}")
        
        try:
            # Read all columns as strings to preserve leading zeros
            temp_df = pd.read_excel(self.file2, dtype=str)
            self.df2 = self.deduplicate_columns(temp_df)
            
            if len(temp_df.columns) != len(set(temp_df.columns)):
                self.log("Note: Duplicate column names in File 2 have been renamed")
            
            self.df2_cols = list(self.df2.columns)
                
        except Exception as e:
            self.log(f"❌ Failed to read File 2: {str(e)}")
            self.file2 = None
            self.df2 = None

    def add_mapping_pair(self):
        if not self.df1_cols or not self.df2_cols:
            messagebox.showwarning("Warning", "Please load both Excel files first")
            return
            
        row = ttk.Frame(self.mapping_inner_frame)
        row.pack(fill="x", pady=2)
        
        var1 = StringVar(self.master)
        var2 = StringVar(self.master)
        var1.set(self.df1_cols[0] if self.df1_cols else "")
        var2.set(self.df2_cols[0] if self.df2_cols else "")
        
        col1_dropdown = ttk.Combobox(row, textvariable=var1, values=self.df1_cols, width=25, state="readonly")
        col1_dropdown.pack(side="left", padx=(10, 5))
        
        ttk.Label(row, text="⟷").pack(side="left", padx=5)
        
        col2_dropdown = ttk.Combobox(row, textvariable=var2, values=self.df2_cols, width=25, state="readonly")
        col2_dropdown.pack(side="left", padx=5)
        
        def remove_row():
            self.mapping_pairs.remove((var1, var2))
            row.destroy()
        
        # Create a properly sized remove button
        remove_btn = ttk.Button(row, text="✖", width=2, command=remove_row)
        remove_btn.pack(side="right", padx=5)
        
        # Configure the button to have proper height
        style = ttk.Style()
        style.configure('Remove.TButton', padding=0, font=('Arial', 8))  # Smaller font size
        remove_btn.configure(style='Remove.TButton')
        
        self.mapping_pairs.append((var1, var2))

    def add_output_column(self):
        if not self.df1_cols and not self.df2_cols:
            messagebox.showwarning("Warning", "Please load at least one Excel file first")
            return
            
        row = ttk.Frame(self.add_cols_inner_frame)
        row.pack(fill="x", pady=2)  # Reduced pady from 5 to 2
        
        source_var = StringVar(self.master)
        source_var.set("File 1")  # Default value
        
        source_dropdown = ttk.Combobox(row, textvariable=source_var, values=["File 1", "File 2"], width=8, state="readonly")
        source_dropdown.pack(side="left", padx=(10, 5))
        
        column_var = StringVar(self.master)
        
        # Initial column values based on selected source
        columns = self.df1_cols if source_var.get() == "File 1" else self.df2_cols
        if columns:
            column_var.set(columns[0])
        
        column_dropdown = ttk.Combobox(row, textvariable=column_var, values=columns, width=25, state="readonly")
        column_dropdown.pack(side="left", padx=5)
        
        # Update column dropdown when source changes
        def update_columns(*args):
            new_columns = self.df1_cols if source_var.get() == "File 1" else self.df2_cols
            column_dropdown['values'] = new_columns
            if new_columns:
                column_var.set(new_columns[0])
                
        source_var.trace("w", update_columns)
        
        def remove_row():
            self.additional_output_cols.remove((source_var, column_var))
            row.destroy()
        
        # Create a properly sized remove button
        remove_btn = ttk.Button(row, text="✖", width=2, command=remove_row)
        remove_btn.pack(side="right", padx=5)
        
        # Configure the button to have proper height
        style = ttk.Style()
        style.configure('Remove.TButton', padding=0, font=('Arial', 8))  # Smaller font size
        remove_btn.configure(style='Remove.TButton')
        
        self.additional_output_cols.append((source_var, column_var))

    def run_match_thread(self):
        if self.running:
            messagebox.showwarning("Warning", "A matching process is already running")
            return
            
        if not self.file1 or not self.file2:
            messagebox.showwarning("Warning", "Please select both Excel files first")
            return
            
        if not self.mapping_pairs:
            messagebox.showwarning("Warning", "Please add at least one column mapping")
            return
            
        self.running = True
        threading.Thread(target=self.run_match, daemon=True).start()

    def run_match(self):
        try:
            self.log("Starting matching process...")
            
            if self.df1 is None or self.df2 is None:
                self.log("❌ Data not loaded properly. Please reload files.")
                return

            idsbr_col = self.idsbr_var.get()
            mapping = [(v1.get(), v2.get()) for v1, v2 in self.mapping_pairs if v1.get() and v2.get()]
            
            # Get additional output columns
            add_cols = [(src.get(), col.get()) for src, col in self.additional_output_cols if col.get()]
            
            if not mapping:
                self.log("❌ No valid column mappings selected")
                return
            
            # Find region columns (if any were mapped)
            region_cols = []
            for col1, col2 in mapping:
                # Simple heuristic: if column name contains "region", "wilayah", or "kode"
                if any(x in str(col1).lower() for x in ['region', 'wilayah', 'kode']):
                    region_cols.append((col1, col2))
            
            if not region_cols:
                self.log("⚠️ No region columns detected in mappings - will compare all records")
                region_cols = [(None, None)]  # Dummy value to proceed without region filtering
            
            required_cols1 = [idsbr_col] + [m[0] for m in mapping]
            required_cols2 = [m[1] for m in mapping]
            
            # Add additional columns to required columns lists
            for src, col in add_cols:
                if src == "File 1" and col not in required_cols1:
                    required_cols1.append(col)
                elif src == "File 2" and col not in required_cols2:
                    required_cols2.append(col)
            
            missing_cols1 = [col for col in required_cols1 if col not in self.df1.columns]
            missing_cols2 = [col for col in required_cols2 if col not in self.df2.columns]
            
            if missing_cols1:
                self.log(f"❌ Missing columns in File 1: {', '.join(missing_cols1)}")
                return
                
            if missing_cols2:
                self.log(f"❌ Missing columns in File 2: {', '.join(missing_cols2)}")
                return

            try:
                df1_filtered = self.df1[required_cols1].copy()
                df2_filtered = self.df2[required_cols2].copy()
                
                # We're already reading everything as strings, but just to be sure
                # Convert all potential region columns to strings explicitly
                for col1, col2 in region_cols:
                    if col1:
                        df1_filtered[col1] = df1_filtered[col1].astype(str)
                    if col2:
                        df2_filtered[col2] = df2_filtered[col2].astype(str)
                
            except Exception as e:
                self.log(f"❌ Error preparing data: {str(e)}")
                return

            best_matches = []
            
            # Process each region pair
            for region1_col, region2_col in region_cols:
                if region1_col and region2_col:
                    # Get common region values
                    wilayah_set = set(df1_filtered[region1_col]).intersection(set(df2_filtered[region2_col]))
                    
                    if not wilayah_set:
                        self.log(f"⚠️ No common values found between '{region1_col}' and '{region2_col}'")
                        continue
                        
                    self.log(f"🔍 Processing {len(wilayah_set)} common regions")
                    
                    for wilayah in wilayah_set:
                        df1_grp = df1_filtered[df1_filtered[region1_col] == wilayah]
                        df2_grp = df2_filtered[df2_filtered[region2_col] == wilayah]
                        
                        self.process_record_matches(df1_grp, df2_grp, idsbr_col, mapping, best_matches, 
                                                    add_cols, wilayah)
                else:
                    # No region filtering - compare all records
                    self.process_record_matches(df1_filtered, df2_filtered, idsbr_col, mapping, best_matches,
                                                add_cols, "All Records")

            if best_matches:
                try:
                    outfile = self.output_file_entry.get()
                    outdf = pd.DataFrame(best_matches)
                    
                    # Always export with text format preservation
                    with pd.ExcelWriter(outfile, engine='openpyxl') as writer:
                        outdf.to_excel(writer, index=False)
                        
                        # Access the workbook and the sheet
                        workbook = writer.book
                        worksheet = writer.sheets['Sheet1']
                        
                        # Set the format to Text for all cells (Excel treats as text not numbers)
                        for i, col in enumerate(outdf.columns):
                            # Excel columns are 1-indexed
                            col_letter = worksheet.cell(row=1, column=i+1).column_letter
                            
                            # Apply text format to the whole column (excluding header)
                            for row in range(2, len(outdf) + 2):  # +2 because Excel is 1-indexed and we have headers
                                cell = worksheet.cell(row=row, column=i+1)
                                cell.number_format = '@'  # '@' is the number format code for Text
                    
                    self.log(f"✅ Matching complete. Results saved to {outfile}")
                    messagebox.showinfo("Success", f"Matching complete!\nResults saved to {outfile}")
                except Exception as e:
                    self.log(f"❌ Error saving results: {str(e)}")
            else:
                self.log("⚠️ No matches found.")
                messagebox.showinfo("No Matches", "No matches were found between the files.")
                
        finally:
            self.running = False

    # Modified to include additional output columns
    def process_record_matches(self, df1, df2, idsbr_col, mapping, best_matches, add_cols, region_name=""):
        record_count1 = len(df1)
        record_count2 = len(df2)
        
        if record_count1 == 0 or record_count2 == 0:
            self.log(f"⚠️ No records to compare in region {region_name}")
            return
            
        self.log(f"🔍 Comparing {record_count1} x {record_count2} records" + 
                (f" in region {region_name}" if region_name else ""))

        for _, r1 in df1.iterrows():
            # Keep track of top 3 matches
            top_matches = []  # List of (score, row) tuples
            
            for _, r2 in df2.iterrows():
                scores = []
                
                for col1, col2 in mapping:
                    val1 = r1[col1] if col1 in r1 and not pd.isna(r1[col1]) else ""
                    val2 = r2[col2] if col2 in r2 and not pd.isna(r2[col2]) else ""
                    
                    # Clean the values
                    t1 = self.clean_text(val1)
                    t2 = self.clean_text(val2)
                    
                    # Get match score using field-specific method
                    score = self.get_match_score(t1, t2, col1)
                    if score > 0:  # Only add non-zero scores
                        scores.append(score)

                if scores:
                    avg_score = sum(scores) / len(scores)
                    
                    # Always store as string to preserve leading zeros
                    idsbr_value = str(r1[idsbr_col]) if idsbr_col in r1 and not pd.isna(r1[idsbr_col]) else "Unknown"
                    
                    match_row = {
                        'IDSBR': idsbr_value,
                        'avg_score': str(round(avg_score, self.settings['score_precision']))
                    }
                    
                    if region_name:
                        match_row['region'] = str(region_name)
                    
                    for idx, (col1, col2) in enumerate(mapping):
                        # Store all values as strings
                        val1 = str(r1[col1]) if col1 in r1 and not pd.isna(r1[col1]) else ""
                        val2 = str(r2[col2]) if col2 in r2 and not pd.isna(r2[col2]) else ""
                        
                        match_row[f'{col1}_df1'] = val1
                        match_row[f'{col2}_df2'] = val2
                        
                        if idx < len(scores):
                            match_row[f'score_{col1}'] = str(round(scores[idx], self.settings['score_precision']))

                    # Add the additional output columns (as strings)
                    for src, col in add_cols:
                        if src == "File 1":
                            val = str(r1[col]) if col in r1 and not pd.isna(r1[col]) else ""
                            match_row[f'add_{col}_df1'] = val
                        elif src == "File 2":
                            val = str(r2[col]) if col in r2 and not pd.isna(r2[col]) else ""
                            match_row[f'add_{col}_df2'] = val
                    
                    # Add match to top matches list
                    top_matches.append((avg_score, match_row))
                    
                    # Keep only top 3 matches
                    top_matches.sort(key=lambda x: x[0], reverse=True)
                    if len(top_matches) > 3:
                        top_matches = top_matches[:3]
            
            # Add all top matches to the results
            for score, match_row in top_matches:
                best_matches.append(match_row)


if __name__ == "__main__":
    root = tk.Tk()
    app = FuzzyMatcherGUI(root)
    root.mainloop()