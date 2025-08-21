#!/usr/bin/env python3
"""
Excel ↔ CSV Converter
Lightweight cross-platform converter for Excel (.xls, .xlsx) and CSV files
Handles large files efficiently with minimal memory usage
"""

import os
import sys
import csv
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import openpyxl
from openpyxl import Workbook
import xlrd
import warnings

# Suppress warnings from xlrd
warnings.filterwarnings("ignore", category=UserWarning, module="xlrd")

class ExcelCSVConverter:
    def __init__(self):
        self.supported_excel = ('.xlsx', '.xlsm', '.xls')
        self.supported_csv = ('.csv', '.txt')
        
    def excel_to_csv(self, excel_path, csv_path=None, sheet_index=0):
        """
        Convert Excel file to CSV with UTF-8 encoding.
        Handles large files efficiently using iterators.
        """
        excel_path = Path(excel_path)
        
        if csv_path is None:
            csv_path = excel_path.with_suffix('.csv')
        
        try:
            file_ext = excel_path.suffix.lower()
            
            if file_ext == '.xls':
                # Handle older .xls files
                workbook = xlrd.open_workbook(str(excel_path))
                sheet = workbook.sheet_by_index(sheet_index)
                
                with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    for row_idx in range(sheet.nrows):
                        row = []
                        for col_idx in range(sheet.ncols):
                            cell = sheet.cell(row_idx, col_idx)
                            # Handle different cell types
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                # Convert Excel date to readable format
                                from xlrd import xldate_as_datetime
                                try:
                                    dt = xldate_as_datetime(cell.value, workbook.datemode)
                                    row.append(dt.strftime('%Y-%m-%d %H:%M:%S'))
                                except:
                                    row.append(str(cell.value))
                            else:
                                row.append(str(cell.value))
                        writer.writerow(row)
            else:
                # Handle .xlsx files with openpyxl (memory efficient for large files)
                workbook = openpyxl.load_workbook(str(excel_path), read_only=True, data_only=True)
                sheet = workbook.worksheets[sheet_index]
                
                with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
                    writer = csv.writer(csvfile)
                    # Use iter_rows for memory efficiency with large files
                    for row in sheet.iter_rows(values_only=True):
                        # Convert None values to empty strings
                        row_data = ['' if cell is None else str(cell) for cell in row]
                        writer.writerow(row_data)
                
                workbook.close()
            
            return True, f"Successfully converted to: {csv_path}"
            
        except Exception as e:
            return False, f"Error converting Excel to CSV: {str(e)}"
    
    def csv_to_excel(self, csv_path, excel_path=None, xlsx=True):
        """
        Convert CSV file to Excel format.
        Handles large files efficiently using streaming.
        """
        csv_path = Path(csv_path)
        
        if excel_path is None:
            extension = '.xlsx' if xlsx else '.xls'
            excel_path = csv_path.with_suffix(extension)
        
        try:
            # Create a new workbook
            workbook = Workbook(write_only=True)  # write_only mode for memory efficiency
            sheet = workbook.create_sheet('Sheet1')
            
            # Read CSV and write to Excel
            with open(csv_path, 'r', encoding='utf-8') as csvfile:
                # Try to detect delimiter
                sample = csvfile.read(1024)
                csvfile.seek(0)
                sniffer = csv.Sniffer()
                try:
                    delimiter = sniffer.sniff(sample).delimiter
                except:
                    delimiter = ','
                
                reader = csv.reader(csvfile, delimiter=delimiter)
                for row in reader:
                    # Convert strings to appropriate types
                    processed_row = []
                    for cell in row:
                        if cell == '':
                            processed_row.append(None)
                        else:
                            # Try to convert to number if possible
                            try:
                                if '.' in cell:
                                    processed_row.append(float(cell))
                                else:
                                    processed_row.append(int(cell))
                            except ValueError:
                                processed_row.append(cell)
                    
                    sheet.append(processed_row)
            
            # Save the workbook
            workbook.save(str(excel_path))
            workbook.close()
            
            return True, f"Successfully converted to: {excel_path}"
            
        except Exception as e:
            return False, f"Error converting CSV to Excel: {str(e)}"
    
    def get_excel_info(self, excel_path):
        """Get information about Excel file (sheet names, dimensions, etc.)"""
        try:
            excel_path = Path(excel_path)
            file_ext = excel_path.suffix.lower()
            
            if file_ext == '.xls':
                workbook = xlrd.open_workbook(str(excel_path))
                sheets = []
                for i in range(workbook.nsheets):
                    sheet = workbook.sheet_by_index(i)
                    sheets.append({
                        'name': sheet.name,
                        'rows': sheet.nrows,
                        'cols': sheet.ncols
                    })
                return sheets
            else:
                workbook = openpyxl.load_workbook(str(excel_path), read_only=True)
                sheets = []
                for sheet_name in workbook.sheetnames:
                    sheet = workbook[sheet_name]
                    sheets.append({
                        'name': sheet_name,
                        'rows': sheet.max_row,
                        'cols': sheet.max_column
                    })
                workbook.close()
                return sheets
        except Exception as e:
            return None

class ConverterGUI:
    def __init__(self):
        self.converter = ExcelCSVConverter()
        self.root = tk.Tk()
        self.root.title("Excel ↔ CSV Converter")
        self.root.geometry("500x300")
        self.root.resizable(False, False)
        
        # Set icon for different platforms
        if sys.platform == "win32":
            self.root.iconbitmap(default='')  # Default Windows icon
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Title
        title = ttk.Label(main_frame, text="Excel ↔ CSV Converter", font=('Helvetica', 16, 'bold'))
        title.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Instructions
        instructions = ttk.Label(main_frame, text="Select a file to convert:", font=('Helvetica', 10))
        instructions.grid(row=1, column=0, columnspan=2, pady=5)
        
        # File info label
        self.file_info = ttk.Label(main_frame, text="No file selected", font=('Helvetica', 9))
        self.file_info.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Sheet selection frame (hidden initially)
        self.sheet_frame = ttk.Frame(main_frame)
        self.sheet_frame.grid(row=3, column=0, columnspan=2, pady=5)
        self.sheet_frame.grid_remove()
        
        ttk.Label(self.sheet_frame, text="Select sheet:").pack(side=tk.LEFT, padx=5)
        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.sheet_frame, textvariable=self.sheet_var, width=30, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=4, column=0, columnspan=2, pady=20)
        
        # Convert Excel to CSV button
        excel_to_csv_btn = ttk.Button(
            button_frame,
            text="Convert Excel to CSV",
            command=self.convert_excel_to_csv,
            width=20
        )
        excel_to_csv_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # Convert CSV to Excel button
        csv_to_excel_btn = ttk.Button(
            button_frame,
            text="Convert CSV to Excel",
            command=self.convert_csv_to_excel,
            width=20
        )
        csv_to_excel_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Batch convert button
        batch_btn = ttk.Button(
            button_frame,
            text="Batch Convert Folder",
            command=self.batch_convert,
            width=20
        )
        batch_btn.grid(row=1, column=0, columnspan=2, pady=5)
        
        # Status bar
        self.status_bar = ttk.Label(main_frame, text="Ready", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.grid(row=5, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
    def convert_excel_to_csv(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[
                ("Excel files", "*.xlsx *.xlsm *.xls"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.file_info.config(text=f"Selected: {Path(file_path).name}")
            
            # Get sheet information
            sheets = self.converter.get_excel_info(file_path)
            if sheets and len(sheets) > 1:
                # Show sheet selection
                self.sheet_frame.grid()
                sheet_names = [f"{s['name']} ({s['rows']} rows, {s['cols']} cols)" for s in sheets]
                self.sheet_combo['values'] = sheet_names
                self.sheet_combo.current(0)
                
                # Wait for user to select sheet
                result = messagebox.askquestion("Multiple Sheets", 
                                               "This file has multiple sheets. Convert the first sheet?")
                if result == 'no':
                    return
                sheet_index = 0
            else:
                sheet_index = 0
                self.sheet_frame.grid_remove()
            
            # Ask for output location
            output_path = filedialog.asksaveasfilename(
                title="Save CSV as",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=Path(file_path).stem + ".csv"
            )
            
            if output_path:
                self.status_bar.config(text="Converting...")
                self.root.update()
                
                success, message = self.converter.excel_to_csv(file_path, output_path, sheet_index)
                
                if success:
                    self.status_bar.config(text="Conversion successful!")
                    messagebox.showinfo("Success", message)
                else:
                    self.status_bar.config(text="Conversion failed")
                    messagebox.showerror("Error", message)
    
    def convert_csv_to_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[
                ("CSV files", "*.csv"),
                ("Text files", "*.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.file_info.config(text=f"Selected: {Path(file_path).name}")
            self.sheet_frame.grid_remove()
            
            # Ask for format
            format_choice = messagebox.askyesno("Excel Format", 
                                               "Save as .xlsx? (Yes for .xlsx, No for .xls)")
            
            # Ask for output location
            ext = ".xlsx" if format_choice else ".xls"
            output_path = filedialog.asksaveasfilename(
                title="Save Excel as",
                defaultextension=ext,
                filetypes=[("Excel files", f"*{ext}"), ("All files", "*.*")],
                initialfile=Path(file_path).stem + ext
            )
            
            if output_path:
                self.status_bar.config(text="Converting...")
                self.root.update()
                
                success, message = self.converter.csv_to_excel(file_path, output_path, format_choice)
                
                if success:
                    self.status_bar.config(text="Conversion successful!")
                    messagebox.showinfo("Success", message)
                else:
                    self.status_bar.config(text="Conversion failed")
                    messagebox.showerror("Error", message)
    
    def batch_convert(self):
        folder_path = filedialog.askdirectory(title="Select folder for batch conversion")
        
        if folder_path:
            # Ask what type of conversion
            choice = messagebox.askyesno("Batch Conversion", 
                                        "Convert Excel to CSV? (No = CSV to Excel)")
            
            folder = Path(folder_path)
            converted = 0
            failed = 0
            
            if choice:  # Excel to CSV
                files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls")) + list(folder.glob("*.xlsm"))
                for file_path in files:
                    output_path = file_path.with_suffix('.csv')
                    success, _ = self.converter.excel_to_csv(file_path, output_path)
                    if success:
                        converted += 1
                    else:
                        failed += 1
            else:  # CSV to Excel
                files = list(folder.glob("*.csv"))
                for file_path in files:
                    output_path = file_path.with_suffix('.xlsx')
                    success, _ = self.converter.csv_to_excel(file_path, output_path)
                    if success:
                        converted += 1
                    else:
                        failed += 1
            
            message = f"Batch conversion complete!\nConverted: {converted} files\nFailed: {failed} files"
            messagebox.showinfo("Batch Conversion Results", message)
            self.status_bar.config(text=f"Batch conversion complete: {converted} succeeded, {failed} failed")
    
    def run(self):
        self.root.mainloop()

def main():
    # Check if running with command line arguments
    if len(sys.argv) > 1:
        converter = ExcelCSVConverter()
        input_file = Path(sys.argv[1])
        
        if not input_file.exists():
            print(f"Error: File '{input_file}' not found")
            sys.exit(1)
        
        # Determine conversion type based on file extension
        if input_file.suffix.lower() in converter.supported_excel:
            # Convert Excel to CSV
            output_file = input_file.with_suffix('.csv') if len(sys.argv) < 3 else Path(sys.argv[2])
            success, message = converter.excel_to_csv(input_file, output_file)
        elif input_file.suffix.lower() in converter.supported_csv:
            # Convert CSV to Excel
            output_file = input_file.with_suffix('.xlsx') if len(sys.argv) < 3 else Path(sys.argv[2])
            success, message = converter.csv_to_excel(input_file, output_file)
        else:
            print(f"Error: Unsupported file type '{input_file.suffix}'")
            sys.exit(1)
        
        print(message)
        sys.exit(0 if success else 1)
    else:
        # Run GUI
        app = ConverterGUI()
        app.run()

if __name__ == "__main__":
    main()