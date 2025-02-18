#!/usr/bin/env python3
"""
Excel Manager GUI

This script provides a modern Tkinter GUI for managing Excel files.
It allows you to combine multiple Excel files or split a single Excel file
based on a column value. Password-protected files are supported.

Author: Obaid Aldosari (Enhanced by ChatGPT)
GitHub: https://github.com/ODosari/manage_excel_sheets_and_files
"""

import os
import sys
import time
import glob
import logging
import tempfile
import io
import traceback
from contextlib import closing

import pandas as pd
import msoffcrypto
from openpyxl import load_workbook
from msoffcrypto.exceptions import InvalidKeyError

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# --- Logging configuration (logs to file only) ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    filename="manage_excel.log",
    filemode="a",
)
logger = logging.getLogger()


# --- Core Functions (modified for GUI use) --- #

def get_timestamped_filename(base_path, prefix, extension):
    """Generate a unique filename using a timestamp."""
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f"{prefix}_{timestamp}{extension}")


def normalize_sheet_name(base_name, sheet_name, existing_names):
    """Generate a safe, unique Excel sheet name."""
    raw_name = f"{base_name}_{sheet_name}".replace(" ", "_")
    safe_name = raw_name[:31]
    counter = 1
    while safe_name in existing_names:
        suffix = f"_{counter}"
        safe_name = raw_name[:31 - len(suffix)] + suffix
        counter += 1
    existing_names.add(safe_name)
    return safe_name


def unprotect_excel_file(file_path, default_password=None, max_attempts=3):
    """
    Unprotect an Excel file and return a temporary unprotected copy.
    Supports a default password for encrypted files.
    """
    try:
        with open(file_path, "rb") as f:
            office_file = msoffcrypto.OfficeFile(f)
            if office_file.is_encrypted():
                attempts = 0
                password = default_password
                while attempts < max_attempts:
                    try:
                        if password is None:
                            raise Exception("No password provided")
                        decrypted = io.BytesIO()
                        office_file.load_key(password=password)
                        office_file.decrypt(decrypted)
                        decrypted.seek(0)
                        wb = load_workbook(decrypted, read_only=False, keep_vba=True)
                        logger.info(f"Password accepted for {os.path.basename(file_path)}")
                        break
                    except InvalidKeyError as ike:
                        attempts += 1
                        logger.error(
                            f"Incorrect password for {os.path.basename(file_path)} (Attempt {attempts}/{max_attempts}): {ike}"
                        )
                        if attempts < max_attempts:
                            raise Exception("Incorrect password") from ike
                        else:
                            raise Exception("Maximum password attempts reached.") from ike
            else:
                wb = load_workbook(file_path, read_only=False, keep_vba=True)

        # Remove protection
        wb.security = None
        for sheet in wb.worksheets:
            sheet.protection.enabled = False
            sheet.protection.sheet = False

        temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        wb.save(temp_file.name)
        temp_file.close()
        logger.debug(f"Unprotected file saved to temporary file: {temp_file.name}")
        return temp_file.name

    except Exception as e:
        logger.error(f"Failed to open {os.path.basename(file_path)}: {e}")
        traceback.print_exc()
        return None


def get_sheets_from_file(file_path, default_password=None):
    """
    Return the list of sheet names for an Excel file.
    """
    temp_file = unprotect_excel_file(file_path, default_password=default_password)
    if temp_file is None:
        return []
    try:
        with closing(pd.ExcelFile(temp_file, engine="openpyxl")) as workbook:
            sheets = workbook.sheet_names
        return sheets
    except Exception as e:
        logger.error(f"Error reading sheets from {os.path.basename(file_path)}: {e}")
        return []
    finally:
        if temp_file and os.path.exists(temp_file):
            os.unlink(temp_file)


def read_sheet_from_file(file_path, sheet_name, default_password=None):
    """
    Read a sheet from an Excel file and return a DataFrame.
    """
    temp_file = unprotect_excel_file(file_path, default_password=default_password)
    if temp_file is None:
        return None
    try:
        df = pd.read_excel(temp_file, sheet_name=sheet_name, engine="openpyxl")
        return df
    except Exception as e:
        logger.error(f"Error reading sheet {sheet_name} from {os.path.basename(file_path)}: {e}")
        return None
    finally:
        if temp_file and os.path.exists(temp_file):
            os.unlink(temp_file)


def combine_files(selected_files, combine_mode="one_sheet", default_password=None):
    """
    Combine selected Excel files.

    combine_mode: "one_sheet" for a single combined sheet,
                  "separate_sheets" for a workbook with one sheet per file.
    Returns the output file path or None.
    """
    if not selected_files:
        return None

    out_dir = os.path.dirname(selected_files[0])
    output_file = get_timestamped_filename(out_dir, "Combined", ".xlsx")

    try:
        if combine_mode == "one_sheet":
            combined_df = pd.DataFrame()
            for file_path in selected_files:
                sheets = get_sheets_from_file(file_path, default_password=default_password)
                if sheets:
                    df = read_sheet_from_file(file_path, sheets[0], default_password=default_password)
                    if df is not None:
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
            combined_df.to_excel(output_file, sheet_name="Combined", engine="openpyxl", index=False)
        else:
            base_names_used = set()
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for file_path in selected_files:
                    base_name = os.path.splitext(os.path.basename(file_path))[0]
                    sheets = get_sheets_from_file(file_path, default_password=default_password)
                    if sheets:
                        df = read_sheet_from_file(file_path, sheets[0], default_password=default_password)
                        if df is not None:
                            safe_sheet = normalize_sheet_name(base_name, sheets[0], base_names_used)
                            df.to_excel(writer, sheet_name=safe_sheet, index=False)
        logger.info(f"Files combined successfully. Output File: {os.path.basename(output_file)}")
        return output_file
    except Exception as e:
        logger.error(f"An error occurred while combining files: {e}")
        traceback.print_exc()
        return None


def split_file(file_path, sheet_name, column, split_mode="files", default_password=None):
    """
    Split an Excel file based on the unique values in a column.

    split_mode: "files" to split into separate files,
                "sheets" to split into separate sheets in one workbook.
    'column' can be either an integer index or a string (column name).
    Returns the output file path (or list of file paths) or None.
    """
    df = read_sheet_from_file(file_path, sheet_name, default_password=default_password)
    if df is None:
        return None

    columns = list(df.columns)
    # Determine the column to split by
    if isinstance(column, int):
        if column < 0 or column >= len(columns):
            logger.error("Invalid column index selected.")
            return None
        split_column = columns[column]
    elif isinstance(column, str):
        if column not in columns:
            logger.error("Invalid column name selected.")
            return None
        split_column = column
    else:
        logger.error("Column must be an int or str.")
        return None

    out_dir = os.path.dirname(file_path)

    try:
        if split_mode == "files":
            output_files = []
            for value in df[split_column].unique():
                safe_value = str(value).replace(" ", "_")
                output_file = get_timestamped_filename(
                    out_dir,
                    f"{os.path.splitext(os.path.basename(file_path))[0]}_{sheet_name}_{split_column}_{safe_value}",
                    ".xlsx"
                )
                filtered_df = df[df[split_column] == value]
                filtered_df.to_excel(output_file, sheet_name=str(value)[:31], engine="openpyxl", index=False)
                output_files.append(output_file)
                logger.info(f"Data for '{value}' written to {os.path.basename(output_file)}")
            logger.info("Data split into files successfully.")
            return output_files
        else:
            output_file = get_timestamped_filename(
                out_dir,
                f"{os.path.splitext(os.path.basename(file_path))[0]}_{sheet_name}_split",
                ".xlsx"
            )
            existing_sheet_names = set()
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for value in df[split_column].unique():
                    raw_sheet_name = str(value)
                    safe_sheet = normalize_sheet_name(raw_sheet_name, "",
                                                      existing_sheet_names) if raw_sheet_name.strip() else "Empty"
                    filtered_df = df[df[split_column] == value]
                    filtered_df.to_excel(writer, sheet_name=safe_sheet, index=False)
            logger.info(f"Data split into sheets successfully. Output File: {os.path.basename(output_file)}")
            return output_file
    except Exception as e:
        logger.error(f"An error occurred while splitting the file: {e}")
        traceback.print_exc()
        return None


# --- GUI Class --- #

class ExcelManagerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Manage Excel Sheets and Files")
        self.geometry("900x600")  # Larger window size
        self.center_window()  # Center the window on the screen

        # Configure dark theme with green accents and a flat modern design
        self.style = ttk.Style(self)
        self.style.theme_use("clam")
        self.configure(bg="#2e2e2e")
        self.style.configure(".", background="#2e2e2e", foreground="#ffffff")
        self.style.configure("TFrame", background="#2e2e2e")
        self.style.configure("TLabel", background="#2e2e2e", foreground="#ffffff")
        self.style.configure("TCheckbutton", background="#2e2e2e", foreground="#ffffff")
        self.style.configure("TButton", background="#4CAF50", foreground="#ffffff", relief="flat", borderwidth=0)
        self.style.map("TButton", background=[("active", "#45a049")])
        self.style.configure("TEntry", fieldbackground="#3e3e3e", foreground="#ffffff")
        self.style.configure("TCombobox", fieldbackground="#3e3e3e", foreground="#ffffff")

        # Variables for Combine tab
        self.combine_dir = tk.StringVar()
        self.combine_mode = tk.StringVar(value="one_sheet")
        self.default_password_combine = tk.StringVar()
        self.select_all_var = tk.BooleanVar(value=True)  # For select/deselect all

        # Dictionary mapping file path -> BooleanVar (for checkbuttons)
        self.files_vars = {}

        # Variables for Split tab
        self.split_file_path = tk.StringVar()
        self.split_sheet = tk.StringVar()
        self.default_password_split = tk.StringVar()
        self.col_combo = None  # Will be created in create_split_tab()

        self.split_mode = tk.StringVar(value="files")

        # Create Notebook
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.create_combine_tab()
        self.create_split_tab()

    def center_window(self):
        """Center the window on the screen."""
        self.update_idletasks()
        width = self.winfo_width() or 900
        height = self.winfo_height() or 600
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ---------- Combine Tab ---------- #
    def create_combine_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Combine Files")

        # Directory selection (automatically load files on selection)
        dir_frame = ttk.Frame(frame)
        dir_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(dir_frame, text="Directory:").pack(side=tk.LEFT)
        entry = ttk.Entry(dir_frame, textvariable=self.combine_dir, width=50)
        entry.pack(side=tk.LEFT, padx=5)
        ttk.Button(dir_frame, text="Browse", command=self.browse_directory).pack(side=tk.LEFT)

        # Optional password
        pwd_frame = ttk.Frame(frame)
        pwd_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(pwd_frame, text="Default Password (if any):").pack(side=tk.LEFT)
        ttk.Entry(pwd_frame, textvariable=self.default_password_combine, width=20, show="*").pack(side=tk.LEFT, padx=5)

        # Select/Deselect All checkbox
        selectall_frame = ttk.Frame(frame)
        selectall_frame.pack(fill=tk.X, padx=10, pady=5)
        chk_all = ttk.Checkbutton(
            selectall_frame,
            text="Select/Deselect All",
            variable=self.select_all_var,
            command=self.update_all_checkbuttons
        )
        chk_all.pack(anchor="w")

        # Files checklist frame with scrollbar
        list_frame = ttk.LabelFrame(frame, text="Excel Files Found")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.canvas = tk.Canvas(list_frame, bg="#2e2e2e", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.checklist_frame = ttk.Frame(self.canvas, style="TFrame")
        self.canvas.create_window((0, 0), window=self.checklist_frame, anchor="nw")

        # Refresh Files button
        ttk.Button(frame, text="Refresh File List", command=self.load_files_list).pack(pady=5)

        # Combine mode radiobuttons
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(mode_frame, text="Combine Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="One Sheet", variable=self.combine_mode, value="one_sheet").pack(side=tk.LEFT,
                                                                                                          padx=5)
        ttk.Radiobutton(mode_frame, text="Separate Sheets", variable=self.combine_mode, value="separate_sheets").pack(
            side=tk.LEFT, padx=5)

        # Combine Button
        ttk.Button(frame, text="Combine Files", command=self.combine_files_action).pack(pady=10)

    def browse_directory(self):
        dir_selected = filedialog.askdirectory()
        if dir_selected:
            self.combine_dir.set(dir_selected)
            self.load_files_list()  # Automatically load files

    def load_files_list(self):
        directory = self.combine_dir.get()
        # Clear any previous checkbuttons
        for widget in self.checklist_frame.winfo_children():
            widget.destroy()
        self.files_vars.clear()

        if not directory or not os.path.isdir(directory):
            messagebox.showerror("Error", "Please select a valid directory.")
            return
        files = glob.glob(os.path.join(directory, "*.xlsx"))
        if not files:
            messagebox.showinfo("No Files", "No Excel (.xlsx) files found in the directory.")
            return

        for file in files:
            var = tk.BooleanVar(value=self.select_all_var.get())
            chk = ttk.Checkbutton(self.checklist_frame, text=os.path.basename(file), variable=var)
            chk.pack(anchor="w")
            self.files_vars[file] = var

    def update_all_checkbuttons(self):
        # Set each file's BooleanVar to match the select_all_var value.
        for var in self.files_vars.values():
            var.set(self.select_all_var.get())

    def combine_files_action(self):
        selected_files = [file for file, var in self.files_vars.items() if var.get()]
        if not selected_files:
            messagebox.showerror("Error", "Please select at least one file from the list.")
            return
        mode = self.combine_mode.get()
        default_pwd = self.default_password_combine.get() or None
        output = combine_files(selected_files, combine_mode=mode, default_password=default_pwd)
        if output:
            messagebox.showinfo("Success", f"Files combined successfully.\nOutput File:\n{output}")
        else:
            messagebox.showerror("Error", "An error occurred while combining files.")

    # ---------- Split Tab ---------- #
    def create_split_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="Split File")

        # File selection
        file_frame = ttk.Frame(frame)
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(file_frame, text="Excel File:").pack(side=tk.LEFT)
        ttk.Entry(file_frame, textvariable=self.split_file_path, width=50).pack(side=tk.LEFT, padx=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_file).pack(side=tk.LEFT)

        # Optional password
        pwd_frame = ttk.Frame(frame)
        pwd_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(pwd_frame, text="Default Password (if any):").pack(side=tk.LEFT)
        ttk.Entry(pwd_frame, textvariable=self.default_password_split, width=20, show="*").pack(side=tk.LEFT, padx=5)

        # Sheet selection
        sheet_frame = ttk.Frame(frame)
        sheet_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(sheet_frame, text="Sheet:").pack(side=tk.LEFT)
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.split_sheet, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # Column selection using a combobox (automatically loaded)
        col_frame = ttk.Frame(frame)
        col_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(col_frame, text="Split by Column:").pack(side=tk.LEFT)
        self.col_combo = ttk.Combobox(col_frame, state="readonly")
        self.col_combo.pack(side=tk.LEFT, padx=5)

        # Split mode radiobuttons
        mode_frame = ttk.Frame(frame)
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        ttk.Label(mode_frame, text="Split Mode:").pack(side=tk.LEFT)
        ttk.Radiobutton(mode_frame, text="Files", variable=self.split_mode, value="files").pack(side=tk.LEFT, padx=5)
        ttk.Radiobutton(mode_frame, text="Sheets", variable=self.split_mode, value="sheets").pack(side=tk.LEFT, padx=5)

        # Split Button
        ttk.Button(frame, text="Split File", command=self.split_file_action).pack(pady=10)

    def browse_file(self):
        file_selected = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_selected:
            self.split_file_path.set(file_selected)
            self.load_sheet_options_and_columns()

    def load_sheet_options_and_columns(self):
        file_path = self.split_file_path.get()
        if not file_path or not os.path.isfile(file_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        default_pwd = self.default_password_split.get() or None
        sheets = get_sheets_from_file(file_path, default_password=default_pwd)
        if not sheets:
            messagebox.showerror("Error", "Could not load sheets from the file.")
            return
        self.sheet_combo["values"] = sheets
        self.sheet_combo.current(0)
        self.split_sheet.set(sheets[0])
        self.load_columns()

    def load_columns(self):
        file_path = self.split_file_path.get()
        sheet = self.split_sheet.get()
        if not file_path or not os.path.isfile(file_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        default_pwd = self.default_password_split.get() or None
        df = read_sheet_from_file(file_path, sheet, default_password=default_pwd)
        if df is None:
            messagebox.showerror("Error", "Could not load columns from the file.")
            return
        columns = list(df.columns)
        if not columns:
            messagebox.showerror("Error", "No columns found in the selected sheet.")
            return
        self.col_combo["values"] = columns
        self.col_combo.current(0)

    def on_sheet_change(self, event):
        self.load_columns()

    def split_file_action(self):
        file_path = self.split_file_path.get()
        sheet = self.split_sheet.get()
        selected_column = self.col_combo.get()
        if not file_path or not sheet or not selected_column:
            messagebox.showerror("Error", "Please select a file, sheet, and column.")
            return
        mode = self.split_mode.get()
        default_pwd = self.default_password_split.get() or None
        output = split_file(file_path, sheet, selected_column, split_mode=mode, default_password=default_pwd)
        if output:
            if mode == "files":
                message = "Data split into files successfully.\nFiles:\n" + "\n".join(output)
            else:
                message = f"Data split into sheets successfully.\nOutput File:\n{output}"
            messagebox.showinfo("Success", message)
        else:
            messagebox.showerror("Error", "An error occurred while splitting the file.")


# --- Main --- #

def main():
    app = ExcelManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
