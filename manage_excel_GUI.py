#!/usr/bin/env python3
"""
Manage Excel Sheets and Files Utility - GUI

This script allows you to combine multiple Excel files into one,
or split a single Excel file into multiple sheets or files based on a specific column.
It includes enhanced handling for password-protected files.

Version = "0.2"
Author: Obaid Aldosari
GitHub: https://github.com/ODosari/manage_excel_sheets_and_files
"""

import os
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

# --- Logging configuration ---
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    filename="manage_excel.log",
    filemode="a",
)
logger = logging.getLogger()


def get_timestamped_filename(base_path, prefix, extension):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f"{prefix}_{timestamp}{extension}")


def normalize_sheet_name(base_name, sheet_name, existing_names):
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
    try:
        import msoffcrypto
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
                        from openpyxl import load_workbook
                        wb = load_workbook(decrypted, read_only=False, keep_vba=True)
                        logger.info(f"Password accepted for {os.path.basename(file_path)}")
                        break
                    except InvalidKeyError as ike:
                        attempts += 1
                        logger.error(f"Incorrect password for {os.path.basename(file_path)} (Attempt {attempts}/{max_attempts}): {ike}")
                        if attempts < max_attempts:
                            raise Exception("Incorrect password") from ike
                        else:
                            raise Exception("Maximum password attempts reached.") from ike
            else:
                from openpyxl import load_workbook
                wb = load_workbook(file_path, read_only=False, keep_vba=True)

        # Remove protection
        wb.security = None
        for sheet in wb.worksheets:
            sheet.protection.enabled = False
            sheet.protection.sheet = False

        import tempfile
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
    temp_file = unprotect_excel_file(file_path, default_password=default_password)
    if temp_file is None:
        return []
    try:
        from contextlib import closing
        import pandas as pd
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
    temp_file = unprotect_excel_file(file_path, default_password=default_password)
    if temp_file is None:
        return None
    try:
        import pandas as pd
        df = pd.read_excel(temp_file, sheet_name=sheet_name, engine="openpyxl")
        return df
    except Exception as e:
        logger.error(f"Error reading sheet {sheet_name} from {os.path.basename(file_path)}: {e}")
        return None
    finally:
        if temp_file and os.path.exists(temp_file):
            os.unlink(temp_file)


def combine_files(selected_files, combine_mode="one_sheet", default_password=None):
    import pandas as pd
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
            import pandas as pd
            base_names_used = set()
            from openpyxl import load_workbook
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
    df = read_sheet_from_file(file_path, sheet_name, default_password=default_password)
    if df is None:
        return None

    columns = list(df.columns)
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
        import pandas as pd
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
            base_names_used = set()
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                for value in df[split_column].unique():
                    raw_sheet_name = str(value)
                    safe_sheet = normalize_sheet_name(raw_sheet_name, "", base_names_used) if raw_sheet_name.strip() else "Empty"
                    filtered_df = df[df[split_column] == value]
                    filtered_df.to_excel(writer, sheet_name=safe_sheet, index=False)
            logger.info(f"Data split into sheets successfully. Output File: {os.path.basename(output_file)}")
            return output_file
    except Exception as e:
        logger.error(f"An error occurred while splitting the file: {e}")
        traceback.print_exc()
        return None


class ExcelManagerGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Manage Excel Sheets and Files")
        self.geometry("1000x700")
        self.center_window()

        # --------------------- Dark Theme with Green Accents ---------------------
        self.style = ttk.Style(self)
        self.style.theme_use("clam")

        self.configure(bg="#2e2e2e")

        # Notebook styling
        self.style.configure("TNotebook", background="#2e2e2e", borderwidth=0)
        self.style.configure("TNotebook.Tab", background="#2e2e2e", foreground="#ffffff", padding=[10, 5])
        self.style.map("TNotebook.Tab",
                       background=[("selected", "#00a67d"), ("active", "#00a67d")],
                       foreground=[("selected", "#ffffff"), ("active", "#ffffff")]
                       )

        self.style.configure("TLabelframe", background="#2e2e2e", borderwidth=0)
        self.style.configure("TLabelframe.Label", background="#2e2e2e", foreground="#ffffff")

        self.style.configure(".", font=("IBM Plex Mono", 16))
        self.style.configure("TFrame", background="#2e2e2e")
        self.style.configure("TLabel", background="#2e2e2e", foreground="#ffffff")
        self.style.configure("TButton", background="#00a67d", foreground="#ffffff",
                             relief="flat", borderwidth=0)
        self.style.map("TButton", background=[("active", "#45A049")])
        self.style.configure("TCheckbutton", background="#2e2e2e", foreground="#ffffff")
        self.style.map("TCheckbutton",
                       background=[("active", "#2e2e2e")],
                       foreground=[("active", "#ffffff")])
        self.style.configure("TEntry", fieldbackground="#3e3e3e", foreground="#ffffff")
        self.style.configure("TCombobox", fieldbackground="#3e3e3e", foreground="#ffffff")
        self.style.map("TCombobox",
                       fieldbackground=[("readonly", "#3e3e3e")],
                       foreground=[("readonly", "#ffffff")])
        self.style.configure("TRadiobutton", background="#2e2e2e", foreground="#ffffff")
        self.style.map("TRadiobutton",
                       background=[("active", "#2e2e2e")],
                       foreground=[("active", "#ffffff")])

        self.style.configure("Vertical.TScrollbar", troughcolor="#2e2e2e", bordercolor="#2e2e2e",
                             background="#2e2e2e", arrowcolor="#ffffff")
        self.style.map("Vertical.TScrollbar",
                       background=[("active", "#45A049")],
                       arrowcolor=[("active", "#ffffff")])
        # --------------------------------------------------------------------------

        # Variables
        self.combine_dir = tk.StringVar()
        self.combine_mode = tk.StringVar(value="one_sheet")
        self.default_password_combine = tk.StringVar()
        self.select_all_var = tk.BooleanVar(value=True)
        self.files_vars = {}

        self.split_file_path = tk.StringVar()
        self.split_sheet = tk.StringVar()
        self.default_password_split = tk.StringVar()
        self.col_combo = None
        self.split_mode = tk.StringVar(value="files")

        self.notebook = ttk.Notebook(self, style="TNotebook")
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.create_combine_tab()
        self.create_split_tab()

    def center_window(self):
        self.update_idletasks()
        width = self.winfo_width() or 1000
        height = self.winfo_height() or 700
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ---------- Combine Tab ---------- #
    def create_combine_tab(self):
        frame = ttk.Frame(self.notebook, style="TFrame")
        self.notebook.add(frame, text="Combine Files")

        # Row 1: Directory
        dir_frame = ttk.Frame(frame, style="TFrame")
        dir_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_dir = ttk.Label(dir_frame, text="Directory:")
        lbl_dir.pack(side=tk.LEFT, anchor="w")
        entry_dir = ttk.Entry(dir_frame, textvariable=self.combine_dir, width=50)
        entry_dir.pack(side=tk.LEFT, anchor="w", padx=5)
        btn_browse = ttk.Button(dir_frame, text="Browse", command=self.browse_directory)
        btn_browse.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 2: Default password
        pwd_frame = ttk.Frame(frame, style="TFrame")
        pwd_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_pwd = ttk.Label(pwd_frame, text="Default Password (if any):")
        lbl_pwd.pack(side=tk.LEFT, anchor="w")
        entry_pwd = ttk.Entry(pwd_frame, textvariable=self.default_password_combine, width=20, show="*")
        entry_pwd.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 3: Combine mode
        mode_frame = ttk.Frame(frame, style="TFrame")
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_mode = ttk.Label(mode_frame, text="Combine Mode:")
        lbl_mode.pack(side=tk.LEFT, anchor="w")
        rbtn_one = ttk.Radiobutton(mode_frame, text="One Sheet", variable=self.combine_mode, value="one_sheet")
        rbtn_one.pack(side=tk.LEFT, anchor="w", padx=5)
        rbtn_sep = ttk.Radiobutton(mode_frame, text="Separate Sheets", variable=self.combine_mode,
                                   value="separate_sheets")
        rbtn_sep.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 4: File list
        list_frame = ttk.Labelframe(frame, text="Excel Files Found", style="TLabelframe")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        header_frame = ttk.Frame(list_frame, style="TFrame")
        header_frame.pack(fill=tk.X, padx=5, pady=5)
        chk_all = ttk.Checkbutton(header_frame, text="Select/Deselect All",
                                  variable=self.select_all_var, command=self.update_all_checkbuttons)
        chk_all.pack(side=tk.LEFT, anchor="w")

        self.canvas = tk.Canvas(list_frame, bg="#2e2e2e", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.canvas.yview, style="Vertical.TScrollbar")
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)
        self.canvas.bind('<Configure>', lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.checklist_frame = ttk.Frame(self.canvas, style="TFrame")
        self.canvas.create_window((0, 0), window=self.checklist_frame, anchor="nw")

        # Row 5: Combine button
        btn_combine = ttk.Button(frame, text="Combine Files", command=self.combine_files_action)
        btn_combine.pack(anchor="w", padx=10, pady=10)

    def browse_directory(self):
        dir_selected = filedialog.askdirectory()
        if dir_selected:
            self.combine_dir.set(dir_selected)
            self.load_files_list()

    def load_files_list(self):
        directory = self.combine_dir.get()
        for widget in self.checklist_frame.winfo_children():
            widget.destroy()
        self.files_vars.clear()

        if not directory or not os.path.isdir(directory):
            messagebox.showerror("Error", "Please select a valid directory.")
            return
        files = [file for file in glob.glob(os.path.join(directory, "*.xlsx"))
                 if not os.path.basename(file).startswith("~")]
        if not files:
            messagebox.showinfo("No Files", "No Excel (.xlsx) files found in the directory.")
            return

        for file in files:
            var = tk.BooleanVar(value=self.select_all_var.get())
            chk = ttk.Checkbutton(self.checklist_frame, text=os.path.basename(file), variable=var)
            chk.pack(anchor="w")
            self.files_vars[file] = var

    def update_all_checkbuttons(self):
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
        frame = ttk.Frame(self.notebook, style="TFrame")
        self.notebook.add(frame, text="Split File")

        # Row 1: File selection
        file_frame = ttk.Frame(frame, style="TFrame")
        file_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_file = ttk.Label(file_frame, text="Excel File:")
        lbl_file.pack(side=tk.LEFT, anchor="w")
        entry_file = ttk.Entry(file_frame, textvariable=self.split_file_path, width=50)
        entry_file.pack(side=tk.LEFT, anchor="w", padx=5)
        btn_browse = ttk.Button(file_frame, text="Browse", command=self.browse_file)
        btn_browse.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 2: Default password
        pwd_frame = ttk.Frame(frame, style="TFrame")
        pwd_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_pwd = ttk.Label(pwd_frame, text="Default Password (if any):")
        lbl_pwd.pack(side=tk.LEFT, anchor="w")
        entry_pwd = ttk.Entry(pwd_frame, textvariable=self.default_password_split, width=20, show="*")
        entry_pwd.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 3: Sheet selection
        sheet_frame = ttk.Frame(frame, style="TFrame")
        sheet_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_sheet = ttk.Label(sheet_frame, text="Sheet:")
        lbl_sheet.pack(side=tk.LEFT, anchor="w")
        self.sheet_combo = ttk.Combobox(sheet_frame, textvariable=self.split_sheet, state="readonly")
        self.sheet_combo.pack(side=tk.LEFT, anchor="w", padx=5)
        self.sheet_combo.bind("<<ComboboxSelected>>", self.on_sheet_change)

        # Row 4: Column selection
        col_frame = ttk.Frame(frame, style="TFrame")
        col_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_col = ttk.Label(col_frame, text="Split by Column:")
        lbl_col.pack(side=tk.LEFT, anchor="w")
        self.col_combo = ttk.Combobox(col_frame, state="readonly")
        self.col_combo.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 5: Split mode
        mode_frame = ttk.Frame(frame, style="TFrame")
        mode_frame.pack(fill=tk.X, padx=10, pady=5)
        lbl_mode = ttk.Label(mode_frame, text="Split Mode:")
        lbl_mode.pack(side=tk.LEFT, anchor="w")
        rbtn_files = ttk.Radiobutton(mode_frame, text="Files", variable=self.split_mode, value="files")
        rbtn_files.pack(side=tk.LEFT, anchor="w", padx=5)
        rbtn_sheets = ttk.Radiobutton(mode_frame, text="Sheets", variable=self.split_mode, value="sheets")
        rbtn_sheets.pack(side=tk.LEFT, anchor="w", padx=5)

        # Row 6: Split button
        btn_split = ttk.Button(frame, text="Split File", command=self.split_file_action)
        btn_split.pack(anchor="w", padx=10, pady=10)

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
        sheets = get_sheets_from_file(file_path, default_pwd)
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
        df = read_sheet_from_file(file_path, sheet, default_pwd)
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


def main():
    app = ExcelManagerGUI()
    app.mainloop()


if __name__ == "__main__":
    main()
