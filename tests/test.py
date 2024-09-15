#!/usr/bin/env python3

# Name = "Manage Excel Sheets and Files"
# Version = "1.0"
# By = "Obaid Aldosari"
# GitHub = "https://github.com/ODosari/manage_excel_sheets_and_files"

import os
import sys
import pandas as pd
import time
import glob
from openpyxl import load_workbook
import argparse
import tempfile
import msoffcrypto
import io
import traceback
import logging

# Configure logging
logging.basicConfig(filename='excel_manager.log', level=logging.INFO,
                    format='%(asctime)s:%(levelname)s:%(message)s')

def print_help_message():
    print("################################################################################")
    print("Welcome to Manage Excel Sheets and Files Utility!")
    print("This utility allows you to combine multiple Excel files into one,")
    print("or split a single Excel file into multiple sheets or files based on a specific column.")
    print("\nCommands:")
    print("C <path> - Combine all Excel files in <path> into a single file.")
    print("S <file> - Split an Excel file into multiple sheets or files based on a column.")
    print("Q - Quit the program.")
    print("################################################################################")

def get_timestamped_filename(base_path, prefix, extension):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f'{prefix}_{timestamp}{extension}')

def unprotect_excel_file(file):
    try:
        with open(file, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            if office_file.is_encrypted():
                password = input(f"Enter password for {file}: ")
                decrypted = io.BytesIO()
                try:
                    office_file.load_key(password=password)
                    office_file.decrypt(decrypted)
                except Exception as e:
                    print(f"Incorrect password for {file}.")
                    logging.error(f"Incorrect password for {file}: {e}")
                    return None
                decrypted.seek(0)
                wb = load_workbook(decrypted, read_only=False, keep_vba=True, data_only=False, keep_links=False)
            else:
                # File is not encrypted
                wb = load_workbook(file, read_only=False, keep_vba=True, data_only=False, keep_links=False)
    except Exception as e:
        print(f"Failed to open {file}: {e}")
        logging.error(f"Failed to open {file}: {e}")
        return None

    # Unprotect workbook
    wb.security = None  # Remove workbook-level protection

    # Unprotect sheets
    for sheet in wb.worksheets:
        sheet.protection.enabled = False
        sheet.protection.sheet = False

    # Save to a temporary file
    temp = tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False)
    wb.save(temp.name)
    temp.close()
    return temp.name

def combine_excel_files(path):
    try:
        if not os.path.isdir(path):
            print(f"The path {path} does not exist or is not a directory.")
            return
        print("Searching for Excel files...")
        files = glob.glob(os.path.join(path, '*.xlsx'))
        if not files:
            print("No Excel files found in the directory.")
            return

        print("Found the following Excel files:")
        for i, file in enumerate(files, 1):
            print(f"{i}. {file}")
        print("Type 'all' to select all files.")

        selected_files_idx = input("Enter the numbers of the files to combine (separated by commas) or type 'all': ").strip()
        if selected_files_idx.lower() == 'all':
            selected_files = files
        else:
            try:
                selected_indices = [int(i.strip()) - 1 for i in selected_files_idx.split(',')]
                selected_files = [files[i] for i in selected_indices]
            except (ValueError, IndexError):
                print("Invalid selection. Please enter valid file numbers separated by commas.")
                return

        if not selected_files:
            print("No files selected.")
            return

        default_output_dir = path
        output_dir = input(f"Enter output directory (default: {default_output_dir}): ").strip() or default_output_dir
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        output_file_default = get_timestamped_filename(output_dir, 'Combined', '.xlsx')
        output_file = input(f"Enter output file name (default: {output_file_default}): ").strip() or output_file_default

        choice = input("Combine into one sheet (O) or into one workbook with different sheets (W)? [O/W]: ").strip().lower()
        while choice not in ('o', 'w'):
            choice = input("Invalid choice. Please enter 'O' for one sheet or 'W' for workbook: ").strip().lower()

        include_header = input("Include header row from each file? [Y/N]: ").strip().lower()
        while include_header not in ('y', 'n'):
            include_header = input("Invalid input. Please enter 'Y' or 'N': ").strip().lower()
        header = True if include_header == 'y' else False

        remove_duplicates = input("Remove duplicate rows after combining? [Y/N]: ").strip().lower()
        while remove_duplicates not in ('y', 'n'):
            remove_duplicates = input("Invalid input. Please enter 'Y' or 'N': ").strip().lower()
        remove_dupes = True if remove_duplicates == 'y' else False

        if choice == 'o':
            combined_df = pd.DataFrame()
            for file in selected_files:
                df_list = choose_sheet_from_file(file)
                if df_list is not None:
                    for sheet_name, df in df_list:
                        if not header:
                            df = df.iloc[1:]
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
            if remove_dupes:
                combined_df.drop_duplicates(inplace=True)
            combined_df.to_excel(output_file, sheet_name='Combined', engine='openpyxl', index=False)
        elif choice == 'w':
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for file in selected_files:
                    df_list = choose_sheet_from_file(file)
                    if df_list is not None:
                        for sheet_name, df in df_list:
                            if not header:
                                df = df.iloc[1:]
                            # Normalize the sheet name to create unique names
                            safe_sheet_name = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}".replace(' ', '_')[:31]
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
        print(f'Files combined successfully. Output File: {output_file}')
        logging.info(f'Files combined successfully. Output File: {output_file}')
    except Exception as e:
        print(f"An error occurred while combining files: {e}")
        logging.error(f"An error occurred while combining files: {e}")
        traceback.print_exc()

def choose_sheet_from_file(file):
    unprotected_file = None
    try:
        unprotected_file = unprotect_excel_file(file)
        if unprotected_file is None:
            return None
        workbook = pd.ExcelFile(unprotected_file, engine='openpyxl')
        sheet_names = workbook.sheet_names
        print(f"Available sheets in {file}: {sheet_names}")

        if len(sheet_names) == 1:
            print(f"Only one sheet available ('{sheet_names[0]}'), automatically selecting it.")
            df = pd.read_excel(unprotected_file, sheet_name=sheet_names[0], engine='openpyxl')
            return [(sheet_names[0], df)]
        else:
            print("Type 'all' to select all sheets.")
            chosen_sheet = input(f"Enter the name(s) of the sheet(s) to combine from {file} (comma-separated) or type 'all': ").strip()
            if chosen_sheet.lower() == 'all':
                df_list = []
                for sheet in sheet_names:
                    df = pd.read_excel(unprotected_file, sheet_name=sheet, engine='openpyxl')
                    df_list.append((sheet, df))
                return df_list
            else:
                chosen_sheets = [s.strip() for s in chosen_sheet.split(',')]
                df_list = []
                for sheet in chosen_sheets:
                    if sheet in sheet_names:
                        df = pd.read_excel(unprotected_file, sheet_name=sheet, engine='openpyxl')
                        df_list.append((sheet, df))
                    else:
                        print(f"Sheet '{sheet}' not found in {file}. Skipping this sheet.")
                return df_list if df_list else None
    except Exception as e:
        print(f"Error reading file {file}: {e}")
        logging.error(f"Error reading file {file}: {e}")
        traceback.print_exc()
        return None
    finally:
        if unprotected_file:
            os.unlink(unprotected_file)

def split_excel_file(file):
    unprotected_file = None
    try:
        if not os.path.isfile(file):
            print(f"The file {file} does not exist.")
            return
        unprotected_file = unprotect_excel_file(file)
        if unprotected_file is None:
            return
        workbook = pd.ExcelFile(unprotected_file, engine='openpyxl')
        sheet_names = workbook.sheet_names
        print(f"Available sheets: {sheet_names}")

        if len(sheet_names) == 1:
            chosen_sheet = sheet_names[0]
            print(f"Only one sheet ('{chosen_sheet}') available, automatically selected.")
        else:
            chosen_sheet = input("Enter the name of the sheet to split: ").strip()
            while chosen_sheet not in sheet_names:
                print(f"Sheet '{chosen_sheet}' not found in the workbook.")
                chosen_sheet = input("Please enter a valid sheet name: ").strip()

        # Read the specified sheet
        df = pd.read_excel(unprotected_file, sheet_name=chosen_sheet, engine='openpyxl')
        cols_name = df.columns.tolist()

        # List columns with indices
        print("Columns available for splitting:")
        for index, col in enumerate(cols_name, 1):
            print(f"{index}. {col}")

        while True:
            try:
                column_index = int(input('Enter the index number of the column to split by: '))
                if 1 <= column_index <= len(cols_name):
                    break
                else:
                    print("Invalid column index. Please enter a number from the list.")
            except ValueError:
                print("Invalid input. Please enter a number.")

        column_name = cols_name[column_index - 1]
        unique_values = df[column_name].dropna().unique()
        if unique_values.size == 0:
            print(f"No unique values found in column '{column_name}' to split by.")
            return
        print(f'Your data will be split based on these values in "{column_name}": {", ".join(map(str, unique_values))}.')

        split_type = input('Split into different Sheets or Files (S/F)? [S/F]: ').strip().lower()
        while split_type not in ('s', 'f'):
            split_type = input("Invalid choice. Please enter 'S' for sheets or 'F' for files: ").strip().lower()

        default_output_dir = os.path.dirname(file)
        output_dir = input(f"Enter output directory (default: {default_output_dir}): ").strip() or default_output_dir
        if not os.path.isdir(output_dir):
            os.makedirs(output_dir, exist_ok=True)

        filter_data = input("Would you like to filter the data before splitting? [Y/N]: ").strip().lower()
        while filter_data not in ('y', 'n'):
            filter_data = input("Invalid input. Please enter 'Y' or 'N': ").strip().lower()

        if filter_data == 'y':
            filter_column = input("Enter the column name to filter on: ").strip()
            while filter_column not in cols_name:
                print(f"Column '{filter_column}' not found.")
                filter_column = input("Please enter a valid column name: ").strip()
            filter_value = input(f"Enter the value to filter '{filter_column}' by: ").strip()
            df = df[df[filter_column] == filter_value]

        if split_type == 'f':
            send_to_file(df, unique_values, column_name, file, chosen_sheet, output_dir)
        elif split_type == 's':
            send_to_sheet(df, unique_values, column_name, file, chosen_sheet, output_dir)
    except Exception as e:
        print(f"An error occurred while splitting the file: {e}")
        logging.error(f"An error occurred while splitting the file: {e}")
        traceback.print_exc()
    finally:
        if unprotected_file:
            os.unlink(unprotected_file)

def send_to_file(df, unique_values, column_name, file, sheet_name, output_dir):
    base_filename = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}"

    for value in unique_values:
        output_file = get_timestamped_filename(output_dir, f'{base_filename}_{column_name}_{value}', '.xlsx')
        subset_df = df[df[column_name] == value]
        subset_df.to_excel(output_file, sheet_name=str(value), engine='openpyxl', index=False)

    print('Data split into files successfully.')
    logging.info('Data split into files successfully.')

def send_to_sheet(df, unique_values, column_name, file, sheet_name, output_dir):
    output_file_default = get_timestamped_filename(output_dir, f'{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}_split', '.xlsx')
    output_file = input(f"Enter output file name (default: {output_file_default}): ").strip() or output_file_default

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for value in unique_values:
            sn = str(value)[:31]  # Excel sheet name limit is 31 characters
            filtered_df = df[df[column_name] == value]
            filtered_df.to_excel(writer, sheet_name=sn, index=False)

    print(f'Data split into sheets successfully. Output File: {output_file}')
    logging.info(f'Data split into sheets successfully. Output File: {output_file}')

def interactive_mode():
    print_help_message()
    while True:
        user_input = input("Enter your command: ").strip()
        if user_input.lower() == 'q':
            print("Exiting the program.")
            logging.info("User exited the program.")
            break
        elif user_input.startswith(('C ', 'c ', 'S ', 's ')):
            parts = user_input.strip().split(maxsplit=1)
            if len(parts) < 2:
                print("Please provide a path after the command.")
                continue
            operation, path = parts
            operation = operation.lower()
            if operation == 'c':
                combine_excel_files(path)
            elif operation == 's':
                split_excel_file(path)
            else:
                print("Invalid command. Type 'Q' to quit or see above for command usage.")
        else:
            print("Invalid command. Type 'Q' to quit or see above for command usage.")

def parse_arguments():
    parser = argparse.ArgumentParser(description="Manage Excel Sheets and Files")
    parser.add_argument('-c', '--combine', help='Path to combine Excel files')
    parser.add_argument('-s', '--split', help='File to split into multiple sheets or files')
    return parser.parse_args()

def main():
    args = parse_arguments()
    if args.combine:
        combine_excel_files(args.combine)
    elif args.split:
        split_excel_file(args.split)
    else:
        interactive_mode()

if __name__ == "__main__":
    main()
