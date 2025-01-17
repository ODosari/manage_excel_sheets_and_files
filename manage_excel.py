#!/usr/bin/env python3

# Name = "Manage Excel Sheets and Files"
# Version = "0.1"
# By = "Obaid Aldosari"
# GitHub = "https://github.com/ODosari/manage_excel_sheets_and_files"

import os
import pandas as pd
import time
import glob
from openpyxl import load_workbook
import argparse
import tempfile
import msoffcrypto
import io
import traceback
from contextlib import closing

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

def print_commands():
    print("\nAvailable commands:")
    print("C <path> - Combine all Excel files in <path> into a single file.")
    print("S <file> - Split an Excel file into multiple sheets or files based on a column.")
    print("Q - Quit the program.")

def get_timestamped_filename(base_path, prefix, extension):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f'{prefix}_{timestamp}{extension}')

def unprotect_excel_file(file):
    try:
        with open(file, 'rb') as f:
            office_file = msoffcrypto.OfficeFile(f)
            if office_file.is_encrypted():
                password = input(f"Enter password for {os.path.basename(file)}: ")
                decrypted = io.BytesIO()
                office_file.load_key(password=password)
                office_file.decrypt(decrypted)
                decrypted.seek(0)
                wb = load_workbook(decrypted, read_only=False, keep_vba=True, data_only=False, keep_links=False)
            else:
                # File is not encrypted
                wb = load_workbook(file, read_only=False, keep_vba=True, data_only=False, keep_links=False)

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
    except Exception as e:
        print(f"Failed to open {os.path.basename(file)}: {e}")
        return None

def combine_excel_files(path):
    try:
        print("Searching for Excel files...")
        files = glob.glob(os.path.join(path, '*.xlsx'))
        if not files:
            print("No Excel files found in the directory.")
            return

        print("Found the following Excel files:")
        for i, file in enumerate(files, 1):
            print(f"{i}. {os.path.basename(file)}")
        print("Type 'all' to select all files.")

        while True:
            selected_files_idx = input("Enter the numbers of the files to combine (separated by commas) or type 'all': ")
            if selected_files_idx.lower() == 'q':
                print("Operation cancelled by the user.")
                return
            if selected_files_idx.lower() == 'all':
                selected_files = files
                break
            else:
                indices = [i.strip() for i in selected_files_idx.split(',')]
                if all(idx.isdigit() and 1 <= int(idx) <= len(files) for idx in indices):
                    selected_files = [files[int(i) - 1] for i in indices]
                    break
                else:
                    print("Invalid input. Please enter valid file numbers or 'all'. Type 'Q' to cancel.")

        while True:
            choice = input("Combine into one sheet (O) or into one workbook with different sheets (W)? (Type 'Q' to cancel): ").lower()
            if choice == 'q':
                print("Operation cancelled by the user.")
                return
            elif choice in ['o', 'w']:
                break
            else:
                print("Invalid choice. Please enter 'O', 'W', or 'Q' to cancel.")

        output_file = get_timestamped_filename(path, 'Combined', '.xlsx')

        if choice == 'o':
            combined_df = pd.DataFrame()
            for file in selected_files:
                df_list = choose_sheet_from_file(file)
                if df_list is not None:
                    for sheet_name, df in df_list:
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
            combined_df.to_excel(output_file, sheet_name='Combined', engine='openpyxl', index=False)

        elif choice == 'w':
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                for file in selected_files:
                    df_list = choose_sheet_from_file(file)
                    if df_list is not None:
                        for sheet_name, df in df_list:
                            # Normalize the sheet name to create unique names
                            safe_sheet_name = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}".replace(' ', '_')[:31]
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

        print(f'Files combined successfully. Output File: {os.path.basename(output_file)}')

    except Exception as e:
        print(f"An error occurred while combining files: {e}")
        traceback.print_exc()
    finally:
        # Ensure that after combining, we return to the main menu
        pass  # The function will return naturally

def choose_sheet_from_file(file):
    unprotected_file = None
    try:
        # For .xlsx files
        unprotected_file = unprotect_excel_file(file)
        if unprotected_file is None:
            return None

        while True:
            # Use contextlib.closing to ensure workbook is closed
            with closing(pd.ExcelFile(unprotected_file, engine='openpyxl')) as workbook:
                sheet_names = workbook.sheet_names
                print(f"\nAvailable sheets in {os.path.basename(file)}:")
                for idx, sheet_name in enumerate(sheet_names, start=1):
                    print(f"{idx}. {sheet_name}")

                if len(sheet_names) == 1:
                    print(f"Only one sheet available ('{sheet_names[0]}'), automatically selecting it.")
                    df = pd.read_excel(unprotected_file, sheet_name=sheet_names[0], engine='openpyxl')
                    return [(sheet_names[0], df)]
                else:
                    print("\nType 'all' to select all sheets or 'Q' to skip this file.")
                    chosen_input = input(f"Enter the index numbers of the sheets to select from {os.path.basename(file)} (separated by commas), 'all', or 'Q' to skip: ").strip()
                    if chosen_input.lower() == 'q':
                        print(f"Skipping file {os.path.basename(file)}.")
                        return None
                    elif chosen_input.lower() == 'all':
                        df_list = []
                        for sheet in sheet_names:
                            df = pd.read_excel(unprotected_file, sheet_name=sheet, engine='openpyxl')
                            df_list.append((sheet, df))
                        return df_list
                    else:
                        indices = chosen_input.split(',')
                        df_list = []
                        for idx in indices:
                            idx = idx.strip()
                            if idx.isdigit():
                                index = int(idx)
                                if 1 <= index <= len(sheet_names):
                                    chosen_sheet = sheet_names[index - 1]
                                    df = pd.read_excel(unprotected_file, sheet_name=chosen_sheet, engine='openpyxl')
                                    df_list.append((chosen_sheet, df))
                                else:
                                    print(f"Invalid index number {index}.")
                            else:
                                print(f"Invalid input '{idx}'. Please enter numbers only.")
                        if df_list:
                            return df_list
                        else:
                            print("No valid sheets selected. Please try again or type 'Q' to skip this file.")
                            continue
    except Exception as e:
        print(f"Error reading file {os.path.basename(file)}: {e}")
        traceback.print_exc()
        return None
    finally:
        if unprotected_file:
            os.unlink(unprotected_file)

def split_excel_file(file):
    unprotected_file = None
    try:
        unprotected_file = unprotect_excel_file(file)
        if unprotected_file is None:
            return

        while True:
            with closing(pd.ExcelFile(unprotected_file, engine='openpyxl')) as workbook:
                sheet_names = workbook.sheet_names
                print(f"\nAvailable sheets in {os.path.basename(file)}:")
                for idx, sheet_name in enumerate(sheet_names, start=1):
                    print(f"{idx}. {sheet_name}")

                if len(sheet_names) == 1:
                    chosen_sheet = sheet_names[0]
                    print(f"Only one sheet ('{chosen_sheet}') available, automatically selected.")
                else:
                    print("Type 'Q' to skip this file.")
                    chosen_input = input(f"Enter the index number of the sheet to split from {os.path.basename(file)}: ").strip()
                    if chosen_input.lower() == 'q':
                        print(f"Skipping file {os.path.basename(file)}.")
                        return
                    if not chosen_input.isdigit():
                        print("Invalid input. Please enter a valid index number or 'Q' to skip.")
                        continue
                    index = int(chosen_input)
                    if 1 <= index <= len(sheet_names):
                        chosen_sheet = sheet_names[index - 1]
                    else:
                        print("Invalid index number. Please try again or type 'Q' to skip.")
                        continue

                df = pd.read_excel(unprotected_file, sheet_name=chosen_sheet, engine='openpyxl')
                cols_name = df.columns.tolist()

                print("\nColumns available for splitting:")
                for index, col in enumerate(cols_name, 1):
                    print(f"{index}. {col}")

                while True:
                    column_input = input("Enter the index number of the column to split by (or type 'Q' to skip): ").strip()
                    if column_input.lower() == 'q':
                        print("Skipping splitting operation.")
                        return
                    if not column_input.isdigit():
                        print("Invalid input. Please enter a number or 'Q' to skip.")
                        continue
                    column_index = int(column_input)
                    if 1 <= column_index <= len(cols_name):
                        break
                    else:
                        print("Invalid column index. Please try again or type 'Q' to skip.")

                column_name = cols_name[column_index - 1]
                cols = df[column_name].unique()
                print(f'Your data will be split based on these values in "{column_name}": {", ".join(map(str, cols))}.')

                while True:
                    split_type = input("Split into different Sheets or Files (S/F)? (Type 'Q' to skip): ").lower()
                    if split_type == 'q':
                        print("Skipping splitting operation.")
                        return
                    elif split_type == 'f':
                        send_to_file(df, cols, column_name, file, chosen_sheet)
                        break
                    elif split_type == 's':
                        send_to_sheet(df, cols, column_name, file, chosen_sheet)
                        break
                    else:
                        print("Invalid choice. Please enter 'S', 'F', or 'Q' to skip.")

                # After successful splitting, break out of the while loop
                break

        # Return to main menu after operation
    except Exception as e:
        print(f"An error occurred while splitting the file: {e}")
        traceback.print_exc()
    finally:
        if unprotected_file:
            os.unlink(unprotected_file)

def send_to_file(df, cols, column_name, file, sheet_name):
    directory = os.path.dirname(file)
    base_filename = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}"

    os.makedirs(directory, exist_ok=True)  # Ensure the directory exists

    for value in cols:
        output_file = get_timestamped_filename(directory, f'{base_filename}_{column_name}_{value}', '.xlsx')
        df[df[column_name] == value].to_excel(output_file, sheet_name=str(value), engine='openpyxl', index=False)

    print('Data split into files successfully.')

def send_to_sheet(df, cols, column_name, file, sheet_name):
    directory = os.path.dirname(file)
    output_file = get_timestamped_filename(directory, f'{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}_split', '.xlsx')

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for value in cols:
            sn = str(value)[:31]  # Excel sheet name limit is 31 characters
            filtered_df = df[df[column_name] == value]
            filtered_df.to_excel(writer, sheet_name=sn, index=False)

    print(f'Data split into sheets successfully. Output File: {os.path.basename(output_file)}')

def interactive_mode():
    print_help_message()
    while True:
        print_commands()
        user_input = input("Enter your command: ").strip()
        if user_input.lower() == 'q':
            break
        elif user_input.lower() == 'help':
            print_help_message()
            continue
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
