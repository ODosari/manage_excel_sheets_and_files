#!/usr/bin/env python3

# Name = "Manage Excel Sheets and Files"
# Version = "0.2"
# By = "Obaid Aldosari"
# GitHub = "https://github.com/ODosari/manage_excel_sheets_and_files"


import os
import pandas as pd
import time
import glob
from openpyxl import load_workbook
import argparse


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


def combine_excel_files(path):
    try:
        print("Searching for Excel files...")
        files = glob.glob(os.path.join(path, '*.xlsx'))
        if not files:
            print("No Excel files found in the directory.")
            return

        print("Found the following Excel files:")
        for i, file in enumerate(files, 1):
            print(f"{i}. {file}")
        print("Type 'all' to select all files.")

        selected_files_idx = input("Enter the numbers of the files to combine (separated by commas) or type 'all': ")
        selected_files = files if selected_files_idx.lower() == 'all' else [files[int(i) - 1] for i in selected_files_idx.split(',')]

        output_file = get_timestamped_filename(path, 'Combined', '.xlsx')
        choice = input("Combine into one sheet (O) or into one workbook with different sheets (W)?: ").lower()

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

        print(f'Files combined successfully. Output File: {output_file}')

    except Exception as e:
        print(f"An error occurred while combining files: {e}")


def choose_sheet_from_file(file):
    try:
        workbook = pd.ExcelFile(file, engine='openpyxl')
        sheet_names = workbook.sheet_names
        print(f"Available sheets in {file}: {sheet_names}")

        if len(sheet_names) == 1:  # If there's only one sheet, return it immediately
            print(f"Only one sheet available ('{sheet_names[0]}'), automatically selecting it.")
            return [(sheet_names[0], pd.read_excel(file, sheet_name=sheet_names[0], engine='openpyxl'))]
        else:
            print("Type 'all' to select all sheets.")
            chosen_sheet = input(f"Enter the name of the sheet to combine from {file} or type 'all': ").strip()
            if chosen_sheet.lower() == 'all':
                return [(sheet, pd.read_excel(file, sheet_name=sheet, engine='openpyxl')) for sheet in sheet_names]
            elif chosen_sheet not in sheet_names:
                print(f"Sheet '{chosen_sheet}' not found in {file}. Skipping this file.")
                return None
            else:
                return [(chosen_sheet, pd.read_excel(file, sheet_name=chosen_sheet, engine='openpyxl'))]
    except Exception as e:
        print(f"Error reading file {file}: {e}")
        return None


def split_excel_file(file):
    try:
        # Load the workbook and get the sheet names
        workbook = pd.ExcelFile(file, engine='openpyxl')
        sheet_names = workbook.sheet_names
        print(f"Available sheets: {sheet_names}")

        if len(sheet_names) == 1:
            chosen_sheet = sheet_names[0]
            print(f"Only one sheet ('{chosen_sheet}') available, automatically selected.")
        else:
            chosen_sheet = input("Enter the name of the sheet to split: ").strip()
            if chosen_sheet not in sheet_names:
                print(f"Sheet '{chosen_sheet}' not found in the workbook. Please try again.")
                return

        # Read the specified sheet
        df = pd.read_excel(file, sheet_name=chosen_sheet, engine='openpyxl')
        cols_name = df.columns.tolist()

        # List columns with indices
        print("Columns available for splitting:")
        for index, col in enumerate(cols_name, 1):
            print(f"{index}. {col}")

        column_index = int(input('Enter the index number of the column to split by: '))
        if column_index < 1 or column_index > len(cols_name):
            print("Invalid column index. Please try again.")
            return

        column_name = cols_name[column_index - 1]
        cols = df[column_name].unique()
        print(f'Your data will be split based on these values in "{column_name}": {", ".join(map(str, cols))}.')

        split_type = input('Split into different Sheets or Files (S/F): ').lower()
        if split_type == 'f':
            send_to_file(df, cols, column_name, file, chosen_sheet)
        elif split_type == 's':
            send_to_sheet(df, cols, column_name, file, chosen_sheet)
        else:
            print("Invalid choice. Please enter 'S' for sheets or 'F' for files.")
    except Exception as e:
        print(f"An error occurred while splitting the file: {e}")


def send_to_file(df, cols, column_name, file, sheet_name):
    directory = os.path.dirname(file)
    base_filename = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}"

    os.makedirs(directory, exist_ok=True)  # Ensure the directory exists without checking if it already exists

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
            filtered_df.to_excel(writer, sheet_name=sn, index=False, engine='openpyxl')

    print(f'Data split into sheets successfully. Output File: {output_file}')


def interactive_mode():
    print_help_message()
    while True:
        user_input = input("Enter your command: ").strip().lower()
        if user_input == 'q':
            break
        elif user_input.startswith(('c ', 's ')):
            operation, path = user_input.split(maxsplit=1)
            if operation == 'c':
                combine_excel_files(path)
            elif operation == 's':
                split_excel_file(path)
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
