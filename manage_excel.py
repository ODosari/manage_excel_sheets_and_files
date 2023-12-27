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


def get_valid_file_path():
    """
    Prompts the user to enter an operation and file path. Validates the input.
    Returns the operation and file path.
    """
    while True:
        user_input = input("Enter 'C' to combine files in a directory, or 'S' to split a file: ").strip()
        if user_input.lower().startswith(('c ', 's ')):
            operation, path = user_input.split(maxsplit=1)
            operation = operation.lower()

            if operation == 'c' or (operation == 's' and os.path.isfile(path)):
                return operation, path
            print("Invalid file path. Please enter a valid path.")
        else:
            print("Invalid input format. Please start with 'C' or 'S' followed by the path.")


def get_timestamped_filename(base_path, prefix, extension):
    """
    Generates a filename with a timestamp.
    """
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f'{prefix}_{timestamp}{extension}')


def combine_excel_files(path):
    """
    Combines Excel files in the given path into a single file.
    Allows user to choose between combining into one sheet or separate sheets.
    """
    files = glob.glob(os.path.join(path, '*.xlsx'))
    if not files:
        print("No Excel files found in the directory.")
        return

    choice = input("Combine into one sheet (O) or into one workbook with different sheets (W)?: ").lower()
    output_file = get_timestamped_filename(path, 'combined', '.xlsx')

    if choice == 'o':
        combined_df = pd.concat([pd.read_excel(f) for f in files], ignore_index=True)
        combined_df.to_excel(output_file, sheet_name='combined', index=False)
        print(f'Files combined into one sheet successfully. Output File: {output_file}')
    elif choice == 'w':
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for f in files:
                df = pd.read_excel(f)
                sheet_name = os.path.splitext(os.path.basename(f))[0][:31]  # Excel sheet name limit is 31 characters
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        print(f'Files combined into one workbook with different sheets successfully. Output File: {output_file}')
    else:
        print("Invalid choice. Please enter 'O' for one sheet or 'W' for workbook with different sheets.")


def split_excel_file(file):
    """
    Splits an Excel file based on a specified column into different sheets or files.
    """
    df = pd.read_excel(file)
    cols_name = df.columns.tolist()
    print(f'Columns name(s): {cols_name}')
    column_name = input('Type the name of Column to split by: ').strip()

    if column_name not in cols_name:
        print(f"The Column name '{column_name}' is not found. Please try again.")
        return

    cols = df[column_name].unique()
    print(f'Your data will be split based on these values: {", ".join(map(str, cols))}.')

    choice = input('Ready to Proceed? (Y/N): ').lower()
    if choice != 'y':
        print("Operation cancelled.")
        return

    split_type = input('Split into different Sheets or Files (S/F): ').lower()
    if split_type == 'f':
        send_to_file(df, cols, column_name, file)
    elif split_type == 's':
        send_to_sheet(df, cols, column_name, file)
    else:
        print("Invalid choice. Please enter 'S' for sheets or 'F' for files.")


def send_to_file(df, cols, column_name, file):
    """
    Splits the DataFrame into separate files based on the column values.
    """
    pth = os.path.dirname(file)
    for value in cols:
        output_file = get_timestamped_filename(pth, f'{column_name}_{value}', '.xlsx')
        df[df[column_name] == value].to_excel(output_file, sheet_name=str(value), index=False)
    print('Data split into files successfully.')


def send_to_sheet(df, cols, column_name, file):
    """
    Splits the DataFrame into separate sheets in a single file based on the column values.
    """
    pth, ext = os.path.splitext(file)
    output_file = get_timestamped_filename(pth, 'split_workbook', ext)

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for value in cols:
            sheet_name = str(value)[:31]  # Excel sheet name limit is 31 characters
            filtered_df = df[df[column_name] == value]
            filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print('Data split into sheets successfully.')


def main():
    """
    Main function to run the Excel file utility.
    """
    print("-------------------------------------------------------------------------------------")
    print("Welcome to the Manage Excel Utility!")
    print("This utility allows you to combine multiple Excel files into one,")
    print("or split a single Excel file into multiple sheets or files based on a specific column.")
    print("-------------------------------------------------------------------------------------\n")

    operation, path = get_valid_file_path()

    if operation == 'c':
        combine_excel_files(path)
    elif operation == 's':
        split_excel_file(path)


if __name__ == "__main__":
    main()
