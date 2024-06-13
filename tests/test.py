#!/usr/bin/env python3

# Name = "Manage Excel Sheets and Files"
# Version = "0.9"
# By = "Obaid Aldosari"
# GitHub = "https://github.com/ODosari/manage_excel_sheets_and_files"

import os
import pandas as pd
import time
import glob
import argparse
import io
import logging
from msoffcrypto import OfficeFile
from pandas import ExcelWriter, DataFrame, read_excel, Series

PASSWORD_TEMPLATE_BASE = 'passwords_template'

logging.basicConfig(filename='excel_manager.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


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


def create_password_template():
    template_path = get_timestamped_filename('.', PASSWORD_TEMPLATE_BASE, '.xlsx')
    df = DataFrame(columns=['FileName', 'Password'])
    df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"Password template created: {template_path}")
    logging.info(f"Password template created: {template_path}")
    return template_path


def read_passwords_from_template(template_path):
    if not os.path.exists(template_path):
        print(f"{template_path} does not exist.")
        logging.error(f"{template_path} does not exist.")
        return {}
    df = read_excel(template_path, engine='openpyxl')
    return Series(df.Password.values, index=df.FileName).to_dict()


def read_protected_excel(file_path, password):
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as f:
        office_file = OfficeFile(f)
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
    decrypted.seek(0)
    return read_excel(decrypted, sheet_name=None, engine='openpyxl')


def read_excel_file(file_path, password=None):
    try:
        if password:
            return read_protected_excel(file_path, password)
        else:
            return read_excel(file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        logging.error(f"Error reading file {file_path}: {e}")
        return None


def get_file_password(file, password_dict):
    return password_dict.get(file, None)


def handle_file_selection(files):
    print("Found the following Excel files:")
    for i, file in enumerate(files, 1):
        print(f"{i}. {file}")
    print("Type 'all' to select all files.")
    selected_files_idx = input(
        "Enter the numbers of the files to combine (separated by commas) or type 'all': ").strip() or 'all'
    return files if selected_files_idx.lower() == 'all' else [files[int(i) - 1] for i in selected_files_idx.split(',')]


def choose_sheet_from_file(file, password=None):
    try:
        excel_data = read_excel_file(file, password)
        if excel_data is None:
            return None

        sheet_names = list(excel_data.keys())
        print(f"Available sheets in {file}: {sheet_names}")

        if len(sheet_names) == 1:
            print(f"Only one sheet available ('{sheet_names[0]}'), automatically selecting it.")
            return [(sheet_names[0], excel_data[sheet_names[0]])]
        else:
            print("Type 'all' to select all sheets.")
            chosen_sheet = input(f"Enter the name of the sheet to combine from {file} or type 'all': ").strip() or 'all'
            if chosen_sheet.lower() == 'all':
                return [(sheet, excel_data[sheet]) for sheet in sheet_names]
            elif chosen_sheet not in sheet_names:
                print(f"Sheet '{chosen_sheet}' not found in {file}. Skipping this file.")
                logging.warning(f"Sheet '{chosen_sheet}' not found in {file}. Skipping this file.")
                return None
            else:
                return [(chosen_sheet, excel_data[chosen_sheet])]
    except Exception as e:
        print(f"Error reading file {file}: {e}")
        logging.error(f"Error reading file {file}: {e}")
        return None


def combine_excel_files(path, password_dict):
    try:
        print("Searching for Excel files...")
        files = glob.glob(os.path.join(path, '*.xlsx'))
        if not files:
            print("No Excel files found in the directory.")
            return

        selected_files = handle_file_selection(files)
        output_file = get_timestamped_filename(path, 'Combined', '.xlsx')
        choice = input(
            "Combine into one sheet (O) or into one workbook with different sheets (W)?: ").strip().lower() or 'o'

        if choice == 'o':
            combined_df = DataFrame()
            for file in selected_files:
                password = get_file_password(file, password_dict)
                df_list = choose_sheet_from_file(file, password)
                if df_list is not None:
                    for _, df in df_list:
                        combined_df = pd.concat([combined_df, df], ignore_index=True)
            combined_df.to_excel(output_file, sheet_name='Combined', engine='openpyxl', index=False)

        elif choice == 'w':
            with ExcelWriter(output_file, engine='openpyxl') as writer:
                for file in selected_files:
                    password = get_file_password(file, password_dict)
                    df_list = choose_sheet_from_file(file, password)
                    if df_list is not None:
                        for sheet_name, df in df_list:
                            safe_sheet_name = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}".replace(' ',
                                                                                                                    '_')[
                                              :31]
                            df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

        print(f'Files combined successfully. Output File: {output_file}')
        logging.info(f'Files combined successfully. Output File: {output_file}')

    except Exception as e:
        print(f"An error occurred while combining files: {e}")
        logging.error(f"An error occurred while combining files: {e}")


def split_excel_file(file, password_dict):
    try:
        password = get_file_password(file, password_dict)
        excel_data = read_excel_file(file, password)
        if excel_data is None:
            return

        sheet_names = list(excel_data.keys())
        print(f"Available sheets: {sheet_names}")

        chosen_sheet = sheet_names[0] if len(sheet_names) == 1 else input(
            "Enter the name of the sheet to split: ").strip() or sheet_names[0]
        if chosen_sheet not in sheet_names:
            print(f"Sheet '{chosen_sheet}' not found in the workbook. Please try again.")
            logging.warning(f"Sheet '{chosen_sheet}' not found in the workbook. Please try again.")
            return

        df = excel_data[chosen_sheet]
        cols_name = df.columns.tolist()
        print("Columns available for splitting:")
        for index, col in enumerate(cols_name, 1):
            print(f"{index}. {col}")

        column_index = int(input('Enter the index number of the column to split by: ').strip() or '1')
        if column_index < 1 or column_index > len(cols_name):
            print("Invalid column index. Please try again.")
            logging.error("Invalid column index.")
            return

        column_name = cols_name[column_index - 1]
        cols = df[column_name].unique()
        print(f'Your data will be split based on these values in "{column_name}": {", ".join(map(str, cols))}.')

        split_type = input('Split into different Sheets or Files (S/F): ').strip().lower() or 's'
        if split_type == 'f':
            save_split_data(df, cols, column_name, file, sheet_name=chosen_sheet, mode='file')
        elif split_type == 's':
            save_split_data(df, cols, column_name, file, sheet_name=chosen_sheet, mode='sheet')
        else:
            print("Invalid choice. Please enter 'S' for sheets or 'F' for files.")
            logging.error("Invalid choice.")
    except Exception as e:
        print(f"An error occurred while splitting the file: {e}")
        logging.error(f"An error occurred while splitting the file: {e}")


def save_split_data(df, cols, column_name, file, sheet_name, mode='file'):
    directory = os.path.dirname(file)
    base_filename = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}"

    os.makedirs(directory, exist_ok=True)

    if mode == 'file':
        for value in cols:
            output_file = get_timestamped_filename(directory, f'{base_filename}_{column_name}_{value}', '.xlsx')
            df[df[column_name] == value].to_excel(output_file, sheet_name=str(value), engine='openpyxl', index=False)
        print('Data split into files successfully.')
        logging.info('Data split into files successfully.')
    elif mode == 'sheet':
        output_file = get_timestamped_filename(directory,
                                               f'{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}_split',
                                               '.xlsx')
        with ExcelWriter(output_file, engine='openpyxl') as writer:
            for value in cols:
                sn = str(value)[:31]
                filtered_df = df[df[column_name] == value]
                filtered_df.to_excel(writer, sheet_name=sn, index=False)
        print(f'Data split into sheets successfully. Output File: {output_file}')
        logging.info(f'Data split into sheets successfully. Output File: {output_file}')


def interactive_mode():
    print_help_message()
    password_dict = {}
    if input("Are the Excel files password protected? (y/n): ").strip().lower() == 'y':
        password_template_exists = os.path.exists(PASSWORD_TEMPLATE_BASE + '.xlsx')
        use_existing = password_template_exists and input(
            f"An existing password file {PASSWORD_TEMPLATE_BASE}.xlsx was found. Do you want to use it? (y/n): ").strip().lower() == 'y'

        if use_existing:
            password_dict = read_passwords_from_template(PASSWORD_TEMPLATE_BASE + '.xlsx')
        else:
            if password_template_exists:
                print(f"The existing password file {PASSWORD_TEMPLATE_BASE}.xlsx will not be used.")
            if input("Do you have an existing password file? (y/n): ").strip().lower() == 'y':
                existing_file = input("Enter the path to your password file: ").strip()
                if os.path.exists(existing_file):
                    password_dict = read_passwords_from_template(existing_file)
                else:
                    print(f"The file {existing_file} does not exist.")
                    logging.error(f"The file {existing_file} does not exist.")
                    return
            else:
                new_template = create_password_template()
                print(f"Fill in the passwords in the new template {new_template} and run the program again.")
                return

    while True:
        command = input("Enter your command (C <path> for combine, S <file> for split, Q to quit): ").strip().lower()
        if command == 'q':
            break
        elif command.startswith('c '):
            _, path = command.split(maxsplit=1)
            combine_excel_files(path, password_dict)
        elif command.startswith('s '):
            _, file = command.split(maxsplit=1)
            split_excel_file(file, password_dict)
        else:
            print("Invalid command. Please try again.")
            print_help_message()


def parse_arguments():
    parser = argparse.ArgumentParser(description="Manage Excel Sheets and Files")
    parser.add_argument('-c', '--combine', help='Path to combine Excel files')
    parser.add_argument('-s', '--split', help='File to split into multiple sheets or files')
    return parser.parse_args()


def main():
    args = parse_arguments()

    if args.combine:
        password_dict = interactive_mode_get_passwords()
        combine_excel_files(args.combine, password_dict)
    elif args.split:
        password_dict = interactive_mode_get_passwords()
        split_excel_file(args.split, password_dict)
    else:
        interactive_mode()


def interactive_mode_get_passwords():
    password_dict = {}
    if input("Are the Excel files password protected? (y/n): ").strip().lower() == 'y':
        password_template_exists = os.path.exists(PASSWORD_TEMPLATE_BASE + '.xlsx')
        use_existing = password_template_exists and input(
            f"An existing password file {PASSWORD_TEMPLATE_BASE}.xlsx was found. Do you want to use it? (y/n): ").strip().lower() == 'y'

        if use_existing:
            password_dict = read_passwords_from_template(PASSWORD_TEMPLATE_BASE + '.xlsx')
        else:
            if password_template_exists:
                print(f"The existing password file {PASSWORD_TEMPLATE_BASE}.xlsx will not be used.")
            if input("Do you have an existing password file? (y/n): ").strip().lower() == 'y':
                existing_file = input("Enter the path to your password file: ").strip()
                if os.path.exists(existing_file):
                    password_dict = read_passwords_from_template(existing_file)
                else:
                    print(f"The file {existing_file} does not exist.")
                    logging.error(f"The file {existing_file} does not exist.")
                    return {}
            else:
                new_template = create_password_template()
                print(f"Fill in the passwords in the new template {new_template} and run the program again.")
                return {}
    return password_dict


if __name__ == "__main__":
    main()
