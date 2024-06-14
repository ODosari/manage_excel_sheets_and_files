#!/usr/bin/env python3

# Name = "Manage Excel Sheets and Files"
# Version = "1.7"
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

# Configure logging
logging.basicConfig(filename='excel_manager.log', level=logging.INFO,
                    format='%(asctime)s - %(levelname)s - %(message)s')


# ANSI escape sequences for colors
class Colors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def print_help_message():
    help_message = f"""
{Colors.HEADER}Welcome to Manage Excel Sheets and Files Utility!{Colors.ENDC}
{Colors.OKCYAN}This utility allows you to:{Colors.ENDC}
  - Combine multiple Excel files from a directory into one
  - Split a single Excel file into multiple sheets or files based on a specific column

{Colors.BOLD}Commands:{Colors.ENDC}
  {Colors.OKGREEN}C <directory>{Colors.ENDC} - Combine all Excel files from the specified directory into a single file.
  {Colors.OKGREEN}S <file>{Colors.ENDC} - Split an Excel file into multiple sheets or files based on a specific column.
  {Colors.OKGREEN}Q{Colors.ENDC} - Quit the program.
"""
    print(help_message)


def get_timestamped_filename(base_path, prefix, extension):
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    return os.path.join(base_path, f'{prefix}_{timestamp}{extension}')


def create_password_template():
    template_path = get_timestamped_filename('.', PASSWORD_TEMPLATE_BASE, '.xlsx')
    df = DataFrame(columns=['FileName', 'Password'])
    df.to_excel(template_path, index=False, engine='openpyxl')
    print(f"{Colors.OKGREEN}Password template created: {template_path}{Colors.ENDC}")
    logging.info(f"Password template created: {template_path}")
    return template_path


def read_passwords_from_template(template_path):
    if not os.path.exists(template_path):
        print(f"{Colors.FAIL}{template_path} does not exist.{Colors.ENDC}")
        logging.error(f"{template_path} does not exist.")
        return {}
    df = read_excel(template_path, engine='openpyxl')
    return Series(df.Password.values, index=df.FileName).to_dict()


def read_protected_excel(file_path, password):
    decrypted = io.BytesIO()
    with open(file_path, 'rb') as f:
        try:
            office_file = OfficeFile(f)
            office_file.load_key(password=password)
            office_file.decrypt(decrypted)
        except Exception as e:
            print(f"{Colors.FAIL}Error decrypting file {file_path}: {e}{Colors.ENDC}")
            logging.error(f"Error decrypting file {file_path}: {e}")
            return None
    decrypted.seek(0)
    try:
        return read_excel(decrypted, sheet_name=None, engine='openpyxl')
    except Exception as e:
        print(f"{Colors.FAIL}Error reading decrypted file {file_path}: {e}{Colors.ENDC}")
        logging.error(f"Error reading decrypted file {file_path}: {e}")
        return None


def read_excel_file(file_path, password=None):
    try:
        if password:
            return read_protected_excel(file_path, password)
        else:
            return read_excel(file_path, sheet_name=None, engine='openpyxl')
    except Exception as e:
        print(f"{Colors.FAIL}Error reading file {file_path}: {e}{Colors.ENDC}")
        logging.error(f"Error reading file {file_path}: {e}")
        return None


def get_file_password(file, password_dict):
    return password_dict.get(file, None)


def handle_file_selection(files):
    print(f"{Colors.OKCYAN}Found the following Excel files:{Colors.ENDC}")
    for i, file in enumerate(files, 1):
        print(f"{Colors.OKBLUE}{i}. {file}{Colors.ENDC}")
    print(f"{Colors.BOLD}Type 'all' to select all files.{Colors.ENDC}")
    selected_files_idx = input(
        f"Enter the numbers of the files to combine (separated by commas) or type 'all': ").strip() or 'all'
    return files if selected_files_idx.lower() == 'all' else [files[int(i) - 1] for i in selected_files_idx.split(',')]


def choose_sheet_from_file(file, password=None):
    try:
        excel_data = read_excel_file(file, password)
        if excel_data is None:
            return None

        sheet_names = list(excel_data.keys())
        print(f"{Colors.OKCYAN}Available sheets in {file}: {sheet_names}{Colors.ENDC}")

        if len(sheet_names) == 1:
            print(
                f"{Colors.OKGREEN}Only one sheet available ('{sheet_names[0]}'), automatically selecting it.{Colors.ENDC}")
            return [(sheet_names[0], excel_data[sheet_names[0]])]
        else:
            print(f"{Colors.BOLD}Type 'all' to select all sheets.{Colors.ENDC}")
            chosen_sheet = input(f"Enter the name of the sheet to combine from {file} or type 'all': ").strip() or 'all'
            if chosen_sheet.lower() == 'all':
                return [(sheet, excel_data[sheet]) for sheet in sheet_names]
            elif chosen_sheet not in sheet_names:
                print(f"{Colors.WARNING}Sheet '{chosen_sheet}' not found in {file}. Skipping this file.{Colors.ENDC}")
                logging.warning(f"Sheet '{chosen_sheet}' not found in {file}. Skipping this file.")
                return None
            else:
                return [(chosen_sheet, excel_data[chosen_sheet])]
    except Exception as e:
        print(f"{Colors.FAIL}Error reading file {file}: {e}{Colors.ENDC}")
        logging.error(f"Error reading file {file}: {e}")
        return None


def combine_excel_files(path, password_dict):
    try:
        if os.path.isfile(path):
            print(f"{Colors.FAIL}Please specify a directory, not a file, for combining Excel files.{Colors.ENDC}")
            return

        print(f"{Colors.OKCYAN}Searching for Excel files...{Colors.ENDC}")
        files = glob.glob(os.path.join(path, '*.xlsx'))

        if not files:
            print(f"{Colors.FAIL}No Excel files found in the specified directory.{Colors.ENDC}")
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

        print(f"{Colors.OKGREEN}Files combined successfully. Output File: {output_file}{Colors.ENDC}")
        logging.info(f'Files combined successfully. Output File: {output_file}')

    except Exception as e:
        print(f"{Colors.FAIL}An error occurred while combining files: {e}{Colors.ENDC}")
        logging.error(f"An error occurred while combining files: {e}")


def split_excel_file(file, password_dict):
    try:
        if not os.path.isfile(file):
            print(f"{Colors.FAIL}Please specify a valid Excel file for splitting.{Colors.ENDC}")
            return

        password = get_file_password(file, password_dict)
        excel_data = read_excel_file(file, password)
        if excel_data is None:
            return

        sheet_names = list(excel_data.keys())
        print(f"{Colors.OKCYAN}Available sheets: {sheet_names}{Colors.ENDC}")

        chosen_sheet = sheet_names[0] if len(sheet_names) == 1 else input(
            "Enter the name of the sheet to split: ").strip() or sheet_names[0]
        if chosen_sheet not in sheet_names:
            print(f"{Colors.WARNING}Sheet '{chosen_sheet}' not found in the workbook. Please try again.{Colors.ENDC}")
            logging.warning(f"Sheet '{chosen_sheet}' not found in the workbook. Please try again.")
            return

        df = excel_data[chosen_sheet]
        cols_name = df.columns.tolist()
        print(f"{Colors.OKCYAN}Columns available for splitting:{Colors.ENDC}")
        for index, col in enumerate(cols_name, 1):
            print(f"{Colors.OKBLUE}{index}. {col}{Colors.ENDC}")

        column_index = int(input('Enter the index number of the column to split by: ').strip() or '1')
        if column_index < 1 or column_index > len(cols_name):
            print(f"{Colors.FAIL}Invalid column index. Please try again.{Colors.ENDC}")
            logging.error("Invalid column index.")
            return

        column_name = cols_name[column_index - 1]
        cols = df[column_name].unique()
        print(
            f"{Colors.OKCYAN}Your data will be split based on these values in '{column_name}': {', '.join(map(str, cols))}.{Colors.ENDC}")

        split_type = input('Split into different Sheets or Files (S/F): ').strip().lower() or 's'
        if split_type == 'f':
            save_split_data(df, cols, column_name, file, sheet_name=chosen_sheet, mode='file')
        elif split_type == 's':
            save_split_data(df, cols, column_name, file, sheet_name=chosen_sheet, mode='sheet')
        else:
            print(f"{Colors.FAIL}Invalid choice. Please enter 'S' for sheets or 'F' for files.{Colors.ENDC}")
            logging.error("Invalid choice.")
    except Exception as e:
        print(f"{Colors.FAIL}An error occurred while splitting the file: {e}{Colors.ENDC}")
        logging.error(f"An error occurred while splitting the file: {e}")


def save_split_data(df, cols, column_name, file, sheet_name, mode='file'):
    directory = os.path.dirname(file)
    base_filename = f"{os.path.splitext(os.path.basename(file))[0]}_{sheet_name}"

    os.makedirs(directory, exist_ok=True)

    if mode == 'file':
        for value in cols:
            output_file = get_timestamped_filename(directory, f'{base_filename}_{column_name}_{value}', '.xlsx')
            df[df[column_name] == value].to_excel(output_file, sheet_name=str(value), engine='openpyxl', index=False)
        print(f"{Colors.OKGREEN}Data split into files successfully.{Colors.ENDC}")
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
        print(f"{Colors.OKGREEN}Data split into sheets successfully. Output File: {output_file}{Colors.ENDC}")
        logging.info(f'Data split into sheets successfully. Output File: {output_file}')


def interactive_mode():
    print_help_message()

    while True:
        command = input(
            f"{Colors.BOLD}Enter your command (C <directory> for combine, S <file> for split, Q to quit):{Colors.ENDC} ").strip().lower()
        if command == 'q':
            break
        elif command.startswith('c '):
            _, path = command.split(maxsplit=1)
            if not os.path.isdir(path):
                print(
                    f"{Colors.FAIL}The path provided is not a directory. Please provide a directory containing Excel files.{Colors.ENDC}")
                continue
            password_dict = get_passwords()
            combine_excel_files(path, password_dict)
        elif command.startswith('s '):
            _, file = command.split(maxsplit=1)
            if not os.path.isfile(file):
                print(f"{Colors.FAIL}The path provided is not a file. Please provide a valid Excel file.{Colors.ENDC}")
                continue
            password_dict = get_passwords()
            split_excel_file(file, password_dict)
        else:
            print(f"{Colors.FAIL}Invalid command. Please try again.{Colors.ENDC}")
            print_help_message()


def ask_yes_no(prompt):
    while True:
        response = input(f"{Colors.BOLD}{prompt} (y/n): {Colors.ENDC}").strip().lower()
        if response in ('y', 'n'):
            return response == 'y'
        print(f"{Colors.FAIL}Invalid input. Please enter 'y' or 'n'.{Colors.ENDC}")


def get_passwords():
    password_dict = {}
    if ask_yes_no("Are the Excel files password protected?"):
        if ask_yes_no("Is there a single password for all files?"):
            single_password = input("Enter the password: ").strip()
            password_dict = {file: single_password for file in glob.glob("*.xlsx")}
        else:
            new_template = create_password_template()
            print(
                f"{Colors.OKCYAN}Fill in the passwords in the new template {new_template} and run the program again.{Colors.ENDC}")
            exit()
    return password_dict


def parse_arguments():
    parser = argparse.ArgumentParser(description="Manage Excel Sheets and Files")
    parser.add_argument('-c', '--combine', help='Path to directory containing Excel files to combine')
    parser.add_argument('-s', '--split', help='Excel file to split into multiple sheets or files')
    return parser.parse_args()


def main():
    args = parse_arguments()

    if args.combine:
        if not os.path.isdir(args.combine):
            print(
                f"{Colors.FAIL}The path provided is not a directory. Please provide a directory containing Excel files.{Colors.ENDC}")
            return
        password_dict = get_passwords()
        combine_excel_files(args.combine, password_dict)
    elif args.split:
        if not os.path.isfile(args.split):
            print(f"{Colors.FAIL}The path provided is not a file. Please provide a valid Excel file.{Colors.ENDC}")
            return
        password_dict = get_passwords()
        split_excel_file(args.split, password_dict)
    else:
        interactive_mode()


if __name__ == "__main__":
    main()
