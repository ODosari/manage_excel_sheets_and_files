# Manage Excel Sheets and Files

**Version**: 1.0  
**Author**: Obaid Aldosari  
**GitHub**: [ODosari](https://github.com/ODosari/manage_excel_sheets_and_files)

## Table of Contents

1. [Introduction](#introduction)
2. [Features](#features)
3. [Installation](#installation)
4. [Usage](#usage)
   - [Interactive Mode](#interactive-mode)
   - [Command-Line Mode](#command-line-mode)
5. [Commands](#commands)
6. [Examples](#examples)
   - [Combining Excel Files](#combining-excel-files)
   - [Splitting an Excel File](#splitting-an-excel-file)
7. [Password Handling](#password-handling)
   - [Single Password](#single-password)
   - [Multiple Passwords](#multiple-passwords)
8. [Contributing](#contributing)
9. [License](#license)

## Introduction

`Manage Excel Sheets and Files` is a utility script designed to streamline the management of Excel files. It allows you to combine multiple Excel files from a directory into one file or split a single Excel file into multiple sheets or files based on a specific column. The script supports password-protected files and offers an intuitive command-line interface.

## Features

- **Combine Excel Files**: Merge multiple Excel files from a directory into one file.
- **Split Excel Files**: Split a single Excel file into multiple files or sheets based on a specified column.
- **Password Protection**: Handle password-protected Excel files with a single password or different passwords for each file.
- **Interactive Mode**: User-friendly interactive mode to guide through combining and splitting tasks.
- **Logging**: Logs activities and errors to `excel_manager.log`.

## Installation

### Prerequisites

- **Python 3.6 or higher** is required.
- Install the required Python packages:

```bash
pip install pandas msoffcrypto-tool openpyxl
```

### Clone the Repository

```bash
git clone https://github.com/ODosari/manage_excel_sheets_and_files.git
cd manage_excel_sheets_and_files
```

## Usage

You can use the script in both interactive and command-line modes.

### Interactive Mode

Run the script without any arguments to enter the interactive mode:

```bash
python manage_excel_sheets_and_files.py
```

### Command-Line Mode

Use the following commands to combine or split Excel files directly from the command line:

#### Combine Excel Files

Combine all Excel files from a specified directory into a single file:

```bash
python manage_excel_sheets_and_files.py -c <directory_path>
```

#### Split Excel File

Split a single Excel file into multiple files or sheets based on a specific column:

```bash
python manage_excel_sheets_and_files.py -s <file_path>
```

## Commands

### Interactive Mode Commands

- **`C <directory>`**: Combine all Excel files from the specified directory into a single file.
- **`S <file>`**: Split an Excel file into multiple sheets or files based on a specific column.
- **`Q`**: Quit the program.

### Command-Line Arguments

- **`-c`, `--combine` `<directory_path>`**: Path to the directory containing Excel files to combine.
- **`-s`, `--split` `<file_path>`**: Path to the Excel file to split into multiple sheets or files.

## Examples

### Combining Excel Files

```bash
python manage_excel_sheets_and_files.py -c ./data/excel_files/
```

This command will combine all Excel files in the `./data/excel_files/` directory into a single file.

### Splitting an Excel File

```bash
python manage_excel_sheets_and_files.py -s ./data/excel_files/sample.xlsx
```

This command will split the `sample.xlsx` file based on a specified column into multiple files or sheets.

## Password Handling

The script supports handling password-protected Excel files. You can provide a single password for all files or use a template for multiple passwords.

### Single Password

If all files use the same password, you will be prompted to enter it once. The script will use this password for all files.

### Multiple Passwords

If different files have different passwords, the script will create a password template:

1. **Create Password Template**: The script generates a password template Excel file where you can specify passwords for each file.
2. **Fill in Passwords**: Fill in the passwords in the generated template.
3. **Run Again**: Rerun the script with the completed password template.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request with any improvements or bug fixes.

1. Fork the repository.
2. Create your feature branch (`git checkout -b feature/new-feature`).
3. Commit your changes (`git commit -m 'Add new feature'`).
4. Push to the branch (`git push origin feature/new-feature`).
5. Open a pull request.

## License

Not specified yet