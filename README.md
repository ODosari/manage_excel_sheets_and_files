# Manage Excel Sheets and Files

## Description
Manage Excel Sheets and Files is a Python utility designed to simplify handling Excel files. It allows users to combine multiple Excel files into a single file or split a single Excel file into multiple sheets or files based on a specific column.

## Features
- **Combine Excel Files**: Merge multiple `.xlsx` files either into one sheet or multiple sheets within one workbook.
- **Split Excel Files**: Divide a single Excel file into multiple sheets or separate files based on the values of a specified column.

## Installation

### Prerequisites
- Python 3.x

### Setting Up a Virtual Environment

A virtual environment is a tool that helps to keep dependencies required by different projects separate. To set up a virtual environment for this project, follow these steps:

1. **Install Virtualenv** (if not already installed):
   ```bash
   pip install virtualenv
   ```

2. **Create a Virtual Environment**:
   Navigate to the project directory and run:
   ```bash
   virtualenv venv
   ```

3. **Activate the Virtual Environment**:
   - On Windows:
     ```bash
     .\venv\Scripts\activate
     ```
   - On Unix or MacOS:
     ```bash
     source venv/bin/activate
     ```

   After activation, your command line prompt will change to indicate that you are now working inside the virtual environment.

4. **Install Dependencies**:
   Within the activated virtual environment, install the required libraries:
   ```bash
   pip install pandas openpyxl
   ```

### Clone the Repository
After setting up the virtual environment, you can clone the repository:
```bash
git clone https://github.com/ODosari/manage_excel_sheets_and_files.git
cd manage_excel_sheets_and_files
```

## Usage

### Interactive Mode
Run the script in interactive mode for a guided experience:
```bash
./manage_excel.py
```

### Command Line Arguments
Alternatively, use command line arguments:

- To combine files:
  ```bash
  ./manage_excel.py -c <path-to-directory>
  ```
- To split a file:
  ```bash
  ./manage_excel.py -s <path-to-file>
  ```

## Deactivating the Virtual Environment
Once you're done, you can deactivate the virtual environment by running:
```bash
deactivate
```

## Contributing
Contributions to enhance Manage Excel Sheets and Files are welcome.

