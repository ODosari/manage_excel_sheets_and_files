import unittest
from unittest.mock import patch, MagicMock
import os
import pandas as pd
from io import StringIO

# Import the functions from experimental_tests.py
import experimental_tests

class TestExcelManagement(unittest.TestCase):
    def setUp(self):
        # Paths to your test files
        self.test_dir = '/Users/hacsec/PycharmProjects/manage_excel_sheets_and_files/TestData'
        self.non_protected_file = os.path.join(self.test_dir, 'HC.xlsx')
        self.protected_file = os.path.join(self.test_dir, 'HC-P.xlsx')
        self.csv_file = os.path.join(self.test_dir, 'test_file.csv')
        self.password = '1234'

        # Ensure the files exist
        self.assertTrue(os.path.isfile(self.non_protected_file), f"Non-protected file not found at {self.non_protected_file}")
        self.assertTrue(os.path.isfile(self.protected_file), f"Protected file not found at {self.protected_file}")

        # Create an output directory for the tests
        self.output_dir = os.path.join(self.test_dir, 'output')
        os.makedirs(self.output_dir, exist_ok=True)

        # Create a CSV file for testing with at least 6 columns
        df = pd.DataFrame({
            'A': [1, 2, 3],
            'B': ['x', 'y', 'z'],
            'C': ['alpha', 'beta', 'gamma'],
            'D': [4, 5, 6],
            'E': ['m', 'n', 'o'],
            'F': ['p', 'q', 'r']
        })
        df.to_csv(self.csv_file, index=False)

        # Mock the is_file_encrypted function to return True for the protected file
        experimental_tests.is_file_encrypted = MagicMock(side_effect=lambda x: x == self.protected_file)

    def tearDown(self):
        # Clean up the output directory after tests
        for f in os.listdir(self.output_dir):
            os.remove(os.path.join(self.output_dir, f))
        os.rmdir(self.output_dir)

        # Remove CSV file
        if os.path.exists(self.csv_file):
            os.remove(self.csv_file)

    @patch('builtins.input')
    @patch('experimental_tests.glob.glob')
    def test_combine_excel_files_to_one_sheet(self, mock_glob, mock_input):
        # Mock glob to return our test files
        mock_glob.return_value = [self.non_protected_file, self.protected_file, self.csv_file]

        # Mock user inputs
        mock_input.side_effect = [
            self.test_dir,    # Directory path
            'all',            # Select all files
            self.password,    # Password for the protected file
            'o',              # Combine into one sheet
            '',               # Output directory (blank for default)
        ]

        # Run the combine_excel_files function
        experimental_tests.combine_excel_files()

        # Check if combined file is created in the test directory
        output_files = [f for f in os.listdir(self.test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No combined file created.")

        # Optionally, check the contents of the combined file
        combined_file = os.path.join(self.test_dir, output_files[0])
        combined_df = pd.read_excel(combined_file)
        self.assertTrue(not combined_df.empty, "Combined DataFrame should not be empty.")

    @patch('builtins.input')
    @patch('experimental_tests.glob.glob')
    def test_combine_excel_files_to_workbook_with_sheets(self, mock_glob, mock_input):
        # Mock glob to return our test files
        mock_glob.return_value = [self.non_protected_file, self.protected_file, self.csv_file]

        # Mock user inputs
        mock_input.side_effect = [
            self.test_dir,     # Directory path
            'all',             # Select all files
            self.password,     # Password for the protected file
            'w',               # Combine into workbook with sheets
            '',                # Output directory (blank for default)
        ]

        # Run the combine_excel_files function
        experimental_tests.combine_excel_files()

        # Check if combined workbook is created
        output_files = [f for f in os.listdir(self.test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No combined workbook created.")

        # Check the number of sheets in the combined workbook
        combined_file = os.path.join(self.test_dir, output_files[0])
        with pd.ExcelFile(combined_file) as xls:
            # We expect at least 3 sheets (from the non-protected file, protected file, and CSV file)
            self.assertTrue(len(xls.sheet_names) >= 3, "Combined workbook should have at least 3 sheets.")

    @patch('builtins.input')
    def test_split_non_protected_file_to_files(self, mock_input):
        # Mock user inputs
        # Since there is only one sheet, the script auto-selects it, so we don't need to provide input for sheet selection
        mock_input.side_effect = [
            self.non_protected_file,  # File path
            '',                       # Output directory (blank for default)
            '6',                      # Column index 6 for splitting
            'f',                      # Split into files
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split files are created
        output_files = [f for f in os.listdir(os.path.dirname(self.non_protected_file)) if f.endswith('.xlsx') and f != os.path.basename(self.non_protected_file)]
        self.assertTrue(len(output_files) > 0, "No split files created for non-protected file.")

    @patch('builtins.input')
    def test_split_non_protected_file_to_sheets(self, mock_input):
        # Mock user inputs
        mock_input.side_effect = [
            self.non_protected_file,  # File path
            '',                       # Output directory (blank for default)
            '6',                      # Column index 6 for splitting
            's',                      # Split into sheets
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split file is created
        output_files = [f for f in os.listdir(os.path.dirname(self.non_protected_file)) if f.endswith('.xlsx') and 'split' in f]
        self.assertTrue(len(output_files) > 0, "No split file created for non-protected file into sheets.")

        # Verify that the output file contains multiple sheets
        output_file = os.path.join(os.path.dirname(self.non_protected_file), output_files[0])
        with pd.ExcelFile(output_file) as xls:
            self.assertTrue(len(xls.sheet_names) > 1, "Output file should contain multiple sheets.")

    @patch('builtins.input')
    def test_split_protected_file_to_files(self, mock_input):
        # Mock user inputs
        mock_input.side_effect = [
            self.protected_file,  # File path
            '',                   # Output directory (blank for default)
            self.password,        # Password for the protected file
            '6',                  # Column index 6 for splitting
            'f',                  # Split into files
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split files are created
        output_files = [f for f in os.listdir(os.path.dirname(self.protected_file)) if f.endswith('.xlsx') and f != os.path.basename(self.protected_file)]
        self.assertTrue(len(output_files) > 0, "No split files created for protected file.")

    @patch('builtins.input')
    def test_split_protected_file_to_sheets(self, mock_input):
        # Mock user inputs
        mock_input.side_effect = [
            self.protected_file,  # File path
            '',                   # Output directory (blank for default)
            self.password,        # Password for the protected file
            '6',                  # Column index 6 for splitting
            's',                  # Split into sheets
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split file is created
        output_files = [f for f in os.listdir(os.path.dirname(self.protected_file)) if f.endswith('.xlsx') and 'split' in f]
        self.assertTrue(len(output_files) > 0, "No split file created for protected file into sheets.")

        # Verify that the output file contains multiple sheets
        output_file = os.path.join(os.path.dirname(self.protected_file), output_files[0])
        with pd.ExcelFile(output_file) as xls:
            self.assertTrue(len(xls.sheet_names) > 1, "Output file should contain multiple sheets.")

    @patch('builtins.input')
    def test_output_directory_option(self, mock_input):
        # Mock user inputs for splitting a file with specified output directory
        mock_input.side_effect = [
            self.non_protected_file,  # File path
            self.output_dir,          # Output directory
            '6',                      # Column index 6 for splitting
            'f',                      # Split into files
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split files are created in the output directory
        output_files = [f for f in os.listdir(self.output_dir) if f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No split files created in the specified output directory.")

    @patch('builtins.input')
    def test_handling_csv_files_in_splitting(self, mock_input):
        # Mock user inputs for splitting a CSV file
        # For CSV files, the sheet selection is not applicable
        mock_input.side_effect = [
            self.csv_file,      # File path
            '',                 # Output directory (blank for default)
            '6',                # Column index 6 for splitting
            'f',                # Split into files
        ]

        # Run the split_excel_file function
        experimental_tests.split_excel_file()

        # Check if split files are created
        output_files = [f for f in os.listdir(os.path.dirname(self.csv_file)) if f.endswith('.xlsx') and 'split' in f]
        self.assertTrue(len(output_files) > 0, "No split files created from CSV.")

    @patch('builtins.input')
    @patch('experimental_tests.glob.glob')
    def test_handling_csv_files_in_combining(self, mock_glob, mock_input):
        # Mock glob to return our CSV file
        mock_glob.return_value = [self.csv_file]

        # Mock user inputs for combining CSV files
        mock_input.side_effect = [
            self.test_dir,     # Directory path
            'all',             # Select all files
            'o',               # Combine into one sheet
            '',                # Output directory (blank for default)
        ]

        # Run the combine_excel_files function
        experimental_tests.combine_excel_files()

        # Check if combined file is created
        output_files = [f for f in os.listdir(self.test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No combined file created from CSV.")

    @patch('builtins.input')
    def test_invalid_file_path(self, mock_input):
        # Mock user inputs with an invalid file path
        invalid_file = '/path/to/nonexistent/file.xlsx'
        mock_input.side_effect = [
            invalid_file,  # File path
        ]

        # Capture the printed output
        with patch('sys.stdout', new_callable=StringIO) as fake_out:
            experimental_tests.split_excel_file()
            output = fake_out.getvalue()
            self.assertIn("File not found. Please try again.", output)

    @patch('builtins.input')
    def test_invalid_directory_path(self, mock_input):
        # Mock user inputs with an invalid directory path
        invalid_dir = '/path/to/nonexistent/directory'
        mock_input.side_effect = [
            invalid_dir,  # Directory path
        ]

        # Capture the printed output
        with patch('sys.stdout', new_callable=StringIO) as fake_out:
            experimental_tests.combine_excel_files()
            output = fake_out.getvalue()
            self.assertIn("Invalid directory path. Please try again.", output)

    @patch('builtins.input')
    def test_user_cancels_operation(self, mock_input):
        # Mock user inputs where the user cancels the operation
        mock_input.side_effect = [
            'q',  # User inputs 'q' to cancel at directory path prompt
        ]

        # Capture the printed output
        with patch('sys.stdout', new_callable=StringIO) as fake_out:
            experimental_tests.combine_excel_files()
            output = fake_out.getvalue()
            self.assertIn("Operation cancelled. Returning to main menu.", output)

    @patch('builtins.input')
    def test_split_file_no_unique_values(self, mock_input):
        # Create a DataFrame with only one unique value in the selected column
        df = pd.DataFrame({
            'A': [1, 1, 1, 1, 1, 1],
            'B': ['x', 'x', 'x', 'x', 'x', 'x'],
            'C': ['alpha', 'alpha', 'alpha', 'alpha', 'alpha', 'alpha'],
            'D': [2, 2, 2, 2, 2, 2],
            'E': ['y', 'y', 'y', 'y', 'y', 'y'],
            'F': ['beta', 'beta', 'beta', 'beta', 'beta', 'beta']
        })
        no_unique_file = os.path.join(self.test_dir, 'no_unique_values.xlsx')
        df.to_excel(no_unique_file, index=False)

        # Mock user inputs
        mock_input.side_effect = [
            no_unique_file,   # File path
            '',               # Output directory
            '6',              # Column index 6 (column 'F') to split by
            'f',              # Split into files
        ]

        # Capture the printed output
        with patch('sys.stdout', new_callable=StringIO) as fake_out:
            experimental_tests.split_excel_file()
            output = fake_out.getvalue()
            self.assertIn("Not enough unique values found in column 'F'. Cannot split.", output)

        # Clean up
        if os.path.exists(no_unique_file):
            os.remove(no_unique_file)

    @patch('builtins.input')
    def test_split_file_no_columns(self, mock_input):
        # Create an Excel file with no columns
        df = pd.DataFrame()
        no_columns_file = os.path.join(self.test_dir, 'no_columns.xlsx')
        df.to_excel(no_columns_file, index=False)

        # Mock user inputs
        mock_input.side_effect = [
            no_columns_file,  # File path
            '',               # Output directory
        ]

        # Capture the printed output
        with patch('sys.stdout', new_callable=StringIO) as fake_out:
            experimental_tests.split_excel_file()
            output = fake_out.getvalue()
            self.assertIn("No columns found in the selected sheet.", output)

        # Clean up
        if os.path.exists(no_columns_file):
            os.remove(no_columns_file)

if __name__ == '__main__':
    unittest.main()
