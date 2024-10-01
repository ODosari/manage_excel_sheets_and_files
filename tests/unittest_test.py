import unittest
from unittest.mock import patch
import os
import pandas as pd
from test import (
    combine_excel_files,
    split_excel_file,
)

class TestManageExcelSheetsAndFiles(unittest.TestCase):

    def setUp(self):
        # Paths to your test files
        self.non_protected_file = '/Users/hacsec/PycharmProjects/manage_excel_sheets_and_files/TestData/HC.xlsx'
        self.protected_file = '/Users/hacsec/PycharmProjects/manage_excel_sheets_and_files/TestData/HC-P.xlsx'
        self.password = '1234'

        # Ensure the files exist
        self.assertTrue(os.path.isfile(self.non_protected_file), f"Non-protected file not found at {self.non_protected_file}")
        self.assertTrue(os.path.isfile(self.protected_file), f"Protected file not found at {self.protected_file}")

        # Create an output directory for the tests
        self.output_dir = '/Users/hacsec/PycharmProjects/manage_excel_sheets_and_files/TestData/output'
        os.makedirs(self.output_dir, exist_ok=True)

    def tearDown(self):
        # Clean up the output directory after tests
        for f in os.listdir(self.output_dir):
            os.remove(os.path.join(self.output_dir, f))
        os.rmdir(self.output_dir)

    @patch('builtins.input')
    def test_split_non_protected_file(self, mock_input):
        # Test splitting non-protected file based on column index 6
        mock_input.side_effect = [
            '1',  # Select sheet index 1
            'f',  # Split into files
            '6',  # Column index 6 for splitting
        ]
        split_excel_file(self.non_protected_file)
        # Check if split files are created
        output_files = [f for f in os.listdir(os.path.dirname(self.non_protected_file)) if f.endswith('.xlsx') and 'split' in f]
        self.assertTrue(len(output_files) > 0, "No split files created for non-protected file.")

    @patch('builtins.input')
    def test_split_protected_file(self, mock_input):
        # Test splitting protected file based on column index 6
        mock_input.side_effect = [
            self.password,  # Password for the protected file
            '1',            # Select sheet index 1
            'f',            # Split into files
            '6',            # Column index 6 for splitting
        ]
        split_excel_file(self.protected_file)
        # Check if split files are created
        output_files = [f for f in os.listdir(os.path.dirname(self.protected_file)) if f.endswith('.xlsx') and 'split' in f]
        self.assertTrue(len(output_files) > 0, "No split files created for protected file.")

    @patch('builtins.input')
    def test_split_into_sheets(self, mock_input):
        # Test splitting non-protected file into sheets based on column index 6
        mock_input.side_effect = [
            '1',  # Select sheet index 1
            's',  # Split into sheets
            '6',  # Column index 6 for splitting
        ]
        split_excel_file(self.non_protected_file)
        # Check if the output file is created
        output_files = [f for f in os.listdir(os.path.dirname(self.non_protected_file)) if f.endswith('.xlsx') and '_split_' in f]
        self.assertTrue(len(output_files) > 0, "No split file created for non-protected file into sheets.")

    @patch('builtins.input')
    def test_split_protected_file_into_sheets(self, mock_input):
        # Test splitting protected file into sheets based on column index 6
        mock_input.side_effect = [
            self.password,  # Password for the protected file
            '1',            # Select sheet index 1
            's',            # Split into sheets
            '6',            # Column index 6 for splitting
        ]
        split_excel_file(self.protected_file)
        # Check if the output file is created
        output_files = [f for f in os.listdir(os.path.dirname(self.protected_file)) if f.endswith('.xlsx') and '_split_' in f]
        self.assertTrue(len(output_files) > 0, "No split file created for protected file into sheets.")

    @patch('builtins.input')
    def test_combine_files(self, mock_input):
        # Test combining both files into one sheet, using column index 6
        mock_input.side_effect = [
            '1,2',          # Select both files by index
            'o',            # Combine into one sheet
            # For first file
            '1',            # Select sheet index 1
            '6',            # Column index 6 for combining
            # For second file (protected)
            self.password,  # Password for the protected file
            '1',            # Select sheet index 1
            '6',            # Column index 6 for combining
        ]
        test_dir = os.path.dirname(self.non_protected_file)
        combine_excel_files(test_dir)
        # Check if combined file is created
        output_files = [f for f in os.listdir(test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No combined file created.")

    @patch('builtins.input')
    def test_combine_files_into_workbook_with_sheets(self, mock_input):
        # Test combining both files into one workbook with different sheets
        mock_input.side_effect = [
            '1,2',          # Select both files by index
            'w',            # Combine into one workbook with sheets
            # For first file
            '1',            # Select sheet index 1
            # For second file (protected)
            self.password,  # Password for the protected file
            '1',            # Select sheet index 1
        ]
        test_dir = os.path.dirname(self.non_protected_file)
        combine_excel_files(test_dir)
        # Check if combined workbook is created
        output_files = [f for f in os.listdir(test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "No combined workbook created.")
        # Check the number of sheets in the combined workbook
        combined_file = os.path.join(test_dir, output_files[0])
        with pd.ExcelFile(combined_file) as xls:
            self.assertEqual(len(xls.sheet_names), 2, "Combined workbook should have 2 sheets.")

    @patch('builtins.input')
    def test_split_with_insufficient_columns(self, mock_input):
        # Test splitting a file where the sheet has fewer than 6 columns
        # Create a test file with fewer columns
        insufficient_columns_file = os.path.join(self.output_dir, 'insufficient_columns.xlsx')
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4], 'C': [5, 6]})
        df.to_excel(insufficient_columns_file, index=False)
        mock_input.side_effect = [
            '1',  # Select sheet index 1
            'f',  # Split into files
            '6',  # Column index 6 for splitting
        ]
        split_excel_file(insufficient_columns_file)
        # Since the sheet doesn't have enough columns, no files should be created
        output_files = [f for f in os.listdir(self.output_dir) if f.endswith('.xlsx') and 'split' in f]
        self.assertEqual(len(output_files), 0, "No split files should be created due to insufficient columns.")

    @patch('builtins.input')
    def test_combine_files_with_insufficient_columns(self, mock_input):
        # Test combining files where one has insufficient columns
        # Create a test file with fewer columns
        insufficient_columns_file = os.path.join(self.output_dir, 'insufficient_columns.xlsx')
        df = pd.DataFrame({'A': [1, 2], 'B': [3, 4], 'C': [5, 6]})
        df.to_excel(insufficient_columns_file, index=False)
        mock_input.side_effect = [
            '1,3',          # Select the non-protected file and the insufficient columns file
            'o',            # Combine into one sheet
            # For first file
            '1',            # Select sheet index 1
            '6',            # Column index 6 for combining
            # For second file
            '1',            # Select sheet index 1
            '6',            # Column index 6 for combining
        ]
        test_dir = os.path.dirname(self.non_protected_file)
        combine_excel_files(test_dir)
        # Check if the combined file is created with only the valid data
        output_files = [f for f in os.listdir(test_dir) if f.startswith('Combined_') and f.endswith('.xlsx')]
        self.assertTrue(len(output_files) > 0, "Combined file should be created even if one file has insufficient columns.")
        # Read the combined file and verify data
        combined_file = os.path.join(test_dir, output_files[0])
        df_combined = pd.read_excel(combined_file)
        # Since one file couldn't be merged, the combined data should only contain data from the valid file
        self.assertTrue(len(df_combined) > 0, "Combined DataFrame should contain data from the valid file.")

if __name__ == '__main__':
    unittest.main()
