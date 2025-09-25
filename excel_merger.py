#!/usr/bin/env python3
"""
Excel File Merger Script

This script merges all Excel files (.xlsx and .xls) in a specified folder 
into a single Excel workbook.

Requirements:
- pandas
- openpyxl 
- xlrd

Author: Generated for Excel file merging task
"""

import os
import sys
import argparse
import pandas as pd
from pathlib import Path
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


class ExcelMerger:
    """Class to handle merging of Excel files."""
    
    def __init__(self):
        self.supported_extensions = ['.xlsx', '.xls']
        self.processed_files = []
        self.failed_files = []
    
    def find_excel_files(self, folder_path):
        """Find all Excel files in the specified folder."""
        excel_files = []
        folder = Path(folder_path)
        
        if not folder.exists():
            raise FileNotFoundError(f"Folder '{folder_path}' does not exist.")
        
        if not folder.is_dir():
            raise NotADirectoryError(f"'{folder_path}' is not a directory.")
        
        for file_path in folder.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in self.supported_extensions:
                excel_files.append(file_path)
        
        if not excel_files:
            print(f"No Excel files found in '{folder_path}'")
            return []
        
        print(f"Found {len(excel_files)} Excel file(s): {[f.name for f in excel_files]}")
        return excel_files
    
    def read_excel_file(self, file_path):
        """Read an Excel file and return a dictionary of DataFrames."""
        sheets_data = {}
        file_name = file_path.stem  # filename without extension
        
        try:
            # Try to read with openpyxl first (for .xlsx files)
            if file_path.suffix.lower() == '.xlsx':
                excel_file = pd.ExcelFile(file_path, engine='openpyxl')
            else:
                # Use xlrd for .xls files
                excel_file = pd.ExcelFile(file_path, engine='xlrd')
            
            for sheet_name in excel_file.sheet_names:
                try:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    if not df.empty:
                        # Create unique sheet name with source file info
                        new_sheet_name = f"{file_name}_{sheet_name}"
                        sheets_data[new_sheet_name] = df
                    else:
                        print(f"Warning: Sheet '{sheet_name}' in '{file_path.name}' is empty, skipping.")
                except Exception as e:
                    print(f"Error reading sheet '{sheet_name}' from '{file_path.name}': {e}")
                    continue
            
            excel_file.close()
            self.processed_files.append(file_path.name)
            print(f"Successfully processed '{file_path.name}' with {len(sheets_data)} sheet(s)")
            
        except Exception as e:
            print(f"Error reading file '{file_path.name}': {e}")
            self.failed_files.append(file_path.name)
            return {}
        
        return sheets_data
    
    def merge_to_single_sheet(self, all_sheets_data, output_path):
        """Merge all data into a single sheet."""
        if not all_sheets_data:
            print("No data to merge.")
            return False
        
        merged_data = []
        
        for sheet_name, df in all_sheets_data.items():
            # Add a column to identify the source
            df_copy = df.copy()
            df_copy.insert(0, 'Source_File_Sheet', sheet_name)
            merged_data.append(df_copy)
        
        try:
            # Concatenate all DataFrames
            final_df = pd.concat(merged_data, ignore_index=True, sort=False)
            
            # Write to Excel
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                final_df.to_excel(writer, sheet_name='Merged_Data', index=False)
            
            print(f"Successfully merged all data into single sheet: '{output_path}'")
            print(f"Total rows: {len(final_df)}, Total columns: {len(final_df.columns)}")
            return True
            
        except Exception as e:
            print(f"Error creating merged file: {e}")
            return False
    
    def merge_to_separate_sheets(self, all_sheets_data, output_path):
        """Keep data in separate sheets."""
        if not all_sheets_data:
            print("No data to merge.")
            return False
        
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for sheet_name, df in all_sheets_data.items():
                    # Truncate sheet name if too long (Excel limit is 31 characters)
                    safe_sheet_name = sheet_name[:31] if len(sheet_name) > 31 else sheet_name
                    df.to_excel(writer, sheet_name=safe_sheet_name, index=False)
            
            print(f"Successfully merged data into separate sheets: '{output_path}'")
            print(f"Total sheets: {len(all_sheets_data)}")
            return True
            
        except Exception as e:
            print(f"Error creating merged file: {e}")
            return False
    
    def merge_excel_files(self, input_folder, output_file, merge_mode='separate'):
        """Main method to merge Excel files."""
        print(f"Starting Excel file merge...")
        print(f"Input folder: {input_folder}")
        print(f"Output file: {output_file}")
        print(f"Merge mode: {merge_mode}")
        print("-" * 50)
        
        # Find Excel files
        excel_files = self.find_excel_files(input_folder)
        if not excel_files:
            return False
        
        # Read all Excel files
        all_sheets_data = {}
        for file_path in excel_files:
            sheets_data = self.read_excel_file(file_path)
            all_sheets_data.update(sheets_data)
        
        if not all_sheets_data:
            print("No valid data found in any Excel files.")
            return False
        
        # Merge based on mode
        if merge_mode == 'single':
            success = self.merge_to_single_sheet(all_sheets_data, output_file)
        else:
            success = self.merge_to_separate_sheets(all_sheets_data, output_file)
        
        # Print summary
        print("\n" + "=" * 50)
        print("MERGE SUMMARY")
        print("=" * 50)
        print(f"Files processed successfully: {len(self.processed_files)}")
        for file in self.processed_files:
            print(f"  ✓ {file}")
        
        if self.failed_files:
            print(f"\nFiles that failed to process: {len(self.failed_files)}")
            for file in self.failed_files:
                print(f"  ✗ {file}")
        
        print(f"\nTotal sheets merged: {len(all_sheets_data)}")
        print(f"Output file: {output_file}")
        print("=" * 50)
        
        return success


def main():
    """Main function to handle command line arguments and run the merger."""
    parser = argparse.ArgumentParser(
        description="Merge all Excel files (.xlsx and .xls) in a folder into a single Excel file.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s -i ./excel_files -o merged_output.xlsx
  %(prog)s -i /path/to/excel/files -o combined.xlsx -m single
  %(prog)s --input-folder ./data --output-file result.xlsx --merge-mode separate

Merge modes:
  separate: Keep each sheet in separate tabs (default)
  single:   Combine all data into one sheet with source identification
        """
    )
    
    parser.add_argument(
        '-i', '--input-folder',
        required=True,
        help='Path to folder containing Excel files to merge'
    )
    
    parser.add_argument(
        '-o', '--output-file',
        required=True,
        help='Name of the output Excel file (e.g., merged_output.xlsx)'
    )
    
    parser.add_argument(
        '-m', '--merge-mode',
        choices=['separate', 'single'],
        default='separate',
        help='How to merge the data: "separate" keeps sheets separate, "single" combines all into one sheet (default: separate)'
    )
    
    args = parser.parse_args()
    
    # Validate input folder
    if not os.path.exists(args.input_folder):
        print(f"Error: Input folder '{args.input_folder}' does not exist.")
        sys.exit(1)
    
    # Ensure output file has .xlsx extension
    output_file = args.output_file
    if not output_file.lower().endswith('.xlsx'):
        output_file += '.xlsx'
    
    # Check if output file already exists
    if os.path.exists(output_file):
        response = input(f"Output file '{output_file}' already exists. Overwrite? (y/N): ")
        if response.lower() not in ['y', 'yes']:
            print("Operation cancelled.")
            sys.exit(0)
    
    # Create merger and run
    merger = ExcelMerger()
    try:
        success = merger.merge_excel_files(args.input_folder, output_file, args.merge_mode)
        if success:
            print(f"\n✅ Excel files merged successfully!")
            sys.exit(0)
        else:
            print(f"\n❌ Failed to merge Excel files.")
            sys.exit(1)
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()