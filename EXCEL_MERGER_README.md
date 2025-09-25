# Excel File Merger

A Python script that merges all Excel files (.xlsx and .xls) in a specified folder into a single Excel workbook.

## Features

- ✅ Supports both .xlsx and .xls file formats
- ✅ Handles multiple sheets within each Excel file
- ✅ Two merge modes: separate sheets or single combined sheet
- ✅ Automatic sheet naming to distinguish data from different source files
- ✅ Robust error handling for corrupted or unreadable files
- ✅ Command-line interface with helpful options
- ✅ Progress tracking and detailed summary reports

## Installation

1. Make sure you have Python 3.6+ installed
2. Install required dependencies:
```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install pandas openpyxl xlrd
```

## Usage

### Basic Usage
```bash
python excel_merger.py -i /path/to/excel/files -o merged_output.xlsx
```

### Command Line Options

- `-i, --input-folder`: Path to folder containing Excel files to merge (required)
- `-o, --output-file`: Name of the output Excel file (required)
- `-m, --merge-mode`: How to merge the data (optional, default: separate)
  - `separate`: Keep each sheet in separate tabs
  - `single`: Combine all data into one sheet with source identification

### Examples

#### Keep sheets separate (default behavior)
```bash
python excel_merger.py -i ./excel_files -o combined_workbook.xlsx
```

#### Combine all data into single sheet
```bash
python excel_merger.py -i ./data -o merged_data.xlsx -m single
```

#### With full paths
```bash
python excel_merger.py --input-folder "/home/user/Documents/excel_files" --output-file "final_report.xlsx" --merge-mode separate
```

## Merge Modes

### Separate Sheets Mode (default)
- Each sheet from each source file becomes a separate sheet in the output
- Sheet names are prefixed with the source filename: `filename_sheetname`
- Preserves original data structure and formatting
- Best for maintaining data organization

### Single Sheet Mode
- All data is combined into one sheet named "Merged_Data"
- A "Source_File_Sheet" column is added to identify the origin of each row
- Useful for data analysis across all files
- May lose some formatting but creates a unified dataset

## Error Handling

The script includes comprehensive error handling:

- **Missing folder**: Clear error message if input folder doesn't exist
- **No Excel files**: Graceful handling when no .xlsx/.xls files are found
- **Corrupted files**: Skip corrupted files and continue with others
- **Empty sheets**: Skip empty sheets with warning messages
- **Permission errors**: Handle file access issues
- **Existing output**: Prompt before overwriting existing files

## Output Summary

After completion, the script provides a detailed summary:
- Number of files processed successfully
- List of failed files (if any)
- Total number of sheets merged
- Output file location
- Row and column counts (for single sheet mode)

## File Naming

### Input Files
- Supports: `.xlsx`, `.xls` (case-insensitive)
- Hidden files and other formats are ignored

### Sheet Names
- **Separate mode**: `{filename}_{original_sheet_name}`
- **Single mode**: Original data with added source column
- Long sheet names are truncated to Excel's 31-character limit

## Troubleshooting

### Common Issues

1. **ImportError**: Install missing dependencies
   ```bash
   pip install pandas openpyxl xlrd
   ```

2. **Permission denied**: Ensure you have read access to input files and write access to output location

3. **Memory issues**: For very large files, consider processing in smaller batches

4. **Encoding issues**: The script handles most encoding issues automatically, but ensure your Excel files are not password-protected

### Getting Help

Run the script with `-h` or `--help` for detailed usage information:
```bash
python excel_merger.py --help
```

## Requirements

- Python 3.6+
- pandas >= 1.5.0
- openpyxl >= 3.0.0 (for .xlsx files)
- xlrd >= 2.0.0 (for .xls files)

## License

This script is provided as-is for educational and practical use.