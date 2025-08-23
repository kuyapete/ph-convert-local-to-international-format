# Phone Converter

A Python script designed to convert local Philippine phone numbers to international format by adding the country code "63" prefix. This tool processes large Excel files and automatically cleans up empty rows while standardizing phone number formats.

## Features

- **Automatic Phone Number Conversion**: Converts local Philippine numbers to international format
- **Excel File Processing**: Handles large datasets (up to 224,680 rows)
- **Smart Format Detection**: Automatically detects and converts different local formats
- **Philippine Number Validation**: Filters out non-Philippine and invalid phone numbers
- **Data Cleanup**: Removes empty rows and non-digit characters
- **Batch Processing**: Processes entire Excel files in one operation
- **Safe Processing**: Uses new worksheet approach to avoid file corruption

## Phone Number Conversion Rules

| Input Format | Output Format | Example |
|--------------|---------------|---------|
| Starts with '9' | Adds '63' prefix | `9123456789` → `639123456789` |
| Starts with '0' | Adds '63' prefix, removes '0' | `09123456789` → `639123456789` |
| Already starts with '63' | No change | `639123456789` → `639123456789` |

## Requirements

- Python 3.7 or higher
- pandas
- openpyxl

## Installation

1. **Clone or download** this repository to your local machine

2. **Install required dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### Basic Usage

1. **Update the input file path** in `localToInternational.py`:
   ```python
   input_file = "your_file.xlsx"  # Change this to your file path
   ```

2. **Run the script**:
   ```bash
   python localToInternational.py
   ```

### Advanced Usage

You can also specify a custom output file:

```python
# Save to a new file instead of overwriting
convert_phone_numbers(input_file, "converted_output.xlsx")
```

## Input File Requirements

- **Format**: Excel file (.xlsx)
- **Phone Numbers**: Must be in column A
- **Starting Row**: Processing begins from row 2
- **Maximum Rows**: Up to 224,680 rows supported

## Output

The script will:
- Convert all phone numbers to international format
- Remove empty rows automatically
- Save the processed file
- Display processing statistics

## Example Output

```
Processing complete!
Converted phone numbers to country code format (63 prefix)
Deleted 15 empty rows
File saved as: your_file.xlsx
```

## Use Cases

- **Telecom Companies**: Converting customer databases to international format
- **Businesses**: Standardizing contact lists for international operations
- **Data Analysts**: Cleaning and formatting large phone number datasets
- **Administrators**: Managing contact databases with consistent formatting

## File Structure

```
Phone Converter/
├── localToInternational.py    # Main conversion script
├── requirements.txt           # Python dependencies
├── README.md                 # This file
└── PBET.xlsx                # Example input file (if provided)
```

## Error Handling

The script includes error handling for:
- File not found errors
- General processing errors
- Invalid file formats

## Notes

- **Backup**: Always backup your original Excel file before processing
- **Large Files**: The script is optimized for large datasets but may take time for very large files
- **Format Preservation**: Only phone numbers are modified; other data remains unchanged

## Support

If you encounter any issues or have questions about the phone number conversion logic, please check:
1. File path is correct
2. Excel file is not corrupted
3. Phone numbers are in the expected format

## License

This project is open source and available for personal and commercial use.
