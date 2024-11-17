# Excel SAP Script

**Notice**: The code in this repository is for review purposes only and is not intended for production use.

This is an Office Script designed to update specific cells in an Excel sheet based on SAP target values. The script processes data in columns and updates related columns accordingly. It handles date formatting and specific conditions for certain entries.

## Features:
- Updates cells in columns E, F, G, I, J, and H based on values in column C.
- Handles specific exceptions (e.g., "PR1", "PA8").
- Date formatting in columns F and J.
- Clears columns G and H for updated rows.

## How to Use:
1. Open the workbook in Excel.
2. Run the script using the built-in Office Script editor.
3. The script will automatically update the relevant columns based on the target values in column C.

## Requirements:
- Excel with support for Office Scripts.

## License:
This project is for review purposes only. It is not intended for production use without modifications. Use of the code is at your own risk.

## Changelog:
- **Version 2.1**: Improved performance and added new exception handling for "PR1", "PA8", and "PS1".
- **Version 2.0**: Initial implementation of the SAP target update script.

