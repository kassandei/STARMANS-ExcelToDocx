# STARMANS Excel to Word Table Processor

This script processes an Excel file to extract specific data ranges and formats them into a Word document template. The file named `DOCXFILE.docx` acts as the template and should not be modified. The script reads data from the specified ranges in the Excel file, formats it, and populates the table in the Word document.

## Requirements

To run this script, you need the following libraries:

- `pandas`
- `python-docx`
- `openpyxl`

You can install these libraries using pip:

```sh
pip install pandas python-docx openpyxl
```

## Usage
```sh
python script_name.py <input_file.xlsx> <output_file.docx>
