# Excel Cleaner 9000

Small collection of Python scripts to help modify Excel data for whatever reason.

## excel_tab_combine.py

This Python script allows you to combine data from multiple sheets in an Excel file into one consolidated sheet. Each sheet must have the same structure, and the script ensures that duplicate header rows are removed in the final output.

---

### Requirements

1. Python 3.x
2. The following Python libraries:
   - `pandas`
   - `openpyxl`
   - `virtualenv` (optional)

You can install the required libraries by running:
```bash
pip install -r requirements.txt
```

---

### Script Overview

The script:
1. Reads all sheets in an Excel file.
2. Removes any duplicate header rows (e.g., rows starting with "Last Name").
3. Combines the cleaned data from all sheets into one.
4. Saves the consolidated data to a new Excel file.

---

### How to Use the Script

1. Place your Excel file (e.g., `your_excel_file.xlsx`) in the same directory as the script.

2. Update the `file_path` variable in the script with the name of your Excel file:
   ```python
   file_path = "your_excel_file.xlsx"
   ```

3. Run the script. It will generate a new Excel file named `combined_data_cleaned.xlsx` in the same directory.

---

### Example Output

#### Input
An Excel file with multiple sheets, each containing data like this:

| Last Name | First Name | Teacher      | School Name | Codes |
|-----------|------------|--------------|-------------|-------|
| jones     | rory       | mrs. teacher | school2     | code2 |

#### Output
A single Excel sheet combining all rows from all sheets, with duplicate headers removed:

| Last Name | First Name | Teacher      | School Name | Codes |
|-----------|------------|--------------|-------------|-------|
| jones     | rory       | mrs. teacher | school2     | code2 |
| johnson   | dustin     | ms. Red      | school3     | code3 |

---

## excel_file_split.py

Splits a single Excel file/sheet into multiple files based on a given column's values. The column to look at can be changed by modifying the `column_name` variable.