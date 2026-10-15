# Excel Data Cleaner

A Python-based Excel data cleaning and validation tool designed to automate common data preparation tasks using **Pandas**.

The application allows users to select an Excel file, inspect its contents, apply multiple data-cleaning operations, and export the results into a separate cleaned workbook while preserving the original file.

## Features

- **Excel File Selection:** Uses Tkinter to provide a graphical file selection dialog.
- **Data Inspection:** Displays dataset dimensions, column data types, missing values, and duplicate counts.
- **Whitespace Cleaning:** Removes leading, trailing, and unnecessary whitespace from text values.
- **Duplicate Removal:** Removes identical records while retaining the first occurrence.
- **Date Standardization:** Identifies potential date columns and converts supported date formats into consistent datetime values.
- **Column Name Standardization:** Removes unnecessary whitespace from column headers and replaces spaces with underscores.
- **Capitalization Standardization:** Applies consistent capitalization to text while excluding certain values such as email addresses and abbreviations.
- **Error Handling:** Handles common file access errors using Python's `try/except` statements.
- **Cleaning Summary:** Reports the original row count, cleaned row count, number of duplicates removed, and remaining missing values.
- **Excel Export:** Saves cleaned data to a separate Excel file without modifying the original dataset.

## Technologies Used

- **Python** — Core programming language
- **Pandas** — Data manipulation, cleaning, and validation
- **OpenPyXL** — Excel file reading and writing
- **Tkinter** — Graphical file selection
- **Pathlib** — File path management
- **Git & GitHub** — Version control and project management

## Installation

**1. Clone the repository**

```bash
git clone https://github.com/jcharle6/excel-data-cleaner.git
```

**2. Navigate to the project directory**

```bash
cd excel-data-cleaner
```

**3. Install the required libraries**

```bash
pip install pandas openpyxl
```

Tkinter is included with most standard Python installations.

## Usage

Run the application:

```bash
python main.py
```

1. Select an Excel `.xlsx` file using the file selection window.
2. The program loads and inspects the dataset.
3. Cleaning functions are applied to the data.
4. A cleaning summary is displayed in the terminal.
5. A cleaned Excel file is saved in the same directory as the original.

### Example

**Original file:** `employee_data.xlsx`

**Cleaned file:** `employee_data_cleaned.xlsx`

The original file remains unchanged. Running the cleaner again on the same original file replaces the existing cleaned copy.

### Example Cleaning Summary

```text
========== CLEANING SUMMARY ==========

Original Rows: 1500
Cleaned Rows: 1475
Duplicates Removed: 25
Missing Values Remaining: 18

======================================
```

## Limitations

- The application currently processes the first worksheet of the selected Excel workbook.
- Date detection relies on pattern recognition and may require manual verification for ambiguous formats.
- Duplicate detection compares complete rows rather than identifying duplicates based on a specific business key.
- Capitalization standardization may not be appropriate for every type of text.
- Exporting through Pandas does not preserve all original Excel formatting, formulas, charts, or workbook features.
- Missing values are reported but are not automatically filled or deleted.


## Project Purpose

This project was developed to strengthen practical skills in Python programming, data cleaning, automation, and data validation.

It demonstrates how repetitive Excel data preparation tasks can be automated through reusable Python functions, helping improve consistency and reduce manual processing.

The project also serves as a foundation for developing more advanced ETL and data-processing workflows.