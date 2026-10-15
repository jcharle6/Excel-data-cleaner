from pathlib import Path
from tkinter import Tk, filedialog
import pandas as pd
import sys



# Create the Tkinter application
root = Tk()

# Hide the empty main application window
root.withdraw()

# Open the Windows file selection dialog
file_path = filedialog.askopenfilename(
    title="Select an Excel File",
    filetypes=[("Excel Files", "*.xlsx")]
)

# Close the Tkinter application
root.destroy()

# # Display the selected file path
# if file_path:
#     df = pd.read_excel(file_path)
#     print(df.head())
#     #print("Selected file:", file_path)
# else:
#     print("No file was selected.")

# Process the selected file
if file_path:
    try:
        df = pd.read_excel(file_path)

        print("\nFIRST FIVE ROWS:")
        print(df.head())

        print("\nDATASET DIMENSIONS:")
        print(f"Rows: {df.shape[0]}")
        print(f"Columns: {df.shape[1]}")

        print("\nCOLUMN DATA TYPES:")
        print(df.dtypes)

        print("\nMISSING VALUES BY COLUMN:")
        print(df.isna().sum())

        print("\nDUPLICATE ROWS:")
        print(df.duplicated().sum())
    except PermissionError:
        print("ERROR: Please close the Excel file and try again.")
        sys.exit()
else:
    print("No file selected.")


#checks every columns for leadng, trailing or extra whitespace
#independent review is still advised
def clean_whitespace(df):
    for column in df.select_dtypes(include=["object", "string"]).columns:
        df[column] = df[column].map(
            lambda value: " ".join(value.split())
            if isinstance(value, str)
            else value
        )

    return df

#delete duplicate records, keeps the original record
def clean_duplicates(df):

    # Remove duplicate rows, keeping the first occurrence
    df = df.drop_duplicates(keep="first").copy()

    return df

def clean_dates(df):
    for column in df.columns:

        # Skip columns already recognized as dates
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            continue

        # Only inspect text-based columns
        if not (
            pd.api.types.is_object_dtype(df[column])
            or pd.api.types.is_string_dtype(df[column])
        ):
            continue

        # Get non-empty values
        values = df[column].dropna()

        if values.empty:
            continue

        # Check whether most values resemble dates
        date_pattern = (
            r"(?i)(?:\d{1,4}[/\-]\d{1,2}[/\-]\d{1,4}"
            r"|[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{4})"
        )

        date_candidates = values.map(
            lambda x: isinstance(x, str)
            and bool(pd.Series([x.strip()]).str.fullmatch(date_pattern).iloc[0])
        )

        # Require at least 90% date-like values
        if date_candidates.mean() < 0.90:
            continue

        # Attempt conversion
        converted = pd.to_datetime(
            df[column],
            format="mixed",
            errors="coerce"
        )

        # Only convert if all non-empty values parse
        if converted[values.index].notna().all():
            df[column] = converted
            print(f"Date column standardized: {column}")

        else:
            print(f"Date column needs review: {column}")

    return df

#column names such as employee name becomes employee_name
def clean_column_names(df):

    # Store the original column names
    original_columns = df.columns.tolist()

    # Clean each column name
    df.columns = [
        "_".join(str(column).split()).title()
        for column in df.columns
    ]

    return df

#capatilizes text when appropriate
def standardize_capitalization(df):

    for column in df.select_dtypes(
        include=["object", "string"]
    ).columns:

        def clean_value(value):

            if not isinstance(value, str):
                return value

            # Skip email addresses
            if "@" in value:
                return value

            # Skip values containing numbers
            if any(char.isdigit() for char in value):
                return value

            # Skip fully uppercase abbreviations
            if value.isupper() and len(value) <= 3:
                return value

            # Standardize capitalization
            return value.title()

        df[column] = df[column].map(clean_value)

    return df

# Store original dataset statistics
original_rows = len(df)
original_duplicates = df.duplicated().sum()

#calling our cleaning functions
df = clean_whitespace(df)
df = clean_dates(df)
df = clean_column_names(df)
df = standardize_capitalization(df)
df = clean_duplicates(df)

# Export cleaned Excel file
original_path = Path(file_path)

output_path = original_path.with_name(
    original_path.stem + "_cleaned.xlsx"
)

try:
    df.to_excel(output_path, index=False)
    print(f"Cleaned file saved: {output_path}")

except PermissionError:
    print("ERROR: Please close the cleaned Excel file.")
    raise SystemExit(1)

# Cleaning summary
cleaned_rows = len(df)
duplicates_removed = original_rows - cleaned_rows
missing_values = df.isna().sum().sum()

print("\n========== CLEANING SUMMARY ==========")
print(f"Original Rows: {original_rows}")
print(f"Cleaned Rows: {cleaned_rows}")
print(f"Duplicates Removed: {duplicates_removed}")
print(f"Missing Values Remaining: {missing_values}")
print("======================================")