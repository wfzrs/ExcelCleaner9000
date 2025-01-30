import random
import string
from hashlib import blake2b
import pandas as pd

def generate_name_hash(str_to_hash):
    '''Creates a unique hash out of a given string'''
    h = blake2b(digest_size=10)
    try:
        h.update(str_to_hash.encode('UTF-8'))
    except AttributeError:
        return "ERROR"
    else:
        return h.hexdigest()


def add_random_code(input_file, output_file):
    '''
    Uses dataframes to read the LastName, FirstName, and Teacher columns to create
    a unique code for every student
    '''
    # Read the Excel file
    df = pd.read_excel(input_file)

    for index, row in df.iterrows():
        try:
            combined = row['Last Name'] + row['First Name'] + row['Teacher']
        except:
            row['Codes'] = "ERROR"
            print(f"Error: Invalid data")
        else:
            row['Codes'] = generate_name_hash(combined)
            print(f"Hashed code: {row['Codes']}")

    # Save the updated dataframe to a new Excel file
    df.to_excel(output_file, index=False)
    print(f"Updated file saved as: {output_file}")


# Path to the Excel file
file_path = "./your_excel_file.xlsx"

# Read all sheets into a dictionary of DataFrames
all_sheets = pd.read_excel(file_path, sheet_name=None, header=None)

# Initialize an empty list to store cleaned DataFrames
cleaned_dataframes = []

# Iterate through each sheet
for sheet_name, df in all_sheets.items():
    # Skip rows matching the header row (assumes header is always the same)
    try:
        df_cleaned = df[df.iloc[:, 0] != "Last Name"]
    except IndexError:
        print(f'Error: {sheet_name}\n')
    else:
        cleaned_dataframes.append(df_cleaned)

# Combine cleaned DataFrames into one
combined_data = pd.concat(cleaned_dataframes, ignore_index=True)

# Optionally, rename the columns
#combined_data.columns = ["Last Name", "First Name", "Teacher", "School Name", "Codes"]

# Save the combined data to a new Excel file
output_file = "Master_Excel_cleaned.xlsx"
combined_data.to_excel(output_file, index=False)

print(f"Cleaned combined data saved to {output_file}")
