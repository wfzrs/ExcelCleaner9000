import os
import pandas as pd

def split_excel_file(input_file, column_name):
    # Read the Excel file
    df = pd.read_excel(input_file)

    # Get unique values in the specified column
    unique_values = df[column_name].unique()

    # Split the dataframe and save to separate Excel files
    for value in unique_values:
        #debug
        print(f"value: {value}")
        subset_df = df[df[column_name] == value]
        output_file = f"./out-files/{value}.xlsx"
        subset_df.to_excel(output_file, index=False)
        #print(f"Created file: {output_file}")

# First check if dir 'out-files' exists and create it if it doesn't.
directory = "out-files"

# Check if directory exists
if not os.path.exists(directory):
    # Create the directory
    os.makedirs(directory)
    print(f"Directory '{directory}' created.")
else:
    print(f"Directory '{directory}' already exists.")

# Example usage
input_file = './2025-fit-for-fun-student-list.xlsx'
column_name = 'School Name'
split_excel_file(input_file, column_name)