import pandas as pd

# Load the Excel file
input_file = "./in-files/2025-Roaring-Reader.xlsx"
df = pd.read_excel(input_file)

print(df['Registered School Name'][1])
exit()

# Transpose the DataFrame
df_transposed = df.transpose()

print(df_transposed)

# Save the transposed DataFrame to a new Excel file
output_file = "./out-files/transposed_output.xlsx"
df_transposed.to_excel(output_file, index=False)

print(f"The transposed data has been saved to {output_file}.")