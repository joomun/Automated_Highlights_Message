import pandas as pd

# Read data from Excel file
excel_file_path = 'sample.xlsx'  # Replace with your actual file path
df = pd.read_excel(excel_file_path)

# Display the first few rows of the DataFrame
print("First few rows:")
print(df.head())

# Access a specific column
column_data = df['Column_Name']  # Replace 'Column_Name' with the actual column name

# Access a specific cell
row_index = 0  # Replace with the desired row index
cell_value = df.at[row_index, 'Column_Name']  # Replace 'Column_Name' with the actual column name
print(f"Value at row {row_index}: {cell_value}")

# Filter rows based on a condition
filtered_df = df[df['Numeric_Column'] > 50]  # Replace 'Numeric_Column' with the actual column name

# Iterate through rows
print("Iterating through rows:")
for index, row in df.iterrows():
    print(index, row['Column_Name'])  # Replace 'Column_Name' with the actual column name
