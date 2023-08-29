import pandas as pd
import tkinter as tk
from tkinter import filedialog

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        excel_file_path.set(file_path)
        process_excel_data(file_path)

def process_excel_data(file_path):
    # Read data from Excel file
    df = pd.read_excel(file_path)

    # Display the first few rows of the DataFrame
    print("First few rows:")
    print(df.head())

    # Access a specific column
    column_data = df['Column_Name']  # Replace 'Column_Name' with the actual column name
    print("Specific column:")
    print(column_data)

    # Access a specific cell
    row_index = 0  # Replace with the desired row index
    cell_value = df.at[row_index, 'Column_Name']  # Replace 'Column_Name' with the actual column name
    print(f"Value at row {row_index}: {cell_value}")

    # Filter rows based on a condition
    filtered_df = df[df['Numeric_Column'] > 50]  # Replace 'Numeric_Column' with the actual column name
    print("Filtered rows:")
    print(filtered_df)

    # Iterate through rows
    print("Iterating through rows:")
    for index, row in df.iterrows():
        print(index, row['Column_Name'])  # Replace 'Column_Name' with the actual column name

# Create a Tkinter window
root = tk.Tk()

# Create a StringVar to store the Excel file path
excel_file_path = tk.StringVar()

# Button to trigger file browsing
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack()

# Display the Tkinter window
root.mainloop()
