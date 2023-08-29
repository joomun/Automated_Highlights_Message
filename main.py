import pandas as pd
import tkinter as tk
from tkinter import filedialog, Listbox, Scrollbar, MULTIPLE, END, Toplevel, Button

def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_paths:
        excel_file_paths.set(file_paths)
        process_excel_files(file_paths)

def process_excel_files(file_paths):
    selected_data_combined = pd.DataFrame()  # Store selected data from all files

    for file_path in file_paths:
        # Read data from Excel file
        df = pd.read_excel(file_path)

        # Display the list of available columns
        available_columns = df.columns.tolist()

        # Create a listbox for column selection
        column_listbox = Listbox(root, selectmode=MULTIPLE)
        column_listbox.pack()

        for col in available_columns:
            column_listbox.insert(tk.END, col)

        def show_selected_columns():
            selected_column_indices = column_listbox.curselection()
            selected_columns = [column_listbox.get(index) for index in selected_column_indices]

            if selected_columns:
                row_selector = RowSelector(root, df, selected_columns, selected_data_combined)

        # Button to trigger column selection
        show_columns_button = tk.Button(root, text="Select Columns", command=show_selected_columns)
        show_columns_button.pack()

# Create a Tkinter window
root = tk.Tk()
root.title("Excel Data Selector")

# Create a StringVar to store the Excel file paths
excel_file_paths = tk.StringVar()

# Button to trigger file browsing
browse_button = tk.Button(root, text="Browse Files", command=browse_files)
browse_button.pack()

class RowSelector:
    def __init__(self, root, df, selected_columns, selected_data_combined):
        self.df = df
        self.selected_columns = selected_columns
        self.selected_data_combined = selected_data_combined

        self.row_window = Toplevel(root)
        self.row_window.title("Row Selector")

        self.create_row_listbox()

    def create_row_listbox(self):
        self.row_listbox = Listbox(self.row_window, selectmode=MULTIPLE)
        self.row_listbox.pack()

        for idx, row_value in enumerate(self.df.iloc[:, 0]):
            self.row_listbox.insert(tk.END, f"{row_value} (Index: {idx})")

        confirm_button = Button(self.row_window, text="Confirm Selection", command=self.show_selected_data)
        confirm_button.pack()

    def show_selected_data(self):
        selected_row_indices = self.row_listbox.curselection()
        selected_rows = [int(item.split(" ")[-1][:-1]) for item in [self.row_listbox.get(index) for index in selected_row_indices]]

        if selected_rows:
            selected_data = self.df.iloc[selected_rows][self.selected_columns]
            self.selected_data_combined = pd.concat([self.selected_data_combined, selected_data], ignore_index=True)

            # Clear previous text and display new data
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, str(self.selected_data_combined))

# Create a Text widget to display the result
result_text = tk.Text(root, height=5, width=50)
result_text.pack()

# Display the Tkinter window
root.mainloop()
