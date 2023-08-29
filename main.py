import pandas as pd
import tkinter as tk
from tkinter import filedialog, Listbox, Scrollbar, MULTIPLE, END

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        excel_file_path.set(file_path)
        process_excel_data(file_path)

def process_excel_data(file_path):
    # Read data from Excel file
    df = pd.read_excel(file_path)

    # Display the list of available columns
    available_columns = df.columns.tolist()

    # Create a listbox for column selection
    listbox = Listbox(root, selectmode=MULTIPLE)
    listbox.pack()

    for col in available_columns:
        listbox.insert(tk.END, col)

    def show_selected_columns():
        selected_indices = listbox.curselection()
        selected_columns = [listbox.get(index) for index in selected_indices]

        if selected_columns:
            selected_data = df[selected_columns]

            # Clear previous text and display new data
            result_text.delete(1.0, tk.END)
            result_text.insert(tk.END, str(selected_data))

            # Resize the result_text widget based on content
            result_text.config(height=min(20, len(selected_data) + 2))

    # Button to trigger displaying selected columns
    show_columns_button = tk.Button(root, text="Show Selected Columns", command=show_selected_columns)
    show_columns_button.pack()

    # Create a Text widget to display the result
    result_text = tk.Text(root, height=5, width=50)
    result_text.pack()

# Create a Tkinter window
root = tk.Tk()

# Create a StringVar to store the Excel file path
excel_file_path = tk.StringVar()

# Button to trigger file browsing
browse_button = tk.Button(root, text="Browse", command=browse_file)
browse_button.pack()

# Display the Tkinter window
root.mainloop()
