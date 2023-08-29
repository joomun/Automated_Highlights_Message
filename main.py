import pandas as pd
import tkinter as tk
from tkinter import filedialog, Listbox, Scrollbar, MULTIPLE, END, Toplevel, Button
from tkinter import ttk
from ttkthemes import ThemedStyle  # Import ThemedStyle from ttkthemes

import tkinter.messagebox as messagebox

class RowSelector:
    
    
    def __init__(self, root, df, selected_columns, file_path):
        self.df = df
        self.selected_columns = selected_columns
        self.file_path = file_path

        self.row_window = Toplevel(root)
        self.row_window.title("Row Selector")
        self.row_window.geometry("400x300")  # Set the window size

        self.create_row_listbox()

    def create_row_listbox(self):
        ttk.Label(self.row_window, text="Select Rows:", font=("Helvetica", 12)).pack(pady=10)

        self.row_listbox = Listbox(self.row_window, selectmode=MULTIPLE)
        self.row_listbox.pack()

        for idx, row_value in enumerate(self.df.iloc[:, 0]):
            self.row_listbox.insert(tk.END, f"{row_value} (Index: {idx})")

        confirm_button = ttk.Button(self.row_window, text="Confirm Selection", command=self.show_selected_data)
        confirm_button.pack(pady=10)

    def show_selected_data(self):
        selected_row_indices = self.row_listbox.curselection()
        selected_rows = [int(item.split(" ")[-1][:-1]) for item in [self.row_listbox.get(index) for index in selected_row_indices]]

        if selected_rows:
            selected_data = self.df.iloc[selected_rows][self.selected_columns]

            ttk.Label(root, text=f"Selected Data from: {self.file_path}", font=("Helvetica", 12)).pack()
            result_text = tk.Text(root, height=5, width=50)
            result_text.pack()

            result_text.insert(tk.END, str(selected_data))

def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_paths:
        excel_file_paths.set(file_paths)
        process_excel_files(file_paths)


def process_excel_files(file_paths):
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

        if "Night Audit Report.xls" in file_path:
            # Display preset columns in a message box
            preset_columns = ["Particulars", "Nett Day", "Nett Year","Test"]
            preset_columns_message = "\n".join(preset_columns)
            messagebox.showinfo("Preset Columns", f"These columns will be preset for Night Audit Report:\n\n{preset_columns_message}")
                    # Check if any of the preset columns are missing
            missing_preset_columns = [col for col in preset_columns if col not in available_columns]
            if missing_preset_columns:
                messagebox.showwarning("Missing Columns", f"The following preset columns are missing in the Excel file:\n\n{', '.join(missing_preset_columns)}")
            # Ask user if they want to proceed with preset columns
            proceed = messagebox.askyesno("Preset Columns", "Do you want to proceed with preset columns?")

            if proceed:
                selected_columns = preset_columns
                for i in range(column_listbox.size()):
                    if column_listbox.get(i) in selected_columns:
                        column_listbox.select_set(i)

                row_selector = RowSelector(root, df, selected_columns, file_path)
                row_selectors[file_path] = row_selector

        else:
            def show_selected_columns():
                selected_column_indices = column_listbox.curselection()
                selected_columns = [column_listbox.get(index) for index in selected_column_indices]

                if selected_columns:
                    row_selector = RowSelector(root, df, selected_columns, file_path)
                    row_selectors[file_path] = row_selector

            # Button to trigger column selection
            show_columns_button = ttk.Button(root, text="Select Columns", command=show_selected_columns)
            show_columns_button.pack()        
# Create a Tkinter window
root = tk.Tk()

root.tk.call("source","azure.tcl")
root.tk.call("set_theme","dark")
root.title("Excel Data Selector")

# Set the window size based on screen size
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = int(screen_width * 0.7)  # Set to 70% of screen width
window_height = int(screen_height * 0.7)  # Set to 70% of screen height
root.geometry(f"{window_width}x{window_height}")


# Create a StringVar to store the Excel file paths
excel_file_paths = tk.StringVar()

# Create a Canvas widget
canvas = tk.Canvas(root)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a Scrollbar for the canvas
scrollbar = tk.Scrollbar(root, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
canvas.configure(yscrollcommand=scrollbar.set)

# Create a Frame inside the canvas to hold the content
content_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

# Button to trigger file browsing with a reduced width
browse_button = ttk.Button(content_frame, text="Browse Files", command=browse_files, width=15)
browse_button.pack(pady=20)

# Dictionary to store RowSelector instances
row_selectors = {}

# Configure the canvas scrolling
def configure_canvas(event):
    canvas.configure(scrollregion=canvas.bbox("all"))

content_frame.bind("<Configure>", configure_canvas)

# Attach the canvas to the scrollbar
canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# Display the Tkinter window
root.mainloop()