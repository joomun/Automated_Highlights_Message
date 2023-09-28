import pandas as pd
import tkinter as tk
from tkinter import filedialog, Listbox, Scrollbar, MULTIPLE, END, Toplevel, Button
from tkinter import ttk
from ttkthemes import ThemedStyle  # Import ThemedStyle from ttkthemes

import tkinter.messagebox as messagebox

class RowSelector:
    def __init__(self, root, df, selected_columns, file_path, preset_rows_values=None):
        self.df = df
        self.selected_columns = selected_columns
        self.file_path = file_path

        self.row_window = Toplevel(root)
        self.row_window.title("Row Selector")
        self.row_window.geometry("400x300")  # Set the window size

        if preset_rows_values:
            self.show_selected_data_by_values(preset_rows_values)
        else:
            self.create_row_listbox()
            
    def show_selected_data(self):
        selected_row_indices = self.row_listbox.curselection()
        selected_rows = [self.df.index[i] for i in selected_row_indices]
        selected_data = self.df.loc[selected_rows, self.selected_columns]
        print(selected_data)
        ttk.Label(self.row_window, text=f"Selected Data from: {self.file_path}", font=("Helvetica", 12)).pack()
        result_text = tk.Text(self.row_window, height=5, width=50)
        result_text.pack()
        result_text.insert(tk.END, str(selected_data))

    def create_row_listbox(self):
        ttk.Label(self.row_window, text="Select Rows:", font=("Helvetica", 12)).pack(pady=10)

        self.row_listbox = Listbox(self.row_window, selectmode=MULTIPLE)
        self.row_listbox.pack()

        for idx, row_value in enumerate(self.df.iloc[:, 0]):
            self.row_listbox.insert(tk.END, f"{row_value} (Index: {idx})")

        confirm_button = ttk.Button(self.row_window, text="Confirm Selection", command=self.show_selected_data)
        confirm_button.pack(pady=10)

    def show_selected_data_by_values(self, row_values):
            # This function selects data based on row values and displays them
            selected_rows = self.df[self.df.iloc[:, 0].isin(row_values)].index.tolist()
            selected_data = self.df.iloc[selected_rows][self.selected_columns]
            print(selected_data)
            ttk.Label(root, text=f"Selected Data from: {self.file_path}", font=("Helvetica", 12)).pack()
            result_text = tk.Text(root, height=5, width=50)
            result_text.pack()
            result_text.insert(tk.END, str(selected_data))
            


preset_columns_NAR = ["Particulars", "Nett Day", "Nett Year"]  # Add other columns you want
preset_rows_NAR_values = ["Room Revenue","Food - All Day FullBoard", "Room Revenue -  No Show","Food & Beverages","Food - All Day HalfBoard","Food - All Day FullBoard","Food - All Day (Menus)","Food Breakfast Package","Food Breakfast (Menus)","Food - Meeting Room","Beverage - Beer","Beverage - Hot Drinks","Beverage - House Wine","Beverage - Soft Drinks","Beverage - Spirit","Beverage - Water","Space Rent - Accelerator","Space Rent - Boardroom","Space Rent - Educator (M)","Space Rent - Incubator","Commission - Car Rental","Commission - Forex Exchge","Commission - Paid Out","Commission -Taxi Services","Commission - Transfer","ICT Service - Room","Laundry - Contracted","Laundry - Inhouse","Misc - Currency Gain/Loss","Misc - Others","Misc - Photocopy","Phone Calls Local","NET REVENUE","TOTAL REVENUE"]
# Function to open the preset configuration window
# Function to open the preset configuration window
def open_preset_config():
    preset_config_window = Toplevel(root)
    preset_config_window.title("Preset Configuration")
    preset_config_window.geometry("400x400")

    # Create labels and entry fields for rows and columns
    ttk.Label(preset_config_window, text="Preset Rows (comma-separated):").pack(pady=10)
    rows_entry = ttk.Entry(preset_config_window)
    rows_entry.pack(pady=10)
    rows_entry.insert(0, ",".join(preset_rows_NAR_values))

    ttk.Label(preset_config_window, text="Preset Columns (comma-separated):").pack(pady=10)
    columns_entry = ttk.Entry(preset_config_window)
    columns_entry.pack(pady=10)
    columns_entry.insert(0, ",".join(preset_columns_NAR))

    # Function to update preset values
    def update_preset_values():
        global preset_rows_NAR_values, preset_columns_NAR
        preset_rows = rows_entry.get().split(",")
        preset_columns = columns_entry.get().split(",")
        preset_rows_NAR_values = [item.strip() for item in preset_rows]
        preset_columns_NAR = [item.strip() for item in preset_columns]
        messagebox.showinfo("Preset Configuration", "Preset values updated successfully.")
        preset_config_window.destroy()

    # Button to update preset values
    update_button = ttk.Button(preset_config_window, text="Update Preset Values", command=update_preset_values)
    update_button.pack(pady=20)

    # Display current preset options
    ttk.Label(preset_config_window, text="Current Preset Rows:").pack(pady=10)
    current_rows_label = ttk.Label(preset_config_window, text=", ".join(preset_rows_NAR_values))
    current_rows_label.pack(pady=10)

    ttk.Label(preset_config_window, text="Current Preset Columns:").pack(pady=10)
    current_columns_label = ttk.Label(preset_config_window, text=", ".join(preset_columns_NAR))
    current_columns_label.pack(pady=10)

    # Function to add new preset row
    def add_new_row():
        new_row = simpledialog.askstring("Add New Row", "Enter a new row value:")
        if new_row:
            preset_rows_NAR_values.append(new_row.strip())
            current_rows_label.config(text=", ".join(preset_rows_NAR_values))

    # Button to add new preset row
    add_row_button = ttk.Button(preset_config_window, text="Add New Row", command=add_new_row)
    add_row_button.pack(pady=10)

    # Function to add new preset column
    def add_new_column():
        new_column = simpledialog.askstring("Add New Column", "Enter a new column value:")
        if new_column:
            preset_columns_NAR.append(new_column.strip())
            current_columns_label.config(text=", ".join(preset_columns_NAR))

    # Button to add new preset column
    add_column_button = ttk.Button(preset_config_window, text="Add New Column", command=add_new_column)
    add_column_button.pack(pady=10)
def browse_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_paths:
        excel_file_paths.set(file_paths)
        process_excel_files(file_paths)

def configure_canvas(event):
    canvas.configure(scrollregion=canvas.bbox("all"))
    content_frame.update_idletasks()  # Update the content frame's size
    
    # Add a vertical scrollbar to the content frame
    canvas_width = content_frame.winfo_reqwidth()
    if canvas_width > window_width:
        canvas.config(scrollregion=canvas.bbox("all"))
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    else:
        scrollbar.pack_forget()

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

        def show_selected_columns():
            selected_column_indices = column_listbox.curselection()
            selected_columns = [column_listbox.get(index) for index in selected_column_indices]

            if selected_columns:
                row_selector = RowSelector(root, df, selected_columns, file_path)
                row_selectors[file_path] = row_selector

        # Button to trigger column selection
        show_columns_button = ttk.Button(root, text="Select Columns", command=show_selected_columns)
        show_columns_button.pack()

        if "Night Audit Report.xls" in file_path:
            # Display preset columns in a message box
            preset_columns = preset_columns_NAR
            preset_rows_values = preset_rows_NAR_values  # Use preset rows
            preset_columns_message = ",".join(preset_columns)
            preset_columns_message = ",".join(preset_rows_values)
            messagebox.showinfo("Preset Columns", f"These columns will be preset for Night Audit Report:\n\n{preset_columns_message}")

            # Check if any of the preset columns are missing
            missing_preset_columns = [col for col in preset_columns if col not in available_columns]

            if missing_preset_columns:
                messagebox.showwarning("Missing Columns", f"The following preset columns are missing in the Excel file:\n\n{', '.join(missing_preset_columns)}")

                # Remove missing preset columns from available_columns list
                preset_columns = [col for col in preset_columns if col not in missing_preset_columns]

            # Ask the user if they want to proceed with preset columns
            proceed = messagebox.askyesno("Preset Columns", "Do you want to proceed with preset columns?")

            if proceed:
                selected_columns = preset_columns
                for i in range(column_listbox.size()):
                    if column_listbox.get(i) in selected_columns:
                        column_listbox.select_set(i)

                # When creating the RowSelector instance:
                row_selector = RowSelector(root, df, selected_columns, file_path, preset_rows_values)
           
            
# Create a Tkinter window
root = tk.Tk()

root.tk.call("source","azure.tcl")
root.tk.call("set_theme","dark")
root.title("Excel Data Selector")

# Set the window size based on screen sizegit 
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = int(screen_width * 1)  # Set to 70% of screen width
window_height = int(screen_height * 1)  # Set to 70% of screen height
root.geometry(f"{window_width}x{window_height}")


# Create a StringVar to store the Excel file paths
excel_file_paths = tk.StringVar()

# Create a Canvas widget
canvas = tk.Canvas(root)
canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

# Create a Scrollbar for the canvas
scrollbar = Scrollbar(root, command=canvas.yview)
scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

# Create a Frame inside the canvas to hold the content
content_frame = tk.Frame(canvas)
canvas.create_window((0, 0), window=content_frame, anchor="nw")

# Button to trigger file browsing with a reduced width
browse_button = ttk.Button(content_frame, text="Browse Files", command=browse_files, width=15)
browse_button.pack(pady=20)

# Button to open the preset configuration window
preset_config_button = ttk.Button(content_frame, text="Configure Preset", command=open_preset_config, width=15)
preset_config_button.pack(pady=10)

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