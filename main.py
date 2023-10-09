import pandas as pd
import tkinter as tk
from tkinter import (filedialog, Listbox, Scrollbar, MULTIPLE, END, Toplevel,
                     ttk, messagebox, simpledialog)
from PIL import Image, ImageTk
from ttkthemes import ThemedStyle
from twilio.rest import Client
import os

class RowSelector:
    def __init__(self, root, df, selected_columns, file_path, preset_rows_values=None):
        self.df = df
        self.selected_columns = selected_columns
        self.file_path = file_path
        self.row_window = Toplevel(root)
        self.row_window.title("Row Selector")
        self.row_window.geometry("400x300")

        if preset_rows_values:
            self.show_selected_data_by_values(preset_rows_values)
        else:
            self.create_row_listbox()

    def show_selected_data(self):
        selected_rows = [self.df.index[i] for i in self.row_listbox.curselection()]
        selected_data = self.df.loc[selected_rows, self.selected_columns]
        self.display_data(selected_data)

    def create_row_listbox(self):
        ttk.Label(self.row_window, text="Select Rows:", font=("Helvetica", 12)).pack(pady=10)
        self.row_listbox = Listbox(self.row_window, selectmode=MULTIPLE)
        self.row_listbox.pack()
        for idx, row_value in enumerate(self.df.iloc[:, 0]):
            self.row_listbox.insert(tk.END, f"{row_value} (Index: {idx})")
        ttk.Button(self.row_window, text="Confirm Selection", command=self.show_selected_data).pack(pady=10)

    def show_selected_data_by_values(self, row_values):
        selected_rows = self.df[self.df.iloc[:, 0].isin(row_values)].index.tolist()
        selected_data = self.df.iloc[selected_rows][self.selected_columns]
        self.display_data(selected_data)

    def display_data(self, data):
        ttk.Label(self.row_window, text=f"Selected Data from: {self.file_path}", font=("Helvetica", 12)).pack()
        result_text = tk.Text(self.row_window, height=5, width=50)
        result_text.pack()
        result_text.insert(tk.END, str(data))       


class App:
    def __init__(self, root):
        self.root = root
        self.configure_ui()
        self.row_selectors = {}
        self.preset_columns_NAR = ["Particulars", "Nett Day", "Nett Year"]
        self.preset_rows_NAR_values = [
            "Room Revenue",
            "Room Revenue - Allowance",
            "Room Revenue - Cancel Fee",
            "Room Revenue - No Show",
            "Food & Beverages",
            "Food - All Day HalfBoard",
            "Food - All Day FullBoard",
            "Food - All Day (Menus)",
            "Food Breakfast Package",
            "Food Breakfast (Menus)",
            "Food - Meeting Room",
            "Beverage - Beer",
            "Beverage - Hot Drinks",
            "Beverage - House Wine",
            "Beverage - Soft Drinks",
            "Beverage - Spirit",
            "Beverage - Water",

            "ICT Service - Accelerator",
            "ICT Service - Boardroom",
            "ICT Service - Educator(M)",
            "ICT Service - Educator 1",
            "ICT Service - Incubator",
            "Space Rent - Accelerator",
            "Space Rent - Boardroom",
            "Space Rent - Educator (M)",
            "Space Rent - Educator 1",
            "Space Rent - Incubator",

            "Commission - Forex Exchge",
            "Commission - Paid Out",
            "Commission -Taxi Services",
            "Commission - Transfer",
            "ICT Service - Room",
            "Laundry - Contracted",
            "Laundry - Inhouse",
            "Misc - Currency Gain/Loss",
            "Misc - Others",
            "Misc - Photocopy",
            "Phone Calls Local",
            "Internet",

            
            "TOTAL REVENUE",
            "TOTAL",
            "NET REVENUE",
            
            
        ]

    def configure_ui(self):
        self.root.tk.call("source", "azure.tcl")
        self.root.tk.call("set_theme", "light")
        self.root.title("Excel Data Selector")
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        window_width = int(screen_width * 1)
        window_height = int(screen_height * 1)
        
        # Calculate the position to center the window
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")

        self.canvas = tk.Canvas(self.root)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar = Scrollbar(self.root, command=self.canvas.yview)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.content_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.content_frame, anchor="nw")

        ttk.Button(self.content_frame, text="Browse Files", command=self.browse_files, width=15).pack(pady=20)
        ttk.Button(self.content_frame, text="Configure Preset", command=self.open_preset_config, width=15).pack(pady=10)

        self.content_frame.bind("<Configure>", self.configure_canvas)
        self.canvas.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

    def configure_canvas(self, event):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def browse_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_paths:
            self.process_excel_files(file_paths)
    
    
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



    def calculate_room_revenue(self,df, time_period):
        
        # Extract values from the dataframe based on the time_period
        if time_period == "Daily":
            nett_column = "Nett Day"
        elif time_period == "Monthly":
            nett_column = "Nett Year"
        else:
            return None

        # Get values for each row
        room_revenue = df[df["Particulars"] == "Room Revenue"][nett_column].sum()
        
        # Ensure the allowance is always positive
        room_revenue_allowance =(df[df["Particulars"] == "Room Revenue - Allowance"][nett_column].sum())

        # Check if "Room Revenue - No Show" exists in the DataFrame
        if "Room Revenue -  No Show" in df["Particulars"].values:
            room_revenue_no_show = df[df["Particulars"] == "Room Revenue -  No Show"][nett_column].sum()
       
        else:
            room_revenue_no_show = 0

        # Calculate total room revenue based on the formula
        total_room_revenue = room_revenue + (room_revenue_allowance) + room_revenue_no_show
        
        return total_room_revenue


    def process_excel_files(self,file_paths):
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
                    self.row_selectors[file_path] = row_selector

            # Button to trigger column selection
            show_columns_button = ttk.Button(root, text="Select Columns", command=show_selected_columns)
            show_columns_button.pack()

            if "Night Audit Report.xls" in file_path:
                # Display preset columns in a message box
                preset_columns = self.preset_columns_NAR
                preset_rows_values = self. preset_rows_NAR_values  # Use preset rows
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
                    selected_data = row_selector.df.iloc[row_selector.df[row_selector.df.iloc[:, 0].isin(preset_rows_values)].index.tolist()][selected_columns]
                    
                    daily_revenue = self.calculate_room_revenue(selected_data, "Daily")
                    monthly_revenue = self.calculate_room_revenue(selected_data, "Monthly")

                    result_text = tk.Text(root, height=5, width=50)
                    result_text.pack()
                    result_text.insert(tk.END, f"\n\nDaily Room Revenue: {daily_revenue}")
                    result_text.insert(tk.END, f"\nMonthly Room Revenue: {monthly_revenue}")
                    send_message = messagebox.askyesno("Send WhatsApp Message", "Do you want to send a WhatsApp message with the revenue details?")
                    if send_message:
                        message_body = f"Daily Room Revenue: {daily_revenue}\nMonthly Room Revenue: {monthly_revenue}"
                        send_whatsapp_message(message_body)
            
def show_splash_screen(root, duration):
    """
    Display a splash screen for the given duration (in milliseconds).
    """
    splash = tk.Toplevel(root)
    splash.geometry("600x338")  # Adjust the size as needed
    splash.title("Loading...")
    splash.overrideredirect(True)  # Remove window decorations

    # Load the image using PIL
    image = Image.open(".\Assets\Excel Data Selector.jpg")

    # Resize the image to fit the container
    image = image.resize((600, 338))

    photo = ImageTk.PhotoImage(image)
    label = tk.Label(splash, image=photo)
    label.image = photo
    label.pack(fill=tk.BOTH, expand=True)  # Fill the container

    # Center the splash screen
    splash.update()  # Update splash window to get accurate width and height
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width - splash.winfo_width()) // 2
    y = (screen_height - splash.winfo_height()) // 2
    splash.geometry("+%d+%d" % (x, y))

    # Destroy the splash screen after the duration
    root.after(duration, splash.destroy)

def send_whatsapp_message(message_body):
    account_sid = os.environ.get('TWILIO_ACCOUNT_SID')
    auth_token = os.environ.get('TWILIO_AUTH_TOKEN')
    print(account_sid,auth_token)
    if not account_sid or not auth_token:
        print("Error: Twilio credentials not found in environment variables.")
        return
    
    client = Client(account_sid, auth_token)

    message = client.messages.create(
        body=message_body,
        from_='whatsapp:+14155238886',
        to='whatsapp:+23057568744'
    )
    print(message.sid)


def start_main_app(root):
    app = App(root)
    root.deiconify()  # Show the main window
          
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    show_splash_screen(root, duration=3000)  
    root.after(3000, start_main_app, root)  # Schedule the creation of the main application window
    root.mainloop()