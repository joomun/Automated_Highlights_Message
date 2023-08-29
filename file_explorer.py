import tkinter as tk
from tkinter import filedialog
import pandas as pd

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        excel_file_path.set(file_path)
        update_display()

def read_excel():
    path = excel_file_path.get()
    if path:
        try:
            df = pd.read_excel(path)
            result_text.config(state=tk.NORMAL)
            result_text.delete("1.0", tk.END)
            result_text.insert(tk.END, df.head().to_string(index=False))
            result_text.config(state=tk.DISABLED)
        except Exception as e:
            result_text.config(state=tk.NORMAL)
            result_text.delete("1.0", tk.END)
            result_text.insert(tk.END, f"Error: {str(e)}")
            result_text.config(state=tk.DISABLED)

def update_display():
    selected_path.config(text=excel_file_path.get())

app = tk.Tk()
app.title("Excel Reader")

excel_file_path = tk.StringVar()

browse_button = tk.Button(app, text="Browse", command=browse_file)
browse_button.pack()

selected_path = tk.Label(app, text="")
selected_path.pack()

read_button = tk.Button(app, text="Read Excel", command=read_excel)
read_button.pack()

result_text = tk.Text(app, state=tk.DISABLED)
result_text.pack()

app.mainloop()
