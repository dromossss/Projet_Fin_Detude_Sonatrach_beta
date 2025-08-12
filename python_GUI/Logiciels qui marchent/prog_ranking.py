import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    return file_path

def sort_data():
    # Choose the Excel file
    file_path = choose_file()
    if not file_path:
        return

    # Read the Excel file
    df = pd.read_excel(file_path)
    headers = df.columns.tolist()

    # Create a new tkinter window for sorting options
    sort_window = tk.Tk()
    sort_window.title("Options de tri")

    def sort_and_save(sort_type):
        # Sort the data based on user's choice
        if sort_type == "decroissant":
            df_sorted = df.sort_values(by=headers, ascending=False)
        elif sort_type == "croissant":
            df_sorted = df.sort_values(by=headers, ascending=True)
        else:
            print("Invalid sort type.")
            sort_window.destroy()
            return

        # Save the sorted data to the same Excel file
        df_sorted.to_excel(file_path, index=False)

        # Close the sort window
        sort_window.destroy()

        # Open the modified Excel file using the OS library
        os.startfile(file_path)

    # Load the theme style from the specified file
    style_path = r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\Forest-ttk-theme-master\Forest-ttk-theme-master\forest-dark.tcl"
    sort_window.tk.call("source", style_path)

    # Create checkboxes for sort options
    decroissant_checkbox = tk.Checkbutton(sort_window, text="Ordre décroissant", command=lambda: sort_and_save("decroissant"))
    decroissant_checkbox.pack(padx=10, pady=5)

    croissant_checkbox = tk.Checkbutton(sort_window, text="Ordre croissant", command=lambda: sort_and_save("croissant"))
    croissant_checkbox.pack(padx=10, pady=5)

    sort_window.mainloop()

root = tk.Tk()
root.title("Programme de Tri")
# Apply the style from the specified file
root.tk.call("source", r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\Forest-ttk-theme-master\Forest-ttk-theme-master\forest-dark.tcl")

button = tk.Button(root, text="Choisir le fichier", command=sort_data)
button.pack(padx=20, pady=10)

root.mainloop()
