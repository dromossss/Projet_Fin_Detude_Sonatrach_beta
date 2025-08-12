import pandas as pd
import tkinter as tk
from tkinter import filedialog
from openpyxl.styles import PatternFill

# Create the main Tkinter window
root = tk.Tk()
root.withdraw()

# Ask the user to choose an Excel file
file_path = filedialog.askopenfilename(title='Sélectionner un fichier Excel')

# Read the Excel file into a DataFrame
df = pd.read_excel(file_path)

# Function to search for duplicates in a column and highlight them
def highlight_duplicates(df, column):
    # Create a copy of the DataFrame
    df_copy = df.copy()

    # Search for duplicates in the specified column
    duplicates = df_copy[column].duplicated(keep=False)

    # Highlight duplicates in the column with a color
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for idx, is_duplicate in enumerate(duplicates):
        if is_duplicate:
            df_copy.loc[idx, column] = df_copy.loc[idx, column].style.apply(fill)

    # Save the modified DataFrame to a new Excel file
    save_path = filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel Files', '*.xlsx')])
    df_copy.to_excel(save_path, index=False)

# Create a Tkinter window to enter the column name
window = tk.Tk()

def process_input():
    # Get the column name entered by the user
    colonne = entry.get()

    # Call the function to highlight duplicates in the specified column
    highlight_duplicates(df, colonne)

    # Close the Tkinter window
    window.destroy()

# Label and Entry widget to get the column name
label = tk.Label(window, text="Entrez le nom de la colonne à analyser:")
label.pack()

entry = tk.Entry(window)
entry.pack()

button = tk.Button(window, text="Valider", command=process_input)
button.pack()

window.mainloop()
