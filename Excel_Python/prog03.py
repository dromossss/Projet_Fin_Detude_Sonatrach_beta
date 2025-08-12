import openpyxl
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename

import openpyxl
from tkinter import Tk, simpledialog
from tkinter.filedialog import askopenfilename, asksaveasfilename

def choisir_fichier():
    Tk().withdraw()  # Hide the main Tkinter window
    fichier = askopenfilename(title="Choisir un fichier Excel", filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
    return fichier

def choisir_colonne():
    root = Tk()
    root.withdraw()  # Hide the main Tkinter window
    colonne = simpledialog.askstring("Choisir une colonne", "Sélectionnez une colonne (par exemple, A, B, C, ...)")
    return colonne

def is_numeric(value):
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False

fichier = choisir_fichier()
colonne = choisir_colonne()

wb = openpyxl.load_workbook(fichier)
feuille = wb['data']

# Retrieve the values from the selected column, ignoring non-numeric values
valeurs_colonne = [cellule.value for cellule in feuille[colonne] if is_numeric(cellule.value)]

# Sum the values in the column
total = sum(valeurs_colonne)

# Find the last row in the column
derniere_ligne = feuille.max_row

# Add the total to the "Total Global" cell at the end of the column
cellule_total = feuille[f"{colonne}{derniere_ligne+1}"]
cellule_total.value = total
cellule_total.number_format = "#,##0.00"  # Number format with two decimal places

# Save the changes to a new file
save_path = asksaveasfilename(title="Enregistrer le fichier modifié", defaultextension=".xlsx", filetypes=[("Fichiers Excel", "*.xlsx")])
wb.save(save_path)
