import openpyxl
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkinter import filedialog


# Créer une fenêtre principale Tkinter
root = tk.Tk()
root.withdraw()

# Deamnder à l'utilisateur de sélectionner le premier fichier Excel
file1_path = filedialog.askopenfilename(title='Sélectionner le premier fichier Excel')

# Demander à l'utilisateur de sélectionner le deuxième fichier Excel
file2_path = filedialog.askopenfilename(title='Sélectionner le deuxième fichier Excel')

# Lire les fichiers Excel sélectionnés dans des DataFrames distincts
df1 = pd.read_excel(file1_path)
df2 = pd.read_excel(file2_path)