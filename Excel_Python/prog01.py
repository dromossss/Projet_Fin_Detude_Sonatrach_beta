import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Importer la bibliothèque pandas pour la manipulation des données
# Importer la bibliothèque tkinter pour l'interface utilisateur
# Importer la boîte de dialogue filedialog de tkinter

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

# Combiner les DataFrames verticalement
combined_df = pd.concat([df1, df2], ignore_index=True)

# Définir une fonction pour remplacer les cellules vides par 'VIDE'
def replace_empty_cells(df):
    df.fillna('VIDE', inplace=True)

# Appeler la fonction pour remplacer les cellules vides par 'VIDE' 
replace_empty_cells(combined_df)

# Afficher le DataFrame combiné
print(combined_df)



# Sauvegarder le DataFrame combiné dans un nouveau fichier Excel
combined_df.to_excel('combined_file.xlsx', index=False)
