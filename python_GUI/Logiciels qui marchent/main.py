import tkinter as tk
from tkinter import filedialog
import pandas as pd
import subprocess
from tkinter import ttk

root = tk.Tk()
root.title("Gestion et manip générale")
root.geometry("700x400")


# Appliquer le style forest dark
root.tk.call("source", r"C:\Users\ryadk\OneDrive\Documents\flash red\APPLICATION\python_GUI\Logiciels qui marchent\Forest-ttk-theme-master\Forest-ttk-theme-master\forest-dark.tcl")

style = ttk.Style(root)  # Configuration du style
style.theme_use("forest-dark")
style.configure(".", font=("Arial", 10))


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        #lire le fichier excel et l'afficher
        df = pd.read_excel(file_path)
        header = df.columns.tolist()
        data = df.values.tolist()
        create_table(header, data)


def run_file(file_path):
    subprocess.run(["python", file_path])


def create_table(header, data):
    separator = ttk.Separator(root, orient="vertical")
    separator.grid(row=0, column=1, rowspan=7, padx=10, pady=5, sticky="ns")

    table_frame = tk.Frame(root)
    table_frame.grid(row=0, column=2, rowspan=7, padx=10, pady=5, sticky="nsew")

    treeview = ttk.Treeview(table_frame, columns=header, show="headings")

    # Configure treeview columns width
    for col in header:
        treeview.column(col, width=100, anchor="center")
        treeview.heading(col, text=col)

    # Add vertical scrollbar
    scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=treeview.yview)
    scrollbar.pack(side="right", fill="y")
    treeview.configure(yscrollcommand=scrollbar.set)

    # Insert data into the treeview
    for row in data:
        treeview.insert("", tk.END, values=row)

    treeview.pack(fill="both", expand=True)


#bouton: Select File
select_file_btn = tk.Button(root, text="Sélectionner le fichier", command=select_file)
select_file_btn.grid(row=1, column=0, padx=10, pady=5, sticky="w")

#bouton: Doublons
doublons_btn = tk.Button(root, text="Trouver doublons",
                         command=lambda: run_file(r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\reperage_doublons.py"))
doublons_btn.grid(row=2, column=0, padx=10, pady=5, sticky="w")

#bouton: Fusion
fusion_btn = tk.Button(root, text="Fusionner deux fichiers",
                       command=lambda: run_file(r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\prog_combine.py"))
fusion_btn.grid(row=3, column=0, padx=10, pady=5, sticky="w")

#bouton: Total
total_btn = tk.Button(root, text="Total Global",
                      command=lambda: run_file(r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\prog_totale_glob.py"))
total_btn.grid(row=4, column=0, padx=10, pady=5, sticky="w")

# bouton: Tri
tri_btn = tk.Button(root, text="Trier les valeurs",
                    command=lambda: run_file(r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\prog_ranking.py"))
tri_btn.grid(row=5, column=0, padx=10, pady=5, sticky="w")

#bouton: Highlight des vides
highlight_btn = tk.Button(root, text="Repérage des vides",
                          command=lambda: run_file(r"C:/Users/user/Documents/PYTHON/python_GUI/Logiciels qui marchent/prog_cellules_vides.py"))
highlight_btn.grid(row=6, column=0, padx=10, pady=5, sticky="w")

#bouton: Insertion
insertion_btn = tk.Button(root, text="Insertion",
                          command=lambda: run_file(r"C:/Users/user/Documents/PYTHON/python_GUI/Logiciels qui marchent/programme_insertion/program_insert.py"))
insertion_btn.grid(row=7, column=0, padx=10, pady=5, sticky="w")

root.mainloop()
