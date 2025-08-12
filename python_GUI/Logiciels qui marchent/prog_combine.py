import tkinter as tk
from tkinter import filedialog
import pandas as pd
import os

def choose_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    return file_path

def show_message(message):
    message_window = tk.Toplevel(root)
    message_window.title("Message")

    message_label = tk.Label(message_window, text=message)
    message_label.pack()

    ok_button = tk.Button(message_window, text="OK", command=message_window.destroy)
    ok_button.pack()

def fusion_files():
    #choisir le premier ficheir
    file_path1 = choose_file()
    if not file_path1:
        return

    #choisir le deuxieme fichier
    file_path2 = choose_file()
    if not file_path2:
        return

    # Lire le premier fichier
    df1 = pd.read_excel(file_path1)
    header1 = df1.columns.tolist()

    #Lire le de le deuxieme fichier
    df2 = pd.read_excel(file_path2)
    header2 = df2.columns.tolist()

    # véerificatoin du nombre de colonnes et de leurs nomination
    if header1 == header2 and df1.shape[1] == df2.shape[1]:
        df_combined = pd.concat([df1, df2], ignore_index=True)# Combiner les tableau avec les headers
        # creer un nouvzl fichzier 
        output_file = "combinaison.xlsx"
        df_combined.to_excel(output_file, index=False)

        os.startfile(output_file)# Ovrir le nouvel fichier

        show_message("Opération réussie") #réussie
    else:
        show_message("Valeurs différentes, réessayez") #erreur

root = tk.Tk()
root.title("Fusion de deux fichiers Excel")
root.geometry("200x100")


style_path = r"C:\Users\user\Documents\PYTHON\python_GUI\Logiciels qui marchent\Forest-ttk-theme-master\Forest-ttk-theme-master\forest-dark.tcl"
root.tk.call("source", style_path)

button = tk.Button(root, text="Choisir les fichiers", command=fusion_files)
button.pack(padx=20, pady=10)

root.mainloop()
