import custom_tkinter as tk
import subprocess

# Fonction pour exécuter le programme sélectionné
def run_program():
    selected_program = program_combobox.get()
    
    # Exécute le programme en utilisant subprocess et récupère la sortie
    result = subprocess.run(["python", selected_program], capture_output=True, text=True)
    
    # Affiche la sortie dans la zone de texte
    output_text.delete("1.0", tk.END)
    output_text.insert(tk.END, result.stdout)

# Liste des programmes à exécuter
program_list = ["program1.py", "program2.py", "program3.py"]

# Crée la fenêtre principale
root = tk.tkinter.Tk()
root.title("Interface d'exécution de programmes")

# Crée une frame pour les éléments de l'interface
frame = tk.Frame(root, padding="20")
frame.pack()

# Crée une liste déroulante pour sélectionner le programme à exécuter
program_label = tk.Label(frame, text="Sélectionner le programme :")
program_label.grid(row=0, column=0, sticky="w")

program_combobox = tk.Combobox(frame, values=program_list)
program_combobox.grid(row=0, column=1, pady=10)

# Crée un bouton pour exécuter le programme sélectionné
run_button = tk.Button(frame, text="Exécuter", command=run_program)
run_button.grid(row=1, column=0, columnspan=2)

# Crée une zone de texte pour afficher la sortie du programme
output_label = tk.Label(frame, text="Sortie du programme :")
output_label.grid(row=2, column=0, sticky="w")

output_text = tk.Text(frame, height=10, width=50)
output_text.grid(row=3, column=0, columnspan=2, pady=10)

root.mainloop()
