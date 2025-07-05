import tkinter as tk
from tkinter import filedialog, messagebox

def select_file():
    # Ouvrir une boîte de dialogue pour sélectionner un fichier
    file_path = filedialog.askopenfilename()
    if file_path:
        file_entry.delete(0, tk.END)  # Effacer le contenu actuel
        file_entry.insert(0, file_path)  # Insérer le chemin du fichier sélectionné

def submit_form():
    # Récupérer les valeurs des champs
    file_path = file_entry.get()
    text_content = text_entry.get("1.0", tk.END).strip()  # Récupérer le texte et supprimer les espaces inutiles

    # Afficher les valeurs dans une boîte de message
    messagebox.showinfo("Information", f"Fichier Excel sélectionné : {file_path}\nNom de la Feuille : {text_content}")

# Créer la fenêtre principale
root = tk.Tk()
root.title("Génération du rapport d'avancement du chantier")

# Champ pour sélectionner un fichier
file_label = tk.Label(root, text="Fichier Excel sélectionné :", anchor="w")
file_label.pack(pady=5)

file_entry = tk.Entry(root, width=50)
file_entry.pack(pady=5)

file_button = tk.Button(root, text="Parcourir", command=select_file)
file_button.pack(pady=5)

# Champ pour saisir du texte
text_label = tk.Label(root, text="Feuille à transformer en rapport Word :")
text_label.pack(pady=5)

text_entry = tk.Text(root, width=15, height=1)
text_entry.pack(pady=5)

# Bouton pour soumettre le formulaire
submit_button = tk.Button(root, text="Soumettre", command=submit_form)
submit_button.pack(pady=20)

# Lancer la boucle principale de l'application
root.mainloop()
