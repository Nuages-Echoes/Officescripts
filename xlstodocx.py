import pandas as pd
from docx import Document
import shutil
import os

def lire_trois_premieres_colonnes_excel(chemin_fichier_excel, nom_feuille):
    # Lire le fichier Excel avec pandas, en sautant les 14 premières lignes
    df = pd.read_excel(chemin_fichier_excel, sheet_name=nom_feuille, skiprows=13)

    # Extraire les trois premières colonnes
    trois_premieres_colonnes = df.iloc[:, :3]  # Sélectionne les trois premières colonnes

    # Retourner les données sous forme de liste de listes
    return trois_premieres_colonnes.values.tolist()

def ajouter_dans_fichier_word(word_en_sortie, donnees):
   
    # Charger le document Word existant
    doc = Document(word_en_sortie)

    # Ajouter chaque élément de la première colonne au document Word
    for valeur in donnees:
        if str(valeur[0]) == 'nan':
            doc.add_paragraph(str(valeur[1]))
        else: 
            doc.add_paragraph(str(valeur[1]), style=str(valeur[0]))

    # Sauvegarder le document Word
    doc.save(word_en_sortie)

# Chemins des fichiers
chemin_fichier_excel = r'C:\Users\maxim\VSCodeProject\Officescripts\TableauEntree.xlsm'
nom_feuille = 'AVP'  # Remplacez par le nom de votre feuille Excel
word_en_entree = r'C:\Users\maxim\VSCodeProject\Officescripts\DSTest.docx'  # Remplacez par le chemin de votre fichier Word
word_en_sortie = r'C:\Users\maxim\VSCodeProject\Officescripts\sortie.docx'  # Remplacez par le chemin de votre fichier Word

try:
    # Vérifier si le fichier source existe
    if not os.path.exists(word_en_entree):
        raise FileNotFoundError(f"Le fichier source {word_en_entree} n'existe pas.")

    # Copier le fichier
    shutil.copy(word_en_entree, word_en_sortie)
    print(f"Le fichier a été copié de {word_en_entree} vers {word_en_sortie}")

except Exception as e:
    print(f"Une erreur est survenue : {e}")   



# Lire la première colonne du fichier Excel
donnees_premiere_colonne = lire_trois_premieres_colonnes_excel(chemin_fichier_excel, nom_feuille)

# Ajouter les données dans le fichier Word existant
ajouter_dans_fichier_word(word_en_sortie, donnees_premiere_colonne)

print(f"Les données ont été ajoutées à {word_en_sortie}")
