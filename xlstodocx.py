import pandas as pd
from docx import Document
import win32com.client as win32
import shutil
import os

def lire_trois_premieres_colonnes_excel(chemin_fichier_excel, nom_feuille):
    # Lire le fichier Excel avec pandas, en sautant les 14 premières lignes
    df = pd.read_excel(chemin_fichier_excel, sheet_name=nom_feuille, skiprows=13)

    # Extraire les trois premières colonnes
    trois_premieres_colonnes = df.iloc[:, :3]  # Sélectionne les trois premières colonnes

    # Retourner les données sous forme de liste de listes
    return trois_premieres_colonnes.values.tolist()

def conversion_style(style_excel):
    # Dictionnaire de conversion des styles Excel vers les styles Word
    conversion_dict = {
        'Titre 1': 'Heading 1',
        'Titre 2': 'Heading 2',
        'Titre 3': 'Heading 3',
        'Normal': 'Normal',
        # Ajoutez d'autres conversions si nécessaire
    }
    return conversion_dict.get(style_excel, 'Normal')  # Retourne 'Normal' par défaut si le style n'est pas trouvé
    

def ajouter_dans_fichier_word(word_en_sortie, donnees):
   
    # Charger le document Word existant
    doc = Document(word_en_sortie)

    # Ajouter chaque élément de la première colonne au document Word
    for valeur in donnees:
        doc.add_paragraph(str(valeur[1]), style=conversion_style(str(valeur[0])))

    # Sauvegarder le document Word
    doc.save(word_en_sortie)

#def mise_a_jour_signets(word_en_sortie, chemin_fichier_excel, param_nom_feuille):
#    # Lire le fichier Excel avec pandas
#    df = pd.read_excel(chemin_fichier_excel, sheet_name=param_nom_feuille)

#    # Charger le document Word existant
#    doc = Document(word_en_sortie)

#    Adresse= df.at[7, 3]
#     # Remplir les signets avec les valeurs des cellules Excel
#    if doc.Bookmarks.Exists("Adresse"):
#        doc.Bookmarks("Adresse").Range.Text = str(Adresse)


if __name__ == "__main__":
    if len(os.sys.argv) != 3:
        print("Usage: python xlstodocx.py <chemin_fichier_excel> <nom_feuille>")
        os.sys.exit(1)

    param_chemin_fichier_excel = os.sys.argv[1]
    param_nom_feuille = os.sys.argv[2]

    # Chemins des fichiers
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
    donnees_premiere_colonne = lire_trois_premieres_colonnes_excel(param_chemin_fichier_excel, param_nom_feuille)

    # Ajouter les données dans le fichier Word existant
    ajouter_dans_fichier_word(word_en_sortie, donnees_premiere_colonne)

    # Mettre à jour les signets du document Word
    # mise_a_jour_signets(word_en_sortie, param_chemin_fichier_excel, param_nom_feuille)

    
    print(f"Les données ont été ajoutées à {word_en_sortie}")
