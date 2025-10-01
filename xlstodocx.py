import pandas as pd
from docx import Document
from docx.shared import RGBColor
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
        
        if (str(valeur[2]) != 'nan'):  # Vérifier si la troisième colonne n'est pas vide
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[2]))
            run.bold = True  # Mettre le texte en gras
            run.font.color.rgb = RGBColor(255, 0, 0)  # Changer la couleur du texte en rouge
        match valeur[0]:
            case 'Titre 1':
                doc.add_paragraph(str(valeur[1]), style='Heading 1')
            case 'Titre 2':
                doc.add_paragraph(str(valeur[1]), style='Heading 2')
            case 'Titre 3':
                doc.add_paragraph(str(valeur[1]), style='Heading 3')
            case _ :
                doc.add_paragraph(str(valeur[1]), style='Normal')
        if "rouge" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.color.rgb = RGBColor(255, 0, 0)  # Changer la couleur du texte en rouge
        elif "bleu" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.color.rgb = RGBColor(0, 0, 255)  # Changer la couleur du texte en bleu
        elif "vert" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.color.rgb = RGBColor(0, 255, 0)  # Changer la couleur du texte en vert
        elif "violet" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.color.rgb = RGBColor(128, 0, 128)  # Changer la couleur du texte en violet
        elif "orange" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.color.rgb = RGBColor(255, 165, 0)  # Changer la couleur du texte en orange
        elif "barre" in str(valeur[0]).lower():
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(str(valeur[1]))
            run.font.strike = True  # Appliquer le style barré


    # Sauvegarder le document Word
    doc.save(word_en_sortie)

def lire_donnees_client_excel(param_chemin_fichier_excel, param_nom_feuille):
    # Lire le fichier Excel avec pandas
    df = pd.read_excel(param_chemin_fichier_excel, sheet_name=param_nom_feuille)

     # Créer une liste pour stocker les valeurs des cellules D10 à D12
    valeurs = []

    # Lire les valeurs des cellules D4 à D12
    for i in range(9, 13):  # Les lignes 10 à 13 (inclus)
        valeur = df.iloc[i - 1, ]
        valeurs.append(valeur)

    # Lire le nom du maître d'ouvrage dans la cellule F10
    #nom_maitre_ouvrage = df.iloc[9 , 5]
    #valeurs.append(nom_maitre_ouvrage)

    return valeurs

def supprimer_dossier_temp(nom_dossier):
    # Obtenir le chemin du répertoire %TEMP%
    temp_dir = os.environ.get('TEMP')

    # Construire le chemin complet du dossier à supprimer
    chemin_dossier = os.path.join(temp_dir, nom_dossier)

    # Vérifier si le dossier existe
    if os.path.exists(chemin_dossier):
        try:
            # Supprimer le dossier et tout son contenu
            shutil.rmtree(chemin_dossier)
            print(f"Le dossier {chemin_dossier} a été supprimé avec succès.")
        except Exception as e:
            print(f"Une erreur est survenue lors de la suppression du dossier : {e}")
    else:
        print(f"Le dossier {chemin_dossier} n'existe pas.")

def mise_a_jour_signets(word_en_sortie, donnees_client):
    
    try:
        # Créer une instance de Word
        word_app = win32.gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False  # Ne pas rendre Word visible (pour un traitement en arrière-plan)
    except Exception as e:
        # supprimer le cache et réessayer
        supprimer_dossier_temp('gen_py')
        word_app = win32.gencache.EnsureDispatch('Word.Application')   
        word_app.Visible = False  # Ne pas rendre Word visible (pour un traitement en arrière-plan)
    
    try:
        # Ouvrir le document Word
        doc = word_app.Documents.Open(word_en_sortie)

        # Vérifier si le signet "AdresseProjet" existe et le remplir
        if doc.Bookmarks.Exists("AdresseProjet"):
            doc.Bookmarks("AdresseProjet").Range.Text = str(donnees_client[1])  # Valeur de D11

        # Sauvegarder et fermer le document Word
        doc.Save()
        doc.Close()
    except Exception as e:
        print(f"Une erreur est survenue : {e}")
    finally:
        # Quitter Word
        word_app.Quit()

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



    # Lire les 3 premières colonnes du fichier Excel
    donnees_premiere_colonne = lire_trois_premieres_colonnes_excel(param_chemin_fichier_excel, param_nom_feuille)

    # Lire les informations client du fichier Excel
    donnees_client = lire_donnees_client_excel(param_chemin_fichier_excel, param_nom_feuille)
    
    # Afficher les valeurs
    print(f"Les valeurs des cellules D10 à D12 sont : {donnees_client}")


    # Ajouter les données dans le fichier Word existant
    ajouter_dans_fichier_word(word_en_sortie, donnees_premiere_colonne)

    # Mettre à jour les signets du document Word
    mise_a_jour_signets(word_en_sortie, donnees_client)

    
    print(f"Les données ont été ajoutées à {word_en_sortie}")
