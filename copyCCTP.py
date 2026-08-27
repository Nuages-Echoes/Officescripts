import win32com.client as win32
import sys

import win32com.client as win32

def creer_feuille_CCTP(chemin_fichier, nom_feuille):
    excel = None
    workbook = None
    try:
        # Se connecter à l'instance existante de Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        # excel.Visible = True  # Décommente pour voir Excel

        # Ouvrir le fichier Excel
        workbook = excel.Workbooks.Open(chemin_fichier)
        print(f"Fichier Excel ouvert : {chemin_fichier}")

        # Copier la feuille
        sheet = workbook.Sheets(nom_feuille)
        sheet.Copy(After=workbook.Sheets(workbook.Sheets.Count))
        new_sheet = workbook.Sheets(workbook.Sheets.Count)
        new_sheet.Name = f"CCTP {nom_feuille.split(' ')[1]}"
        print(f"Feuille copiée et renommée en : {new_sheet.Name}")

        # Effacer les données de la colonne C
        new_sheet.Range("C:C").ClearContents()
        print("Colonne C effacée.")

        workbook.Save()  # Enregistrer les modifications
        print("Modifications enregistrées.")

    except Exception as e:
        print(f"Une erreur est survenue : {e}")


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python copydata.py <chemin_fichier_excel> <nom_feuille>")
        sys.exit(1)

    param_chemin_fichier_excel = sys.argv[1]
    param_nom_feuille = sys.argv[2]

    creer_feuille_CCTP(param_chemin_fichier_excel, param_nom_feuille)
