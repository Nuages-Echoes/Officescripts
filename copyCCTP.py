import win32com.client as win32
import sys

def creer_feuille_CCTP(chemin_fichier, nom_feuille):
    try:
        # Se connecter à l'instance existante de Excel
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = True  # Rendre Excel visible

        # Ouvrir le fichier Excel
        workbook = excel.Workbooks.Open(chemin_fichier)

        # Copier la feuille
        sheet = workbook.Sheets(nom_feuille)
        sheet.Copy(After=workbook.Sheets(workbook.Sheets.Count))
        new_sheet = workbook.Sheets(workbook.Sheets.Count)
        new_sheet.Name = f"CCTP {nom_feuille.split(' ')[1]}"

        # Effacer les données de la colonne C
        new_sheet.Range("C:C").ClearContents()

        


    except Exception as e:
        print(f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python copydata.py <chemin_fichier_excel> <nom_feuille>")
        sys.exit(1)

    param_chemin_fichier_excel = sys.argv[1]
    param_nom_feuille = sys.argv[2]

    creer_feuille_CCTP(param_chemin_fichier_excel, param_nom_feuille)
