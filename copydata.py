import numbers
import win32com.client as win32
import sys

def creer_feuille_chiffrage(chemin_fichier, nom_feuille):
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
        new_sheet.Name = f"Chiffrage {nom_feuille.split(' ')[1]}"

        # Effacer les données de la colonne C
        new_sheet.Range("C:C").ClearContents()

        new_sheet.Range("C14").Value = "Unité"
        new_sheet.Range("D14").Value = "Qté"
        new_sheet.Range("E14").Value = "Prix HT Unitaire"
        new_sheet.Range("F14").Value = "Montant HT"
        new_sheet.Range("G14").Value = "Taux TVA"
        new_sheet.Range("H14").Value = "Montant TTC"

        new_sheet.Range("F15:F300").Value = "=E15*D15"
        new_sheet.Range("H15:H300").Value = "=(F15*(1+G15))"

        print("Formules insérées dans les colonnes F et H.")
        # Insérer une ligne de total sur la ligne 15
        new_sheet.insert_rows(15)
        new_sheet.Range("A15").Value = "Total"
        new_sheet.Range("F15").Formula = "=SUM(F16:F300)"
        new_sheet.Range("H15").Formula = "=SUM(H16:H300)"




    except Exception as e:
        print(f"Une erreur est survenue : {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python copydata.py <chemin_fichier_excel> <nom_feuille>")
        sys.exit(1)

    param_chemin_fichier_excel = sys.argv[1]
    param_nom_feuille = sys.argv[2]

    creer_feuille_chiffrage(param_chemin_fichier_excel, param_nom_feuille)
