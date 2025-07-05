
Sub AppelerScriptPython()
    Dim cheminScriptPython As String
    Dim cheminFichierExcel As String
    Dim commandeShell As String

    ' Obtenir le chemin du fichier Excel actuel
    cheminFichierExcel = ThisWorkbook.Path

    ' Vérifier si le chemin est valide
    If cheminFichierExcel = "" Then
        MsgBox "Veuillez enregistrer le fichier Excel avant d'exécuter ce script.", vbExclamation
        Exit Sub
    End If

    ' Chemin vers le script Python (supposé être dans le même dossier que le fichier Excel)
    cheminScriptPython = cheminFichierExcel & "\bibExcel.py"

    ' Construire la commande Shell pour exécuter le script Python
    commandeShell = "python """ & cheminScriptPython & """"

    ' Exécuter le script Python
    Shell commandeShell, vbNormalFocus
End Sub