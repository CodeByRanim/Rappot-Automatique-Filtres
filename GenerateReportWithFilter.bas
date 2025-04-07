Sub GenerateReportWithFilter()
    Dim ws As Worksheet
    Dim newWB As Workbook
    Dim newWS As Worksheet
    
    Set ws = ThisWorkbook.Sheets("Data") ' Remplacez par le nom de votre feuille de données
    Set newWB = Workbooks.Add
    Set newWS = newWB.Sheets(1)
    
    ' Appliquer un filtre pour sélectionner les données spécifiques
    ws.Range("A1:D1").AutoFilter Field:=3, Criteria1:=">100" ' Filtrer les valeurs de la colonne C > 100
    
    ' Copier les données filtrées
    ws.AutoFilter.Range.Copy
    
    ' Coller dans la nouvelle feuille
    newWS.Paste
    
    ' Enregistrer le rapport généré
    newWB.SaveAs "C:\chemin\vers\le\rapport_généré.xlsx"
    
    newWB.Close
    MsgBox "Rapport généré avec succès !"
End Sub
