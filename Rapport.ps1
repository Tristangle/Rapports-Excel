#Objectif » Récupérer les noms et champs de tables dans un Excel, puis générer un nouvel Excel avec un Script PowerShell

# Acceder à un dossier
$emplacement = Set-Location -Path "C:\Users\tb50919\Documents\databaseExport"
# Lister les fichiers présents
$dossier = Get-ChildItem($emplacement)

#Write-Output($liste)

# Créer le fichier Excel de rapport
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $True
$excel.DisplayAlerts = $False

# Ouvre une page dans le fichier de Rapport
$Rapport = $excel.Workbooks.add()
$FeuilleRapport = $Rapport.worksheets.Item(1)
# Boucle sur la liste des Excels présents dans le dossier
$emplacementNom = 1
$emplacementChamps = 3
foreach($fichier in $dossier){
    # Ouvrir le fichier existant
    $nomFichier = $fichier.FullName
    $fichierExcel = $excel.Workbooks.Open($nomFichier)

    # Copier le nom du fichier, puis le coller dans le Rapport
    $FeuilleRapport.Cells.Item(1, $emplacementNom) = $fichier.Name
    # Accéder à la feuille de calcul du fichier Excel ouvert
    $feuilleExcel = $fichierExcel.Worksheets.Item(1)
    # Boucle sur la liste des champs de la table
    foreach ($cellule in $feuilleExcel.UsedRange.Columns) {
        $champ = $cellule.Cells.Item(1, 1).Value2
        # Copier les champs de table puis les coller dans l'Excel de rapport
        $FeuilleRapport.Cells.Item($emplacementChamps, $emplacementNom) = $champ
        $emplacementChamps++
    }
    $emplacementChamps = 3
    $emplacementNom = $emplacementNom +=2


    $fichierExcel.Close()
}
#Redimensionner la cellule pour avoir un résultat lisible 
$FeuilleRapport.Columns.AutoFit()
# Enregistre le fichier Excel dans le dossier
#$Rapport.SaveAs($emplacement)
#$Rapport.Close()
#$excel.Quit()
