#Objectif » Récupérer les noms et champs de tables dans un Excel, puis générer un nouvel Excel avec un Script PowerShell

# Acceder à un dossier
$emplacement = Set-Location -Path "F:\08-DAFI\06-Informatique\01. Projet IFS\_technique\database.export"
# Lister les fichiers présents
$liste = Get-ChildItem($emplacement)
Write-Output($liste)
# Créer le fichier Excel de rapport
$FichierRapport = New-Object -ComObject Excel.Application
$FichierRapport.Visible = $True
# Ouvre une page dans le fichier de Rapport
$Rapport = $FichierRapport.Workbooks.add()
$FeuilleRapport = $Rapport.worksheets.Item(1)
# Boucle sur la liste des Excels présents dans le dossier

    # Ouvrir le fichier

    # Copier le nom du fichier, puis le coller dans le Rapport

        # Boucle sur la liste des champs de la table

            # Copier les champs de table puis les coller dans l'Excel

# Enregistre le fichier Excel dans le dossier