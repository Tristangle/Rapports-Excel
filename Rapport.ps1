#Objectif » Récupérer les noms et champs de tables dans un Excel, puis générer un nouvel Excel avec un Script PowerShell


# Définir l'emplacement dans une fênetre parcourir 
 Add-Type -AssemblyName System.Windows.Forms
 $browser = New-Object System.Windows.Forms.FolderBrowserDialog

 $null = $browser.ShowDialog()
# Enregistrer le chemin choisi
 $cheminDossier= $browser.SelectedPath

# Lister les fichiers présents, en excluant les dossier
$dossier = Get-ChildItem -Path $cheminDossier -Filter "*.xlsx" | Where-Object { !$_.PSIsContainer }

# Créer un objet Excel afin d'utiliser les fonctions associées à Excel
$excel = New-Object -ComObject Excel.Application

# Permet de voir l'excel, sans cela l'utilisateur ne verra pas les excels
$excel.Visible = $false

# Permet d'éviter les confirmations manuelles, comme pour la fermeture d'un Excel
$excel.DisplayAlerts = $False

# Créer un excel et une feuille excel
$Rapport = $excel.Workbooks.add()
$FeuilleRapport = $Rapport.worksheets.Item(1)

# Définit les valeurs de positions
$emplacementNom = 1
$emplacementChamps = 3

# Boucle sur la liste des Excels présents dans le dossier
foreach($fichier in $dossier){

    # Ouvrir le fichier existant
    $nomFichier = $fichier.FullName
    $fichierExcel = $excel.Workbooks.Open($nomFichier)

    # Copier le nom du fichier sans extensions, puis le coller dans le Rapport
    $nomSansExtension = [System.IO.Path]::GetFileNameWithoutExtension($fichier.Name)
    $FeuilleRapport.Cells.Item(1, $emplacementNom) = $nomSansExtension

    # Accéder à la feuille de calcul du fichier Excel ouvert
    $feuilleExcel = $fichierExcel.Worksheets.Item(1)

    # Boucle sur la liste des champs de la table
    foreach ($cellule in $feuilleExcel.UsedRange.Columns) {

        #Copier la valeur brute de la cellule sans tenir compte du format
        $champ = $cellule.Cells.Item(1, 1).Value2

        # Copier les champs de table puis les coller dans l'Excel de rapport
        $FeuilleRapport.Cells.Item($emplacementChamps, $emplacementNom) = $champ

        #Incrémenter la nouvelle position du champs de table vide
        $emplacementChamps++
    }

    #Réinitialise la position du champs au départ
    $emplacementChamps = 3

    # Actualise la nouvelle position pour le nom
    $emplacementNom = $emplacementNom +=2

    #Ferme le fichier
    $fichierExcel.Close()
}
#Redimensionner les cellules pour avoir un résultat lisible 
$FeuilleRapport.Columns.AutoFit()

#Récupère la date et la transforme en date française 
$dateActuelle = Get-Date
$dateFrancaise = $dateActuelle.ToString("dd MMMM yyyy", [System.Globalization.CultureInfo]::GetCultureInfo("fr-FR"))

#Définit le nom du fichier Excel prochainement enregistré
$nomRapport = "Rapport $dateFrancaise"
# Définit le chemin du dossier
$nomDossierRapport = "$cheminDossier\$nomRapport"

$variableDoublon = 1

while(Test-Path -Path $nomDossierRapport) {

    # Définit un nouveau chemin dans le cas où le chemin est déjà existant
    $nomDossierRapport = "$cheminDossier\$nomRapport\$variableDoublon"
    # Augmente la variable
    $variableDoublon++

}

# Créer un dossier de rapport
$dossierRapport = New-Item -Path $nomDossierRapport -ItemType Directory -Force

#Définit le chemin du fichier prochainement enregistré
$cheminSauvegarde = Join-Path -Path $dossierRapport -ChildPath $nomRapport

# Enregistre le fichier Excel dans le dossier
$FeuilleRapport.SaveAs($cheminSauvegarde)

#Ferme Excel
$excel.Quit()
