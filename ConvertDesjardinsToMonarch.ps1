
<# 
.SYNOPSIS
    Ce script sert à convertir un fichier CSV provenant d'une exportation de données de compte Desjardins vers un fichier CSV lisible par Monarch Money.

.PARAMETER [InFile]
    Fichier CSV source de Desjardins

.NOTES
    Author    : Pierre L. Vallee
    Version   : 1.0
    Last Updated : 2025-05-08

#>


param(
	[Parameter(Mandatory=$true)]
	[string]$InFile
)

### Variables principales
$compteMonarch = "Nom du compte correspondant dans Monarch"
$tagMonarch = $null # Changer $null pour "Nom du tag" si vous souhaitez ajouter un tag dans Monarch pour toutes les transactions importées
$noteMonarch = $null # Changer $null pour "Contenu de la note" si vous souhaitez ajouter une note dans Monarch pour toutes les transactions importées


# Charger mapping categories
$database = import-csv .\monarch_mapping_database.csv -encoding utf8

# Importer fichier source Desjardins dans une variable
$source = Import-Csv $InFile -Encoding ansi -Header 'Caisse','NumCompte','TypeCompte','Date','NumLigne','Description','Vide1','Retrait','Depot','Vide2','Vide3','Vide4','Vide5','Solde'

# Créer nouvelle variable nettoyée
$clean = @()

Foreach ($s in $source) {
	$c_NumLigne = [int]($s.NumLigne).TrimStart('0')
	$c_Date = $s.date -replace "/", "-"

	# Vérifier si fournisseur présent
	$description = $s.description -split "/"
	If ($description.count -gt "1") { 
		$c_Merchant = ($description | Select -Index 1).TrimEnd() 
	}
	Else { 
		$c_Merchant = $null
	}
	
	# Trouver catégorie selon type et fournisseur
	$find = $null
	$type = ($s.description -split "/" | Select -Index 0).TrimEnd()
	$find = $database | where {$_.Type -eq $type -and $_.Fournisseur -eq $c_Merchant}
	If ($find) { $c_Category = $find.Monarch_Category }
	Else { $c_Category = "Uncategorized" }
	
	$c_Account = $compteMonarch
	$c_OriginalStatement = $s.Description

	$c_Notes = $noteMonarch
	
	# Définir montant
	If ($s.Retrait) { $c_amount = $("-" + $s.Retrait) }
	If ($s.Depot) { $c_amount = $("+" + $s.Depot) }
	
	$c_Tags = $tagMonarch
	
	$c_Solde = $s.Solde
	
	$clean += @(
    [PSCustomObject]@{ NumLigne=$c_NumLigne; Date=$c_Date; Merchant=$c_Merchant; Category=$c_Category; Account=$c_Account; "Original Statement"=$c_OriginalStatement; Notes=$c_Notes; Amount=$c_amount; Tags=$c_Tags; Balance=$c_Solde }	
	)
}

# Afficher variable nettoyée et demander à partir de quelle ligne faire la conversion
Write-Host ""
Write-Host "Transactions du fichier CSV source:"
$clean | Select-Object @{Label="Numéro de ligne"; Expression={$_.NumLigne}}, Date, @{Label="Description"; Expression={$_."Original Statement"}}, @{Label="Montant"; Expression={$_."Amount"}} | Format-Table
$choixligne = Read-Host "Entrer numéro de ligne de départ pour la conversion"

# Construire CSV pour Monarch
$lignes = $clean | where {$_.numligne -ge $choixligne} | select Date,Merchant,Category,Account,"Original Statement",Notes,Amount,Tags

# Construire CSV de solde pour Monarch
$lignesBalance = $clean | where {$_.numligne -ge $choixligne} | select Date,Balance,Account
$lignesBalance = $lignesBalance | Group-Object -Property date | ForEach-Object { $_.Group | Select-Object -Last 1 }

# Définir noms des fichiers pour exportation
$OutFileMonarchTransactions = $InFile.TrimEnd(".csv")
$OutFileMonarchTransactions = $OutfileMonarchTransactions + " Monarch Transactions.csv"
$OutFileMonarchBalance = $InFile.TrimEnd(".csv")
$OutFileMonarchBalance = $OutFileMonarchBalance + " Monarch Balance.csv"

# Exporter fichiers CSV
$lignes | ConvertTo-Csv -NoTypeInformation | Out-File -Force -Encoding utf8 $OutFileMonarchTransactions
$lignesBalance | ConvertTo-Csv -NoTypeInformation | Out-File -Force -Encoding utf8 $OutFileMonarchBalance
Write-Host ""
Write-Host "***** Les transactions sélectionnées ont été converties et exportées *****"
Write-Host "Fichier de transactions: `"$OutFileMonarchTransactions`""
Write-Host "Fichier de solde: `"$OutFileMonarchBalance`""
Write-Host ""
