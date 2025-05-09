
<# 
.SYNOPSIS
    Ce script sert à convertir un fichier CSV provenant d'une exportation de données de compte Desjardins vers un fichier lisible par Monarch Money.

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

# Charger mapping categories
$database = import-csv .\monarch_mapping_database.csv -encoding utf8

# Importer fichier dans une variable
$source = Import-Csv $InFile -Encoding ansi -Header 'Caisse','NumCompte','TypeCompte','Date','NumLigne','Description','Vide1','Retrait','Depot','Vide2','Vide3','Vide4','Vide5','Solde'

# Créer nouvelle variable nettoyée
$clean = @()

Foreach ($s in $source) {
	$c_NumLigne = [int]($s.NumLigne).TrimStart('0')
	$c_Date = $s.date -replace "/", "-"
	$c_Merchant = $s.description -split "/" | Select -Index 1
	
	#Trouver catégorie
	$find = $null
	$type = ($s.description -split "/" | Select -Index 0).TrimEnd()
	$supplier = ($s.description -split "/" | Select -Index 1).TrimEnd()
	$find = $database | where {$_.Type -eq $type -and $_.Fournisseur -eq $supplier}
	If ($find) { $c_Category = $find.Categorie }
	Else { $c_Category = "Uncategorized" }
	
	$c_Account = "Nom du compte dans Monarch" # Changer pour le nom du compte correspondant dans Monarch
	$c_OriginalStatement = $s.Description

	#Définir montant
	If ($s.Retrait) { $c_amount = $("-" + $s.Retrait) }
	If ($s.Depot) { $c_amount = $("+" + $s.Depot) }
	
	
	$c_Tags = $null # Peut-être changé si désiré. Remplacer $null par le nom du tag entres guillemets anglais
	
	$c_Solde = $s.Solde
	
	$clean += @(
    [PSCustomObject]@{ NumLigne=$c_NumLigne; Date=$c_Date; Merchant=$c_Merchant; Category=$c_Category; Account=$c_Account; "Original Statement"=$c_OriginalStatement; Notes=$null; Amount=$c_amount; Tags=$c_Tags; Balance=$c_Solde }	
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
