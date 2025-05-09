# ConvertDesjardinsToMonarch

**Fichiers**
ConvertDesjardinsToMonarch.ps1 : Script faisant la conversion des données du format CSV de compte Desjardins vers le format CSV de compte Monarch
monarc_mapping_database.csv : Base de données contenant les correspondances entre une paire "Type / Fournisseur" et une catégorie Monarch

**Conventions**
Les transactions d'un compte Desjardins sont généralement désignées en 2 parties dans l'état de compte exporté.
  
Exemple:
Paiement facture - AccèsD Internet /HYDRO-QUEBEC
Dépôt direct /DESJARDINS SECUR. FIN.

Dans ce script, on désigne la portion à gauche du caractère '/' comme étant le type
et la portion à droite du caractère '/' comme étant le fournisseur.
Ces champs correspondent à ce que vous trouverez dans le fichier monarc_mapping_database.csv.

**Instructions**
1.   Mettre les fichiers ConvertDesjardinsToMonarch.ps1 et monarc_mapping_database.csv dans le même répertoire.
2.   Editer le fichier monarc_mapping_database.csv et populer les lignes en fonction de la catégorisation désirée dans Monarch.
3.   Editer le fichier ConvertDesjardinsToMonarch.ps1 et modifier les variables de la section ### Variables principales vers le haut.
4.   Executer le script en spécifiant le chemin du fichier à importer:
     Example:  .\ConvertDesjardinsToMonarch.ps1 "C:\Desjardins\releve1.csv"
5.   Le script vous demandera le numéro de la ligne à partir de laquelle vous souhaitez faire la conversion
6.   Après l'exécution, deux fichiers seront créés:
     Un fichier "Monarch Transactions" à importer dans le compte de destination avec le bouton Edit > Upload transactions.
     Un fichier "Monarch Balance" à importer dans le compte de destination avec le bouton Edit > Upload balance history
   
