# Création des comptes Office 365 à aparti d'un CSV.

## Struture du CSV

* Le CSV doit avoit comme séparateur des ","
* Le nom des colonne doit être : Email, prenom, Nom
* Les nom des étuidants peuvent avoir des caractères spéciaux ou accent, le script se chargera de faire le ménnage.

  ** Note ** : Le titre "prenom" de la colonne prénom du csv ne oit pas contenir d'accent.

  Exemple :
  
Email,prenom,nom
test@gmail.com, PrénomEtudiant, Nom-étudiant

## Script 

Ce script PowerShell a été développé pour fonctionner sur des machines Linux avec powershell 7.4
Il installe les modules powershell suivant à sa première execution : 
- Microsoft.Graph -Scope CurrentUser
- MSOnline

  ## Paramétrage :
Il est nécessaire de modifier certain paramétrage avant la premère utilisation :
* Sur la ligne 11 : Connect-ExchangeOnline -UserPrincipalName "*******@h3campus.fr"  -> Choissir une adresse ayant les droits de création de compte.
* Sur les ligne 19 et 20 : Remplacer l'adresse qui enverra les emails : $smtpUser = "Compte_EnvoieMail@h3hitema.fr" # L'adresse email de l'expéditeur
$smtpPassword = "********"
* Ligne 28 : Verifier le chemin et le nom du fichier contenant les étudiants.

  ## Evolution :
  Afin de ne pas se retouver avec trop de doublons ou adresses inutile, le script a été verrouillé pour ne pas créer d'adresse pour des étudiants dont les noms sont existant.
  Dans une prochaine version une option sera disponible.
