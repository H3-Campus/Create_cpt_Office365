#Prérequis :
Install-Module Microsoft.Graph -Scope CurrentUser
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"

#Install-Module -Name ExchangeOnlineManagement
#Import-Module -Name ExchangeOnlineManagement
Install-Module -Name MSOnline
Import-Module -Name MSOnline

######## CONNEXION ###########
Connect-ExchangeOnline -UserPrincipalName "*******@h3campus.fr"
# Domaine utilisé pour les adresses email
#$domain = "h3hitema.fr"  
$domain = "h3campus.fr"  

# Paramètres pour Office 365
$smtpServer = "smtp.office365.com"
$smtpPort = 587
$smtpUser = "Compte_EnvoieMail@h3hitema.fr" # L'adresse email de l'expéditeur
$smtpPassword = "********"

# Définir les codes de couleur
$red = "e[31m" 
$green = "e[32m"
$reset = "`e[0m"

# Charger les informations des étudiants à partir du fichier CSV
$csvPath = "students.csv"
$etudiants = Import-Csv -Path $csvPath

function Normalize-String {
param (
[string]$inputString
)

# Remplacer les caractères accentués et non valides
$normalizedString = $inputString -replace '[éèêë]', 'e' `
-replace '[àâä]', 'a' `
-replace '[ôö]', 'o' `
-replace '[ûü]', 'u' `
-replace '[ç]', 'c' `
-replace '[^a-zA-Z0-9]', '' # Supprime les caractères non alphanumériques

return $normalizedString
}

# Fonction pour vérifier si une adresse email existe déjà dans Office 365
function Adresse-Existe {
param (
[string]$email
)
try {
# Filtre correctement formatté avec la syntaxe de chaîne
$filter = "PrimarySmtpAddress -eq '{0}'" -f $email
$user = Get-Mailbox -RecipientTypeDetails UserMailbox -Filter $filter
if ($user) {
return $true
} else {
return $false
}
} catch {
return $false
}
}


# Fonction pour générer une adresse email unique
function Generer-Email {
param (
[string]$prenom,
[string]$nom,
[string]$domain = "h3campus.fr"
)

# Normaliser les chaînes et les convertir en minuscules
$prenomNormalise = (Normalize-String $prenom).ToLower()
$nomNormalise = (Normalize-String $nom).ToLower()
$domainLower = $domain.ToLower()


$index = 1
# Générer l'email en minuscules
$email = "{0}.{1}@{2}" -f $prenomNormalise.Substring(0, $index), $nomNormalise, $domainLower

# Boucle tant qu'une adresse existe
while (Adresse-Existe $email) {
$index++
$email = "{0}.{1}@{2}" -f $prenomNormalise.Substring(0, $index), $nomNormalise, $domainLower
}

return $email
}


function Generer-MotDePasse {
param (
[int]$length = 12,
[int]$specialCharsCount = 2
)

# Vérification de la longueur
if ($length -le $specialCharsCount) {
throw "La longueur du mot de passe doit être supérieure au nombre de caractères spéciaux."
}

# Définir les ensembles de caractères
$letters = 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
$numbers = '0123456789'
$specialChars = '!@#$%^&*()-+='

# Initialiser le mot de passe
$password = ''

# Ajouter des lettres
for ($i = 0; $i -lt ($length - $specialCharsCount - 1); $i++) {
$index = Get-Random -Minimum 0 -Maximum $letters.Length
$password += $letters[$index]  # Prendre une lettre en utilisant l'index
}

# Ajouter un chiffre aléatoire
$index = Get-Random -Minimum 0 -Maximum $numbers.Length
$password += $numbers[$index]  # Ajouter un chiffre en utilisant l'index

# Ajouter les caractères spéciaux
for ($i = 0; $i -lt $specialCharsCount; $i++) {
$index = Get-Random -Minimum 0 -Maximum $specialChars.Length
$password += $specialChars[$index]  # Prendre un caractère spécial en utilisant l'index
}

# Mélanger les caractères
$finalPassword = -join ($password.ToCharArray() | Sort-Object { Get-Random })

return $finalPassword
}


# Fonction pour envoyer un email avec les informations de connexion
function Envoyer-Email {
param (
[string]$destinataire,
[string]$login,
[string]$password
)

# Contenu de l'email avec mise en forme HTML
$subject = "Compte - $domain"
$body = @"
<html>
<body>
<p>Bonjour,</p>
<p>C-dessous vous trouverz vos informations de connexion pour accéder à votre compte Teams :</p>
<table style='border-collapse: collapse;'>
<tr>
<th style='border: 1px solid black; padding: 8px;'>Login</th>
<th style='border: 1px solid black; padding: 8px;'>Mot de passe</th>
</tr>
<tr>
<td style='border: 1px solid black; padding: 8px;'>$login</td>
<td style='border: 1px solid black; padding: 8px;'>$password</td>
</tr>
</table>

<p>Un changement du mot de passe vous sera demandé à la première connexion.</p>
<p>Veuillez conserver ces informations en lieu sûr.</p><br>
<p>Cordialement,<br>Service Informatique</p>
</body>
</html>
"@

# Envoyer l'email
$smtpClient = New-Object Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtpClient.EnableSsl = $true
$smtpClient.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPassword)

$mailMessage = New-Object Net.Mail.MailMessage
$mailMessage.From = $smtpUser
$mailMessage.To.Add($destinataire)
$mailMessage.Subject = $subject
$mailMessage.Body = $body
$mailMessage.IsBodyHtml = $true

try {
$smtpClient.Send($mailMessage)
Write-Host "Email envoyé avec succès à $destinataire"
} catch {
Write-Host "Échec de l'envoi de l'email à $destinataire : $"
}
}

function Creer-Utilisateur {
param (
[string]$prenom,
[string]$nom,
[string]$email,
[string]$domain = "h3campus.fr",
[string]$usageLocation = "FR"
)

$password = Generer-MotDePasse

$PasswordProfile = @{
Password = $password
ForceChangePasswordNextSignIn = $true
}

# Générer un mailNickname valide
$mailNickname = ($email -split '@')[0] -replace '[^a-zA-Z0-9]', ''
if ([string]::IsNullOrEmpty($mailNickname)) {
$mailNickname = "user" + (Get-Random -Minimum 1000 -Maximum 9999)
}

try {
# Vérifier si l'utilisateur existe déjà
$existingUser = Get-MgUser -Filter "UserPrincipalName eq '$email'" -ErrorAction SilentlyContinue

if ($existingUser) {
Write-Host "L'utilisateur $email existe déjà." -ForegroundColor Yellow
return @{
Success = $true
Email = $email
Password = $null
IsNewUser = $false
}
}

Write-Host "Création de l'utilisateur : $email ..." -ForegroundColor Yellow

$newUser = New-MgUser -DisplayName "$prenom $nom" `
-GivenName $prenom `
-Surname $nom `
-UserPrincipalName $email `
-MailNickname $mailNickname `
-PasswordProfile $PasswordProfile `
-AccountEnabled:$true `
-UsageLocation $usageLocation

if ($null -ne $newUser) {
$exchangeLicense = Get-MgSubscribedSku | Where-Object SkuPartNumber -eq 'EXCHANGESTANDARD_STUDENT'

if ($null -ne $exchangeLicense) {
Set-MgUserLicense -UserId $newUser.Id `
-AddLicenses @{SkuId = $exchangeLicense.SkuId} `
-RemoveLicenses @()

Write-Host "Licence Exchange Online for Students attribuée à $email" -ForegroundColor Green

return @{
Success = $true
Email = $email
Password = $password
IsNewUser = $true
}
} else {
Write-Host "Erreur : Licence Exchange Online for Students non trouvée" -ForegroundColor Red
}
} else {
Write-Host "Erreur : L'utilisateur $email n'a pas pu être créé." -ForegroundColor Red
}

} catch {
Write-Host "Erreur lors de la création de l'utilisateur : $_" -ForegroundColor Red
}

return @{
Success = $false
Email = $email
Password = $null
IsNewUser = $false
}
}
# Parcourir chaque étudiant dans le fichier CSV et traiter
foreach ($etudiant in $etudiants) {
	$prenom = $etudiant.Prenom
	$nom = $etudiant.Nom

	Write-Host "Traitement de l'étudiant : $prenom $nom" -ForegroundColor Cyan

	# Normaliser les chaînes et les convertir en minuscules
	$prenomuser = (Normalize-String $prenom).ToLower()
	$nomuser = (Normalize-String $nom).ToLower()
	$domainuser = $domain.ToLower()

	$Usermail = "$($prenomuser.substring(0,1)).$($nomuser)@$($domainuser)"
	$AddressExist = Adresse-Existe $Usermail

	if ($AddressExist) {
	Write-Host "Adresse mail $Usermail existante !" -ForegroundColor Blue
	} else {
	Write-Host "Création d'un nouvel utilisateur avec l'adresse $Usermail" -ForegroundColor Yellow

	# Créer l'utilisateur dans Office 365
	$result = Creer-Utilisateur -prenom $prenom -nom $nom -email $Usermail -domain $domainuser

	# Vérifier si l'adresse a bien été créée
	if ($result.Success) {
	# Envoyer un email avec les informations de connexion
	$messageEnvoi = Envoyer-Email -destinataire $etudiant.Email -login $result.Email -password $result.Password
	Write-Host "E-mail envoyé à $($etudiant.Email) pour le compte $($result.Email)" -ForegroundColor Green
	Write-Host $messageEnvoi -ForegroundColor Green
	} else {
	Write-Host "Erreur : Le compte $($result.Email) n'a pas été créé. Aucun e-mail envoyé." -ForegroundColor Red
	}
}

# Faire une pause de 10 secondes
Write-Host "Pause de 10 secondes avant le traitement du prochain étudiant..." -ForegroundColor Gray
Start-Sleep -Seconds 10
}
