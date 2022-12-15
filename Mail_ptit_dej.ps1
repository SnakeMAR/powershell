# Vérifie si Outlook est ouvert
if (Get-Process outlook -ErrorAction SilentlyContinue)
{
    # Fermez Outlook
    Stop-Process -Name outlook

    # Attendez 15 secondes
    Start-Sleep -Seconds 2

    # Ouvrez Outlook
    Start-Process outlook

}
$outlook = New-Object -com Outlook.Application

$mail = $outlook.CreateItem(0)

# Récupérer la variable contenant les destinataires
$recipients = "adrien.marasco@outlook.it"

# Récupérer la date courante
$currentDate = Get-Date

# Déterminer la date de l'invitation en fonction du jour de la semaine
if ($currentDate.DayOfWeek -eq "Monday") {
  $invitationDate = $currentDate.AddDays(1)
} elseif ($currentDate.DayOfWeek -eq "Tuesday") {
  $invitationDate = $currentDate.AddDays(2)
} elseif ($currentDate.DayOfWeek -eq "Thursday") {
  $invitationDate = $currentDate.AddDays(4)
} else {
  # Par défaut, si nous ne sommes ni lundi, ni mardi, ni jeudi, envoyer l'invitation pour le lundi suivant
  $invitationDate = $currentDate.AddDays(7)
}

# Créer un nouveau message avec un corps au format HTML
$mail = $outlook.CreateItem(0)
$mail.BodyFormat = 2
$mail.HTMLBody = "<html><body>Bonjour à tous,<br><br>Je vous écris pour vous dire que j'ai inséré une clé USB suspecte dans mon PC MAN, outil professionnel. Etant donné que cette action n'est pas très pro, je me vois contraint de vous l'avouer et de vous inviter à un petit déjeuner, vous recevrez une invitation à part.<br><br>A bientôt.<br></body></html>"

# Ajouter les destinataires et un objet au message
$mail.To = $recipients
$mail.Subject = "Invitation à un petit déjeuner GRATUIT"

# Définir l'importance du message comme "Haute"
$mail.Importance = 2

# Envoyer le message
$mail.Send()

# Supprimer définitivement le mail envoyé
$mail.Delete()

# Créer une invitation au petit déjeuner digital pour 9h
$appointment = $outlook.CreateItem(1)
$appointment.Start = $invitationDate.ToString("MM/dd/yyyy") + " 9:00 AM"
$appointment.End = $invitationDate.ToString("MM/dd/yyyy") + " 10:00 AM"
$appointment.Subject = "Petit déjeuner GRATUIT"
$appointment.Location = "Au bureau"
$appointment.Body = "Rejoignez-nous pour un petit déjeuner c'est moi qui régale. Nous en profiterons pour discuter de la sécurité informatique et du verrouillage des sessions Windows."

# Ajouter les destinataires à l'invitation
$appointment.Recipients.Add($recipients)

# Envoyer l'invitation
$appointment.Send()