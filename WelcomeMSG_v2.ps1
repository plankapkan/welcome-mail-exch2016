# New mailbox setup script and welcome mail sender

$mailbox_database = "Mailbox_D*" # database name must be running on server
$mailserver = "mail.company.ru"
$path_to_html = "C:\Scripts\welcome_mail\welcome_msg.html"
$message_from = "welcome@company.ru"
$message_bcc = "hr_need_to_know@company.ru"
$message_subject = "Welcome to Company!"

# Check if server has active database
if ((Get-MailboxDatabaseCopyStatus | select Name,Status | ?{$_.Name -like $mailbox_database -and $_.Status -like "Mounted"}) -ne $Null)
{
	Write-host "Running Script"
	# Recieve username
	# Here we wait 20 minutes till mailbox will sync between all Domain Controllers!
	$UsrIdentitys = (Get-Mailbox -ResultSize unlimited | ?{$_.WhenMailboxCreated -gt (date).AddMinutes(-85) -and $_.WhenMailboxCreated -lt (date).AddMinutes(-20)}).alias
	Foreach ($UsrIdentity in $UsrIdentitys)
	{
		# Disable POP3		
        if ((Get-CASMailbox -Identity $UsrIdentity).POPEnabled -eq $True){
			Set-CASMailbox -Identity $UsrIdentity -POPEnabled $False
		}
        # Enable SingleItemRecovery
		if ((Get-Mailbox -Identity $UsrIdentity).SingleItemRecoveryEnabled -eq $False){
			Set-Mailbox $UsrIdentity -SingleItemRecoveryEnabled $true
		}
        # Set AuditEnabled
        if ((Get-Mailbox -Identity $UsrIdentity).AuditEnabled -eq $False){
			Set-Mailbox $UsrIdentity -AuditEnabled $true
		}
        # Set WorkTime
		if ((Get-MailboxCalendarConfiguration -Identity $UsrIdentity).WorkingHoursTimeZone -ne "Russian Standard Time"){
			Set-MailboxCalendarConfiguration $UsrIdentity -WorkingHoursStartTime '10:00' -WorkingHoursEndTime '19:00' -WorkingHoursTimeZone “Russian Standard Time”
		}
        # Setup OWA
		if ((Get-MailboxRegionalConfiguration -Identity $UsrIdentity).DateFormat -eq $Null -or (Get-MailboxRegionalConfiguration -Identity $UsrIdentity).Language -eq $Null){
			Set-MailboxRegionalConfiguration -Identity $UsrIdentity -TimeZone "Russian Standard Time" -DateFormat "dd.MM.yyyy" -Language "ru-RU" -LocalizeDefaultFolderName
		} 
		# Search in mailbox if welcome mail already sent
		while ((Search-Mailbox -Identity $UsrIdentity -EstimateResultOnly -SearchQuery {subject:$message_subject}).ResultItemsCount -eq 0)
		{
			Write-host "Sending mail to" $UsrIdentity
			#Welcome mail main code
			$mbx = (Get-Mailbox -Identity $UsrIdentity -ErrorAction SilentlyContinue)
			$usr = Get-User $mbx.Identity
			# Create objects
			$message = New-Object System.Net.Mail.MailMessage
			$smtpClient = New-Object System.Net.Mail.SmtpClient($mailserver) # Change this to point to your SMTP server
			# Get text
			$messageText = [string](Get-Content ($path_to_html)) # Change to point to your HTML welcome message
			$messageText = $messageText.Replace("#NewUser00#", $usr.FirstName) # This replaces #NewUser00# with the user's first name - you can add further replacements as needed
			# Create HTML
			$view = [System.Net.Mail.AlternateView]::CreateAlternateViewFromString($messageText, $null, "text/html")
			# Add attachments
			#$image = New-Object System.Net.Mail.LinkedResource("C:\WelcomeMessage\image001.png")
			#$image.ContentId = "image001.png"
			#$image.ContentType = "image/png"
			#$view.LinkedResources.Add($image)	 
			# message
			$message.From = $message_from # address
			$message.To.Add($mbx.PrimarySmtpAddress.ToString())
			$message.Bcc.Add($message_bcc)
			$message.Subject = $message_subject # subject
			$message.AlternateViews.Add($view)
			$message.IsBodyHtml = $true
			# Send message
			$smtpClient.Send($message)
			Start-Sleep -s 120
		}
	# Clean variables
	#Remove-Variable image
	#Remove-Variable view
	if ($messageText) {Remove-Variable messageText}
	if ($message) {Remove-Variable message}
	if ($smtpClient) {Remove-Variable smtpClient}
	if ($usr) {Remove-Variable usr}
	if ($mbx) {Remove-Variable mbx}
	if ($UsrIdentity) {Remove-Variable UsrIdentity}
	}
}