<#
	.SYNOPSIS 
	Dean Thomas: Hey, look! Neville's got a Rememberall!
	Hermione Granger: I've read about those! When the smoke turns red, it means you've forgotten something.
	Neville Longbottom: Only problem is, I can't remember what I've forgotten.
#>

#################################################
<#
This script will send reminder emails to your designated destination (such as your Security or IT teams) to remind them
that it is time to send a particular email to perform a task. The tasks that you will be reminded of are specifically
built with the intention of complying with PCI-DSS and/or SOC2 certifications.

Set to run on any server in your environment authorized to use your SMTP relay as SYSTEM once per week. Monday or 
Tuesday recommended.

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: May 22, 2020
Modified: June 1, 2020
#>

### Email Settings ##############################
$fromAddr = "Company IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
$toAddr = "itadmins@MyDomain.TLD" # Enter the TO address for the e-mail alert - Must be inside quotes.
$testingAddr = "testingAddress@MyDomain.TLD" # Email address to send reports and alerts to when testing.
$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.


### Options #####################################
$Testing = "YES" #Testing mode to avoid spamming your group
$AnnualPolicyAcknowledgement = "YES" #Reminder to send annual policy review confirmation
$FirewallReview = "YES" #Reminder to review firewall rules
	$FRFrequency = "EVEN" #How often - Yearly, Even (even numbered quarters), Odd (odd numbered quarters), Quarterly
$AwarenessEmails = "YES" #Reminder to send awareness/training emails
	$AEFrequency = "6" #How often - Number of weeks between sending
$AccessReview = "YES" #Reminder to perform access reviews
	$ACFrequency = "EVEN" #How often - Yearly, Even (even numbered quarters), Odd (odd numbered quarters), Quarterly
$OffWeeks = "YES" #Send an update email when nothing is due?


### Validations #################################
$Testing = $Testing.ToUpper()
$AnnualPolicyAcknowledgement = $AnnualPolicyAcknowledgement.ToUpper()
$FirewallReview = $FirewallReview.ToUpper()
$FRFrequency = $FRFrequency.ToUpper()
$AwarenessEmails = $AwarenessEmails.ToUpper()
$AccessReview = $AccessReview.ToUpper()
$ACFrequency = $ACFrequency.ToUpper()
$OffWeeks = $OffWeeks.ToUpper()


### Necessary Transforms ########################
$actionDate = Get-Date -Format "dddd, d MMMM, yyyy"
$reportDate = Get-Date -DisplayHint Date
$actionCount = 0
$emailBody = @("These are your recurring reminders:<br /><br />")


### Calculate Week Number #######################
Function Get-WeekNumber([datetime]$DateTime = (Get-Date)) {
    $cultureInfo = [System.Globalization.CultureInfo]::CurrentCulture
    $cultureInfo.Calendar.GetWeekOfYear($DateTime,$cultureInfo.DateTimeFormat.CalendarWeekRule,$cultureInfo.DateTimeFormat.FirstDayOfWeek)
}
$weekNumber = Get-WeekNumber


### Annual Policy Acknowledgement ###############
If ($AnnualPolicyAcknowledgement -eq "YES") {
	If ($weekNumber -eq 2) {
		$emailBody += "<b>Annual Policy Acknowledgement:</b> Send reminders to all employees to review policy documents including Acceptable Use, Security, and any other appropriate company/information management policies.<br /><br />"
		$actionCount += 1
	}
}


### Firewall Review #############################
If ($FirewallReview -eq "YES") {
	If ($FRFrequency -eq "YEARLY") {
		If ($weekNumber -eq 2) {
			$emailBody += "<b>Firewall Review:</b> Schedule the yearly firewall rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($FRFrequency -eq "EVEN") {
		If (($weekNumber -eq 16) -OR ($weekNumber -eq 42)) {
			$emailBody += "<b>Firewall Review:</b> Schedule the bi-yearly firewall rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($FRFrequency -eq "ODD") {
		If (($weekNumber -eq 2) -OR ($weekNumber -eq 29)) {
			$emailBody += "<b>Firewall Review:</b> Schedule the bi-yearly firewall rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($FRFrequency -eq "QUARTERLY") {
		If (($weekNumber -eq 2) -OR ($weekNumber -eq 16) -OR ($weekNumber -eq 29) -OR ($weekNumber -eq 42)) {
			$emailBody += "<b>Firewall Review:</b> Schedule the quarterly firewall rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
}


### Awareness Emails ############################
If ($AwarenessEmails -eq "YES") {
	If ($weekNumber % $AEFrequency -eq 0) {
		$emailBody += "<b>Awareness/Training Emails:</b> Send awareness or training emails to all employees on a subject determined by your organization.<br /><br />"
		$actionCount += 1
	}
}


### Access Review ###############################
If ($AccessReview -eq "YES") {
	If ($ACFrequency -eq "YEARLY") {
		If ($weekNumber -eq 2) {
			$emailBody += "<b>Access Review:</b> Schedule the yearly access rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($ACFrequency -eq "EVEN") {
		If (($weekNumber -eq 16) -OR ($weekNumber -eq 42)) {
			$emailBody += "<b>Access Review:</b> Schedule the bi-yearly access rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($ACFrequency -eq "ODD") {
		If (($weekNumber -eq 2) -OR ($weekNumber -eq 29)) {
			$emailBody += "<b>Access Review:</b> Schedule the bi-yearly access rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
	If ($ACFrequency -eq "QUARTERLY") {
		If (($weekNumber -eq 2) -OR ($weekNumber -eq 16) -OR ($weekNumber -eq 29) -OR ($weekNumber -eq 42)) {
			$emailBody += "<b>Access Review:</b> Schedule the quarterly access rule review meeting.<br /><br />"
			$actionCount += 1
		}
	}
}


### Send Email ##################################
If ($actionCount -ne 0) {
	If ($Testing -eq "YES") { $toAddr = $testingAddr }
	$emailBody += "<br /><br />$weekNumber"
	$emailSubject = "Action Required: $actionCount Reminders Due"
	Send-MailMessage -To $toAddr -From $fromAddr -Subject $emailSubject -Body "$emailBody" -SmtpServer $smtpServer -BodyAsHtml -Priority High
}

If (($actionCount -eq 0) -AND ($OffWeeks -eq "YES")) {
	If ($Testing -eq "YES") { $toAddr = $testingAddr }
	$emailSubject = "Reminders - Nothing Due"
	$emailBody += "<H1>NOTHING!</H1><br />You are all caught up, this is just letting you know the script is still running.<br />"
	$emailBody += "<br /><br />$weekNumber"
	Send-MailMessage -To $toAddr -From $fromAddr -Subject $emailSubject -Body "$emailBody" -SmtpServer $smtpServer -BodyAsHtml
}
