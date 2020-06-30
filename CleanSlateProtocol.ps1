<#
	.SYNOPSIS 
	Jarvis: All wrapped up here, sir. Will there be anything else?
	Tony Stark: You know what to do.
	Jarvis: The Clean Slate Protocol, sir?
	Tony Stark: Screw it, it's Christmas! Yes, yes!

	.PARAMETER Path
	Path to directory to be protected.

	.PARAMETER UserName
	UserName to be monitored for activity.

	.PARAMETER Days
	Number of days of inactivity required.

	.PARAMETER Testing
	Do you want this to be run in Testing mode? Default is "YES".

	.DESCRIPTION
	Usage: .\CleanSlateProtocol.ps1 -Path E:\Some\Directory -UserName UserName -Days 14 -Testing NO
#>
#################################################
<#
Looks for when a specific user last logged in then deletes the contents of the specified directory if the
designated number of days has passed. If you provide a valid address to send to, warnings will be sent
starting 4 days before the event. All Domain Controllers are queried to get the most recent login date.

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    https://github.com/dlindridge/UsefulPowershell
    Created: October 8, 2019
    Modified: October 22, 2019
#>
#################################################

	Param (
        [Parameter(Mandatory=$True)]
		$Path,
		[Parameter(Mandatory=$True)]
		$UserName,
		[Parameter(Mandatory=$True)]
		$Days,
		$Testing = "YES"
	)

### Script Settings
	$emailAlerts = "YES" 
	$fromAddr = "MyCompany IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
	$toAddr = "itadmins@MyDomain.TLD" # Enter the TO address for the e-mail alert - Must be inside quotes.
	$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.
	$inactiveDays = $Days

#################################################

### Input Validation ############################
	$emailAlerts = $emailAlerts.ToUpper()
	$Testing = $Testing.ToUpper()


### Script Actions ##############################
	Import-Module ActiveDirectory
	$inactivityTime = (Get-Date).AddDays(-$inactiveDays)
	# Check all DCs for most recent login time
	$DCs = Get-ADDomainController -Filter * | ForEach { Get-ADUser -Identity $UserName -Properties LastLogon -Server $_.Hostname | Select-Object SamAccountName,LastLogon }
	$logonData = ForEach ($UserEntry in $DCs | Group-Object SamAccountName){ $UserEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }

	# Convert logon entry to human readable date format
	ForEach ($dc in $DCs) { 
		$lngexpires = $logonData.LastLogon
		If (-not ($lngexpires)) { $lngexpires = 0 }
		If (($lngexpires -eq 0) -or ($lngexpires -gt [DateTime]::MaxValue.Ticks)) {
			$LastLogon = "<Never>"
		}
		Else {
			$Date = [DateTime]$lngexpires
			$LastLogon = $Date.AddYears(1600).ToLocalTime()
		}
	}

	Write-Host -ForegroundColor Yellow $UserName "last logged on at:" $LastLogon

	# Determine Expiration 
    $expiresOn = $LastLogon.AddDays($inactiveDays)
    $today = (Get-Date)
    $daysToExpire = (New-TimeSpan -Start $today -End $expiresOn).Days


	if ($daysToExpire -le 5) {
		$body = @("
		<p style='font-family:calibri'>Your account has not recorded a logon to a DC since $LastLogon.<br />
		The Safety Protocol will activate if there have not been any logons of this account in $inactiveDays days.<br />
		You have $daysToExpire to take action. You know what to do.</p>
		<p style='font-family:calibri'>Last Logon: $LastLogon<br />
		Expiration: $expiresOn<br />
		Days Remaining: $daysToExpire</p>")
	}

	if ($daysToExpire -lt 0) {
		$body = @("
		<p style='font-family:calibri'>Your account has not recorded a logon to a DC since $LastLogon.<br />
		The Safety Protocol has been activated and the matter is now closed.</p>
		<p style='font-family:calibri'>Last Logon: $LastLogon<br />
		Expiration: $expiresOn<br />
		Days Remaining: $daysToExpire</p>")
	}

 	# Delete files and empty folders
	if ($Testing -eq "NO" -AND $daysToExpire -lt 0) {
		Get-ChildItem -Path $Path -Recurse -Force | Where-Object { !$_.PSIsContainer } | Remove-Item -Force
		Get-ChildItem -Path $Path -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse
	}

	if ($emailAlerts -eq "YES" -AND $daysToExpire -le 4 -AND $daysToExpire -gt -1) {
		Send-MailMessage -To $toAddr -From $fromAddr -Subject "Safety Alert!" -Body "$body" -SmtpServer $smtpServer -BodyAsHtml
	}

	if ($Testing -eq "YES") {
		Send-MailMessage -To $toAddr -From $fromAddr -Subject "Safety Test" -Body "This is only a test. Carry on. $daysToExpire - $Path" -SmtpServer $smtpServer -BodyAsHtml
	}
