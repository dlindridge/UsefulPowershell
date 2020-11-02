<#
	.SYNOPSIS
	Captain O'Hagan: I swear to God I'll pistol whip the next guy who says "Shenanigans."
	Mac: Hey Farva what's the name of that restaurant you like with all the goofy shit on the walls and the mozzarella sticks?

	.PARAMETER Testing
	Do you want this to be run in Testing mode? Default is "YES". This is a Master Switch and will enable testing on the ENABLED
	functions. While in Testing Mode reports will be generated for enabled functions but they won't actually DO anything. This
	is intentional for you to see what they would have done.

	.DESCRIPTION
	Usage: .\Shenanigans.ps1 -Testing NO
#>

#################################################
<#
This script is a bit monolithic, but it is filled with routine maintenance tasks you should be running for Active Directory
anyways. You can pick which actions you want to run by setting the correct option in the Enable Script Actions section. I've
tried to comment this as throughly as practical so it should make sense as you go through it. It was written as a consolidation
of multiple other scripts so there may be some leftover oddities, but it does work as of writing this. Searches that could be
impacted by querying a specific Domain Controller (i.e. LastLogon) compensate by searching ALL Domain Controllers in the
specified domain and taking the most recent result.

To use this script make sure that service account you run this under has permission to change objects in the OUs specified. This
script must be run from a server that has the Active Directory Module for Powershell or RSAT installed. I recommend setting it up
as a daily scheduled task - I run mine daily at 3am.

"Script Options" section should handle all of your customization needs, but pay particular attention to the email notification
in the "Domain Password Notice" section if your preferred method of changing passwords is not on the client PC or your password
requirements are different than the standard Complexity options in Active Directory.

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
https://github.com/dlindridge/UsefulPowershell
Created: September 17, 2019
Modified: November 2,2020
#>
#################################################

	Param (
		$Testing = "YES",
	)

### Script Options ##############################
$searchRoot = "OU=MyRoot,DC=MyDomain,DC=TLD" # Where to begin your recursive search - Must be in distinguished form and inside quotes.
$disableRoot = "OU=Disabled,DC=MyDomain,DC=TLD" # Where your disabled objects are kept - Must be in distinguished form and inside quotes.
$domain = "MyDomain.TLD" # Domain you are searching - Must be inside quotes.
$companyName = "MyCompany" # The name you use to identify your organization in email.
$pwdMinimum = "10" # Minimum required character length of passwords.
$serviceDesk = "https://servicedesk.MyDomain.TLD" # The URL or email address of your ticketing system.
$expireInDays = 15 # Number of days before password expiration to start warning your users. (e.g. 15)
$timeSinceCreation = 30 # Integer for number of "grace" days since the account was created (to prevent disabling of brand new accounts)
$warningDays = 60 # Integer for number of days of object inactivity for reminder. (e.g. 60)
$inactiveDays = 90 # Integer for number of days of object inactivity (e.g. 90)
$deleteDays = 200 # Integer for number of days of inactivity before object deletion (e.g. 200)
$emailPattern = "*MyDomain.TLD" # UPN domain for your email addresses
$termDesignator = "TERM" # Text used in object description to mark it as a terminated users - Must be inside quotes.
$disableExempt = "DISABLE-EXEMPT" # Text used in object description to exempt it from being disabled by this script - Must be inside quotes.
$deleteExempt = "DO-NOT-DELETE" # Text used in object description to exempt it from being deleted by this script - Must be inside quotes.
$Path = "C:\Scripts\ADMaintenance" # Path to output for consolidation and emailing.
$fromAddr = "MyCompany IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
$toAddr = "itadmins@MyDomain.TLD" # Enter the TO address for the e-mail alert - Must be inside quotes.
$TestAddress = "someone@MyDomain.TLD" # Enter the email address test emails will be sent to.
$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.

### Enable Script Actions (YES or NO) ##############
	$emailReports = "YES" # Enable emailing reports after running.
	$enableQuotes = "YES" # Add quotes to outbound alert emails - these will go to your users too.
	$enablePassNeverExpire = "YES" # Create list of accounts who's passwords never expire.
	$enablePrivilegedAccounts = "YES" # Create list of accounts that are either Domain or Enterprise Admins.
	$enableExpiredUsers = "YES" # Archive User Objects that have met or passed their Expiration Date.
	$enableExpiringComps = "YES" # Notify users that their computer has not connected in too long.
	$enableInactiveComps = "YES" # Archive Computer Objects that have not connected to the domain in too long.
	$enableInactiveUsers = "YES" # Archive User Objects that have not logged in to the domain in too long.
	$enableTempUserWarnings = "YES" # Send email notice to users and their manager when their accounts near expiration date.
	$enableTermUsers = "YES" # Archive User Objects for terminated users after the specified time.
	$enableDisabledComps = "YES" # Archive Compouter Objects that are disabled but still in the production OUs.
	$enableDisabledUsers = "YES" # Archive User Objects that are disabled but not marked as Terminated.
	$enableDeleteComps = "NO" # Delete Computer Objects that are disabled and have not been on network in a very long time.
	$enableDeleteUsers = "NO" # Delete User Objects that are disabled and have not logged in to the domain in a very long time.
	$enableDomainPasswords = "YES" # Notify users starting 15 days before their domain password expires.
	$enableADHealthChecks = "YES" # Check AD Replication and Diagnostics.
		$adHealthDayOfWeek = "MONDAY" # Day to run the replication reports. "ALL" will run these reports every day.
	$enableDFSExport = "YES" # Make a report of all current DFS locations.

#################################################

### Cleanup Old Reports #########################
	Clear-Host
	Write-Host -ForegroundColor Yellow "Cleaning up old records..."
	Get-ChildItem -Path $Path -Filter *.csv | Remove-Item -Force | Wait-Job
	Get-ChildItem -Path $Path -Filter *.txt | Remove-Item -Force | Wait-Job


### Input Validation ############################
	$emailReports = $emailReports.ToUpper()
	$Testing = $Testing.ToUpper()
	$enableQuotes = $enableQuotes.ToUpper()
	$enablePassNeverExpire = $enablePassNeverExpire.ToUpper()
	$enablePrivilegedAccounts = $enablePrivilegedAccounts.ToUpper()
	$enableExpiredUsers = $enableExpiredUsers.ToUpper()
	$enableExpiringComps = $enableExpiringComps.ToUpper()
	$enableInactiveComps = $enableInactiveComps.ToUpper()
	$enableInactiveUsers = $enableInactiveUsers.ToUpper()
	$enableTempUserWarnings = $enableTempUserWarnings.ToUpper()
	$enableTermUsers = $enableTermUsers.ToUpper()
	$enableDisabledComps = $enableDisabledComps.ToUpper()
	$enableDisabledUsers = $enableDisabledUsers.ToUpper()
	$enableDeleteComps = $enableDeleteComps.ToUpper()
	$enableDeleteUsers = $enableDeleteUsers.ToUpper()
	$enableDomainPasswords = $enableDomainPasswords.ToUpper()
	$enableADHealthChecks = $enableADHealthChecks.ToUpper()
	$enableDFSExport = $enableDFSExport.ToUpper()
	$adHealthDay = (Get-Date).DayOfWeek
	$adHealthDay = $adHealthDay.ToString()
	$adHealthDay = $adHealthDay.ToUpper()
	$adHealthDayOfWeek = $adHealthDayOfWeek.ToUpper()


### Necessary Queries and Filters ###############
	Import-Module ActiveDirectory
	$actionDate = Get-Date -Format "dddd, d MMMM, yyyy"
	$csvDate = Get-Date -Format "yyyy-MM-dd"
	$reportDate = Get-Date -DisplayHint Date
	$createdDate = (Get-Date).AddDays(-$timeSinceCreation)
	$deleteTime = (Get-Date).AddDays(-$deleteDays)
	$body = @("These tasks were run today, $actionDate. For full details please check out the attached CSV files.<br /><br />")
	If ($Testing -eq "YES" -AND ($TestAddress -eq $Null -OR $TestAddress -eq "")) {
		Write-Host -ForegroundColor Red "Need address to send test reports to: " -NoNewLine
		$TestAddress = Read-Host
	}
	if ($Testing -eq "YES") { $toAddr = $TestAddress }


### Pick A Quote ################################
<#	Most of the formatting is done as the quote is inserted into the message. You only need to add HTML format tags where you need parts 
	customized. HTML formatting tags include <i>italics</i> <b>bold</b> and <u>underline</u>.	#>
	Function Get-Quote {
		$QuoteList = @(
		"Before you judge a man, walk a mile in his shoes. After that who cares?... He's a mile away and you've got his shoes! -Billy Connolly",
		"All the things I really like to do are either immoral, illegal or fattening. -Alexander Woollcott",
		"I don't believe in astrology; I'm a Sagittarius and we're skeptical. -Arthur C. Clarke",
		"Wine is constant proof that God loves us and loves to see us happy. -Benjamin Franklin",
		"The surest sign that intelligent life exists elsewhere in the universe is that it has never tried to contact us. -Bill Watterson",
		"My favorite machine at the gym is the vending machine. -Caroline Rhea",
		"A study in the Washington Post says that women have better verbal skills than men. I just want to say to the authors of that study: 'Duh.' -Conan O'Brien",
		"It is a scientific fact that your body will not absorb cholesterol if you take it from another person's plate. -Dave Barry",
		"I used to jog but the ice cubes kept falling out of my glass. -David Lee Roth",
		"Human beings, who are almost unique in having the ability to learn from the experience of others, are also remarkable for their apparent disinclination to do so. -Douglas Adams",
		"There is a theory which states that if ever anyone discovers exactly what the Universe is for and why it is here, it will instantly disappear and be replaced by something even more bizarre and inexplicable.There is another theory which states that this has already happened. -Douglas Adams",
		"Have you ever noticed that anybody driving slower than you is an idiot, and anyone going faster than you is a maniac? -George Carlin",
		"I refuse to join any club that would have me as a member. -Groucho Marx",
		"The two most common elements in the universe are hydrogen and stupidity. -Harlan Ellison",
		"I've got all the money I'll ever need, if I die by four o'clock. -Henny Youngman",
		"You tried your best and you failed miserably. The lesson is 'never try.' -Homer Simpson",
		"True terror is to wake up one morning and discover that your high school class is running the country. -Kurt Vonnegut",
		"Of all the things I've lost I miss my mind the most. -Ozzy Osbourne",
		"To err is human, but to really foul things up you need a computer. -Paul R. Ehrlich",
		"The man who smiles when things go wrong has thought of someone to blame it on. -Robert Bloch",
		"I found there was only one way to look thin: hang out with fat people. -Rodney Dangerfield",
		"I looked up my family tree and found out I was the sap. -Rodney Dangerfield",
		"I wish I were dumber so I could be more certain about my opinions. It looks fun. -Scott Adams",
		"A clear conscience is usually the sign of a bad memory. -Steven Wright",
		"Every man is guilty of all the good he did not do. -Voltaire",
		"When I die, I want to die like my grandfather who died peacefully in his sleep. Not screaming like all the passengers in his car. -Will Rogers",
		"It's like deja vu all over again. -Yogi Berra",
		"You better cut the pizza in four pieces because I'm not hungry enough to eat six. -Yogi Berra",
		"I'm about to do to you what Limp Bizkit did to music in the late '90s. -Deadpool <i>Deadpool</i>",
		"Leave the gun. Take the cannoli. -Peter Clemenza <i>The Godfather</i>",
		"Gentlemen, you can't fight in here. This is the war room. -President Merkin Muffley <i>Dr. Strangelove</i>",
		"Your mother was a hamster and your father smelt of elderberries. -The Insulting Frenchman <i>Monty Python and the Holy Grail</i>",
		"I am your father's brother's nephew's cousin's former roommate. -Lord Dark Helmet <i>Spaceballs</i>",
		"Human sacrifice! Dogs and cats living together! Mass Hysteria! -Dr. Peter Venkman <i>Ghostbusters</i>",
		"I swear to God I'll pistol whip the next guy who says `Shenanigans`. -Captain O'Hagan <i>Super Troopers</i>",
		"Hello. My name is Inigo Montoya. You killed my father. Prepare to die. -Ingio Montoya <i>The Princess Bride</i>",
		"Get to the chopper! -Dutch <i>Predator</i>",
		"Ladies and gentlemen, this is your stewardess speaking... We regret any inconvenience the sudden cabin movement might have caused, this is due to periodic air pockets we encountered, there's no reason to become alarmed, and we hope you enjoy the rest of your flight... By the way, is there anyone on board who knows how to fly a plane? -Elaine Dickinson <i>Airplane!</i>",
		"There's only two things I hate in this world: people who are intolerant of other people's cultures and the Dutch. -Nigel Powers <i>Goldmember</i>",
		"I'm <i>a</i> god. I'm not <i>the</i> God... I don't think. -Phil Connors <i>Groundhog Day</i>",
		"Ned, I would love to stay here and talk with you... but I'm not going to. -Phil Connors <i>Groundhog Day</i>",
		"Do what I do. Hold tight and pretend it's a plan! -The Doctor <i>Doctor Who, Season 7, Christmas Special</i>",
		"Never ignore coincidence. Unless, of course, you're busy. In which case, always ignore coincidence. -The Doctor <i>Doctor Who, Season 5, Episode 12</i>",
		"I'm not bad. I'm just drawn that way. -Jessica Rabbit <i>Who Framed Roger Rabbit</i>",
		"Looks like I picked the wrong week to quit sniffing glue. -Steve McCroskey <i>Airplane!</i>",
		"Why don't you make like a tree, and get out of here? -Biff Tannen <i>Back to the Future</i>",
		"Fat, drunk, and stupid is no way to go through life. -Dean Wormer <i>Animal House</i>",
		"Listen, strange women lyin' in ponds distributin' swords is no basis for a system of government. Supreme executive power derives from a mandate from the masses, not from some farcical aquatic ceremony. -Dennis <i>Monty Python and the Holy Grail</i>",
		"If we hit that bullseye, the rest of the dominoes should fall like a house of cards. Checkmate. -Zapp Brannigan",
		"Don't be such a chicken, Kif. Teenagers smoke, and they seem pretty on-the-ball. -Zapp Brannigan",
		"I got your distress call and came here as soon as I wanted to. -Zapp Brannigan",
		"In the game of Chess, you must never let your adversary see your pieces. -Zapp Brannigan",
		"When I'm in command, every mission is a suicide mission. -Zapp Brannigan",
		"The snozzberries taste like snozzberries! -College Boy 3 <i>Super Troopers</i>",
		"Do I look like a cat to you boy? Am I jumpin' around all nimbly bimbly from tree to tree? Am I drinking milk from a saucer? DO YOU SEE ME EATING MICE? -Foster <i>Super Troopers</i>",
		"Look, Your Worshipfulness, let's get one thing straight. I take orders from just one person: me. -Han Solo <i>Star Wars: Episode IV - A New Hope</i>"
		)
		$QuoteList | Get-Random
	}


### Password Never Expires ######################
<#	Special Note: By design this query looks at the entire domain, not just the search root.
	We want to know about EVERY account that doesn't have password expiration.	#>
	if ($enablePassNeverExpire -eq "YES") {
		Write-Host -ForegroundColor Green "Looking up passwords that don't expire..." -NoNewLine
		# Create the CSV for the report
		$passNoExpireDestination = Join-Path -Path $Path -ChildPath "PasswordNoExpire-$csvDate.csv"
		# Get user accounts whose passwords are set to never expire
		$passNeverExpire = (Get-ADUser -Server $domain -Filter * -Properties Name,PasswordNeverExpires,PasswordLastSet,whenCreated | Where-Object { $_.passwordNeverExpires -eq "true" } | Where-Object {$_.enabled -eq "true"})
		# Add information generated to the CSV
		if ($passNeverExpire -ne $Null) {
			Add-Content $passNoExpireDestination "Name,SamAccountName,WhenCreated" | Wait-Job
			ForEach ($user in $passNeverExpire) {
				$passExpName = $user.Name
				$passSamName = $user.SamAccountName
				$passCreated = $user.whenCreated
				Add-Content $passNoExpireDestination "$passExpName,$passSamName,$passCreated" | Wait-Job
			}
		}
		$passNeverExpireCount = $passNeverExpire | Measure-Object
		Write-Host -ForegroundColor Red $passNeverExpireCount.Count
		$body += "Accounts With Password Never Expires: <strong>$($passNeverExpireCount.Count)</strong><br />"
	}


### Privileged Admins ###########################
	if ($enablePrivilegedAccounts -eq "YES") {
		Write-Host -ForegroundColor Green "Looking up privileged accounts..." -NoNewLine
		# Get user accounts that are either Domain or Enterprise Admins
		$domainAdmins = (Get-ADGroupMember -Server $domain "Domain Admins")
		# If this is a Child Domain there is not an EA group so this handles that situation
		Try { $enterpriseAdmins = (Get-ADGroupMember -Server $domain "Enterprise Admins") }
			Catch { $enterpriseAdmins = $Null }
		# Create the CSV for the report
		$privilegedAdminsDestingation = Join-Path -Path $Path -ChildPath "PrivilegedAdmins-$csvDate.csv"
		Add-Content $privilegedAdminsDestingation "Name,SamAccountName,Membership" | Wait-Job
		# Add information generated to the CSV for Domain Admins
		if ($domainAdmins -ne $Null) {
			ForEach ($user in $domainAdmins) {
				$daName = $user.Name
				$daSamName = $user.samAccountName
				Add-Content $privilegedAdminsDestingation "$daName,$daSamName,Domain Admin" | Wait-Job
			}
		}
		# Add information generated to the CSV for Enterprise Admins
		if ($enterpriseAdmins -ne $Null) {
			ForEach ($user in $enterpriseAdmins) {
				$eaName = $user.Name
				$eaSamName = $user.samAccountName
				Add-Content $privilegedAdminsDestingation "$eaName,$eaSamName,Enterprise Admin" | Wait-Job
			}
		}
		$domainAdminsCount = $domainAdmins | Measure-Object
		$enterpriseAdminsCount = $enterpriseAdmins | Measure-Object
		Write-Host -ForegroundColor Red " DomainAdmins:" $domainAdminsCount.Count " EnterpriseAdmins:" $enterpriseAdminsCount.Count
		$body += "Domain Administrator Accounts: <strong>$($domainAdminsCount.Count)</strong><br />"
		$body += "Enterprise Administrator Accounts: <strong>$($enterpriseAdminsCount.Count)</strong><br />"
	}


### Expired Users ###############################
	if ($enableExpiredUsers -eq "YES") {
		Write-Host -ForegroundColor Green "Removing expired user accounts..." -NoNewLine
		# Get user accounts that have met or passed their expiration date if given one
		$expiredUsersSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADUser -Server $domain -SearchBase $searchRoot -Filter * -Properties * | Where-Object { ($_.AccountExpirationDate -ne $Null -AND $_.AccountExpirationDate -lt (Get-Date)) } }
		$expiredUsers = ForEach ($searchEntry in $expiredUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogonTimestamp -Descending | Select-Object -First 1 }
		if ($expiredUsers -ne $Null) {
			# Create the CSV for the report
			$expiredUserDestination = Join-Path -Path $Path -ChildPath "ExpiredUsers-$csvDate.csv"
			Add-Content $expiredUserDestination "Name,SamAccountName,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($user in $expiredUsers) {
				$Name = $user.Name
				$SamAccountName = $user.SamAccountName
				$userExpires = $user.LastLogonTimestamp
				if (-not ($userExpires)) {$userExpires = 0 }
				if (($userExpires -eq 0) -or ($userExpires -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$userExpires
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $expiredUserDestination "$Name,$SamAccountName,$LastLogon" | Wait-Job
				# If Testing is disabled then go ahead and disable the account and move it
				if ($Testing -eq "NO") {
					Set-ADUser -Identity ($user) -Description "Expired Account Disabled on $csvDate" -Enabled $false
					Move-ADObject -Identity ($user) -TargetPath $disableRoot
				}
			}
		}
		$expiredUserCount = $expiredUsers | Measure-Object
		Write-Host -ForegroundColor Red $expiredUserCount.Count
		$body += "Expired Users Disabled and Moved: <strong>$($expiredUserCount.Count)</strong><br />"
	}


### Check-In Computers ##########################
	if ($enableExpiringComps -eq "YES") {
		Write-Host -ForegroundColor Green "Looking up computers that have been offline too long..." -NoNewLine
		# Get computer accounts that have not connected to the domain in a certain amount of time and warn their users to do so
		$expiringComputersTime = (get-date).adddays(-$warningDays)
		$expiringCompsSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADComputer -Server $_.Hostname -SearchBase $searchRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $expiringComputersTime) -AND (PasswordLastSet -le $expiringComputersTime)) -AND (Enabled -eq $True)} -Properties * | Where-Object {$_.description -notmatch $disableExempt} | Sort-Object Name }
		$expiringComputers = ForEach ($searchEntry in $expiringCompsSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($expiringComputers -ne $Null) {
			# Create the CSV for the report
			$expiredComputerDestination = Join-Path -Path $Path -ChildPath "ExpiringComputers-$csvDate.csv"
			Add-Content $expiredComputerDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach ($computer in $expiringComputers) {
				$expiringCompsEmail = $computer.Description
				$expiringCompsName = $computer.Name
				$computerExpiring = $computer.LastLogon
				if (-not ($computerExpiring)) { $computerExpiring = 0 }
				if (($computerExpiring -eq 0) -or ($computerExpiring -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$computerExpiring
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				# Gets owner email from object Description field and composes email message
				if (($expiringCompsEmail) -Like "*coriosgroup.com") {
					$expiringCompsSubject="Your Computer Has Not Checked In Recently"
					$eCompsbody = "
						<span style='font-family:calibri;'>
						<p>Hello, <br /> Your computer ($expiringCompsName) has not connected to the $companyName network in more than $warningDays days. If you do not connect to the $companyName network in an office or through VPN regularly, your account may be disabled. Connecting to the network allows us to make sure your computer is kept current and safe as well as allowing us to provide aditional helpful applications, inventory management, and support options.</p>
						<p>To avoid interruptions to your work, please connect to the VPN as soon as possible so that your computer can get the latest security updates and system patches. If you have any questions please submit a Service Desk ticket at $serviceDesk. </p>
						<p>Thank you, <br /> $companyName IT</p></span>"
					if ($enableQuotes -eq "YES") {
						$expiringComputersQuote = Get-Quote
						$eCompsbody += "<br /><br /><span style='font-family:arial;color:green;'><blockquote>$expiringComputersQuote</blockquote></span>"
					}
				}
				# If no email address in object Description filed, send the notice to the admins instead
				Else {
					$expiringCompsEmail = $toAddr
					$expiringCompsSubject = "Expiring Computer With No Owner Email"
					$eCompsbody = "
						<span style='font-family:calibri;'>
						<p>The machine <strong>$expiringCompsName</strong> has not checked in for more than $warningDays days and does not have an email associated with it in the description field.</p>
						<p>Please enter an email address in the description field as soon as possible so the owner will be notified.</p></span>"
				}
				Add-Content $expiredComputerDestination "$expiringCompsName,$LastLogon" | Wait-Job
				# Send Email Message
					if ($Testing -eq "YES") { $expiringCompsEmail = $toAddr }
					Send-Mailmessage -SmtpServer $smtpServer -From $fromAddr -To $expiringCompsEmail -Subject $expiringCompsSubject -Body $eCompsbody -BodyAsHTML -Priority High
			}
		}
		$expiringComputersCount = $expiringComputers | Measure-Object
		Write-Host -ForegroundColor Red $expiringComputersCount.Count
		$body += "Computer Check-In Warnings Sent: <strong>$($expiringComputersCount.Count)</strong><br />"
	}


### Inactive Computers ##########################
	if ($enableInactiveComps -eq "YES") {
		Write-Host -ForegroundColor Green "Removing inactive computer accounts..." -NoNewLine
		$inactiveComputersTime = (get-date).adddays(-$inactiveDays)
		# Get computer accounts that have not connected to the domain in a certain amount of time and disable them
		$inactiveCompsSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADComputer -Server $_.Hostname -SearchBase $searchRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $inactiveComputersTime) -AND (PasswordLastSet -le $inactiveComputersTime)) -AND (Enabled -eq $True)} -Properties * | Where-Object {$_.description -notmatch $disableExempt} | Sort-Object Name }
		$inactiveComputers = ForEach ($searchEntry in $inactiveCompsSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($inactiveComputers -ne $Null) {
			# Create the CSV for the report
			$inactiveCompsDestination = Join-Path -Path $Path -ChildPath "InactiveComputers-$csvDate.csv"
			Add-Content $inactiveCompsDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($computer in $inactiveComputers){
				$InactiveComputerName = $computer.Name
				$compInactive = $computer.LastLogon
				if (-not ($compInactive)) { $compInactive = 0 }
				if (($compInactive -eq 0) -or ($compInactive -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$compInactive
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $inactiveCompsDestination "$InactiveComputerName,$LastLogon" | Wait-Job
				# If Testing is disabled, then go ahead and disable the computer object and move it
				if ($Testing -eq "NO") {
					Set-ADComputer -Identity ($computer) -Description "Disabled on $csvDate for Inactivity" -Enabled $false
					Move-ADObject -Identity ($computer) -TargetPath $disableRoot
				}
			}
		}
		$inactiveComputersCount = $inactiveComputers | Measure-Object
		Write-Host -ForegroundColor Red $inactiveComputersCount.Count
		$body += "Inactive Computer Objects Disabled: <strong>$($inactiveComputersCount.Count)</strong><br />"
	}


### Inactive Users ##############################
	if ($enableInactiveUsers -eq "YES") {
		Write-Host -ForegroundColor Green "Removing inactive user accounts..." -NoNewLine
		# Get user accounts that have not signed in to the domain in a certain amount of time and disable them
		$inactiveUsersTime = (get-date).adddays(-$inactiveDays)
		$inactiveUsersSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADUser -Server $_.Hostname -SearchBase $searchRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $inactiveUsersTime) -AND (PasswordLastSet -le $inactiveUsersTime) -AND (whenCreated -le $createdDate)) -AND (Enabled -eq $True)} -Properties * | Where-Object {$_.description -notmatch $disableExempt} | Sort-Object Name }
		$inactiveUsers = ForEach ($searchEntry in $inactiveUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($inactiveUsers -ne $Null) {
			# Create the CSV for the report
			$inactiveUsersDestination = Join-Path -Path $Path -ChildPath "InactiveUsers-$csvDate.csv"
			Add-Content $inactiveUsersDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($user in $inactiveUsers){
				$InactiveUserName = $user.Name
				$usersInactive = $user.LastLogon
				if (-not ($usersInactive)) { $usersInactive = 0 }
				if (($usersInactive -eq 0) -or ($usersInactive -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$usersInactive
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $inactiveUsersDestination "$InactiveUserName,$LastLogon" | Wait-Job
				# If Testing is disabled, then go ahead and disable the user object and move it
				if ($Testing -eq "NO") {
					Set-ADUser -Identity ($user) -Description "Disabled on $csvDate for Inactivity" -Enabled $false
					Move-ADObject -Identity ($user) -TargetPath $disableRoot
				}
			}
		}
		$inactiveUsersCount = $inactiveUsers | Measure-Object
		Write-Host -ForegroundColor Red $inactiveUsersCount.Count
		$body += "Inactive User Objects Disabled: <strong>$($inactiveUsersCount.Count)</strong><br />"
	}


### Temporary User Accounts #####################
	if ($enableTempUserWarnings -eq "YES") {
		Write-Host -ForegroundColor Green "Finding expiring/temporary users..." -NoNewLine
		$TempUsers = (Get-ADUser -Server $domain -SearchBase $searchRoot -Filter * -Properties AccountExpirationDate,Manager,Mail | Where-Object { ($_.AccountExpirationDate -NE $Null -AND $_.AccountExpirationDate -LT ($expireDate) -AND $_.Enabled -eq "True") } | Sort-Object Name)
		if ($TempUsers -ne $Null) {
			$TempUserDestination = Join-Path -Path $Path -ChildPath "TempUsers-$csvDate.csv"
			Add-Content $TempUserDestination "Name,DaysToExpire" | Wait-Job
			ForEach ($user in $TempUsers) {
				$name = $user.Name
				$givenName = $user.GivenName
				$samName = $user.SamAccountName
				$manager = $user.Manager
				if ($manager -eq $Null) { $managerEmail = $Null }
				else {
					$managerAD = (Get-ADUser -Identity $manager -Properties Mail)
					$managerEmail = $managerAD.Mail
					$managerName = $managerAD.GivenName
				}
				$expiresOn = $user.AccountExpirationDate
				$daysToExpire = (New-TimeSpan -Start $today -End $expiresOn).Days
				$expireOnDate = (Get-Date $user.AccountExpirationDate -UFormat "%A, %d %B, %Y")

				# If Testing Is Enabled - Email Administrator
				if ($testing -eq "YES") { $emailAddress = $TestAddress }
				else { $emailAddress = $user.Mail } 

				# Check Number of Days to Expiry
				$messageDays = $daysToExpire

				# Set Greeting based on Number of Days to Expiry
				if (($messageDays) -gt "1") { $messageDays = "in " + "$daysToExpire" + " days" }
				else { $messageDays = "today" }

				# Email Subject Set Here
				$subjectUser="Your account will expire $messageDays"
				$subjectManager="Account for $name will expire $messageDays"
			  
				# Email Body Set Here, Note You can use HTML, including Images
				$bodyUser ="
				<p style='font-family:calibri'>Hello $givenName,</p>
				<p style='font-family:calibri'>Your Viewpoint account <strong>VIEWPOINT\$samName</strong> will expire <font color=red><strong>$messageDays</strong></font>.</p>
				<p style='font-family:calibri'>Your account has been set to expire most likely because you are listed as either a Temp or a Contractor. If your contract is being extended, please speak with your manager as soon as possible to have your account expiration date extended. If you do not do this, your account will be disabled on $expireOnDate.</p>
				<p style='font-family:calibri'>Thank you,<br /> Viewpoint IT</p>
				"

				$bodyManager ="
				<p style='font-family:calibri'>Hello $managerName,</p>
				<p style='font-family:calibri'>The user account for <strong>$name</strong> will expire <font color=red><strong>$messageDays</strong></font>.</p>
				<p style='font-family:calibri'>If this person's contract is to be extended, please contact HR to extend the date to avoid interruption. Account actions such as contract extensions must start with an HR request in order to proceed.</p>
				<p style='font-family:calibri'>Thank you,<br /> Viewpoint IT</p>
				"
				
				Add-Content $TempUserDestination "$Name,$DaysToExpire" | Wait-Job
				
				# If a user has no email address listed
				if (($emailAddress) -eq $Null) { $emailAddress = $toAddr }

				# Send messages
				Send-Mailmessage -smtpServer $smtpServer -From $fromAddr -to $emailAddress -subject $subjectUser -body $bodyUser -BodyAsHTML -Priority High -Encoding $textEncoding
				if ($managerEmail -ne $Null) {
					Send-Mailmessage -smtpServer $smtpServer -From $fromAddr -to $managerEmail -subject $subjectManager -body $bodyManager -BodyAsHTML -Priority High -Encoding $textEncoding
				}
			}
		}
		$TempUsersCount = $TempUsers | Measure-Object
		Write-Host -ForegroundColor Red $TempUsersCount.Count
		$body += "Temp Users Notified: <strong>$($TempUsersCount.Count)</strong><br />"
	}


### Terminated User Cleanup #####################
	if ($enableTermUsers -eq "YES") {
		Write-Host -ForegroundColor Green "Removing accounts for terminated users..." -NoNewLine
		# Get disabled user accounts belonging to terminated users and move them when they've reached the Inactive Days limit
		$termUsersTime = (get-date).adddays(-$inactiveDays)
		$termUsersSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADUser -Server $_.Hostname -SearchBase $searchRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $termUsersTime) -AND (Enabled -eq $False))} -Properties * | Where-Object {$_.description -match $termDesignator} | Sort-Object Name }
		$termUsers = ForEach ($searchEntry in $termUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($termUsers -ne $Null) {
			# Create the CSV for the report
			$termUsersDestination = Join-Path -Path $Path -ChildPath "TermedUsers-$csvDate.csv"
			Add-Content $termUsersDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($user in $termUsers) {
				$TermedUserName = $user.Name
				$termUser = $user.LastLogon
				if (-not ($termUser)) { $termUser = 0 }
				if (($termUser -eq 0) -or ($termUser -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$termUser
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $termUsersDestination "$TermedUserName,$LastLogon" | Wait-Job
				# If Testing is disabled, then go ahead and move the object
				if ($Testing -eq "NO") {
					Set-ADUser -Identity ($user) -Description "Auto-Cleanup on $csvDate" -Enabled $false
					Move-ADObject -Identity ($user) -TargetPath $disableRoot
				}
			}
		}
		$termUsersCount = $termUsers | Measure-Object
		Write-Host -ForegroundColor Red $termUsersCount.Count
		$body += "Terminated User Objects Moved: <strong>$($termUsersCount.Count)</strong><br />"
	}


### Disabled Computer Cleanup #######################
	if ($enableDisabledComps -eq "YES") {
		Write-Host -ForegroundColor Green "Removing disabled computer accounts..." -NoNewLine
		# Get disabled computer accounts and move them
		$disabledComps = (Get-ADComputer -Server $domain -SearchBase $searchRoot -Filter {(Enabled -eq $False)} -Properties * | Where-Object {$_.description -notmatch $termDesignator -AND $_.description -notmatch $deleteExempt} | Sort-Object Name)
		if ($disabledComps -ne $Null) {
			# Create the CSV for the report
			$disabledCompsDestination = Join-Path -Path $Path -ChildPath "DisabledComps-$csvDate.csv"
			Add-Content $disabledCompsDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($comp in $disabledComps) {
				$DisabledComputerName = $comp.Name
				$compDisabled = $comp.LastLogon
				if (-not ($compDisabled)) { $compDisabled = 0 }
				if (($compDisabled -eq 0) -or ($compDisabled -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$compDisabled
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $disabledCompsDestination "$DisabledComputerName,$LastLogon" | Wait-Job
				# If Testing is disabled, then go ahead and move the object
				if ($Testing -eq "NO") {
					Set-ADComputer -Identity ($comp) -Description "Auto-Cleanup on $csvDate" -Enabled $false
					Move-ADObject -Identity ($comp) -TargetPath $disableRoot
				}
			}
		}
		$disabledCompsCount = $disabledComps | Measure-Object
		Write-Host -ForegroundColor Red $disabledCompsCount.Count
		$body += "Disabled Computer Objects Moved: <strong>$($disabledCompsCount.Count)</strong><br />"
	}


### Disabled User Cleanup #######################
	if ($enableDisabledUsers -eq "YES") {
		Write-Host -ForegroundColor Green "Removing disabled user accounts..." -NoNewLine
		# Get disabled user accounts and move them
		$disabledUsers = (Get-ADUser -Server $domain -SearchBase $searchRoot -Filter {(Enabled -eq $False)} -Properties * | Where-Object {$_.description -notmatch $termDesignator -AND $_.description -notmatch $deleteExempt} | Sort-Object Name)
		if ($disabledUsers -ne $Null) {
			# Create the CSV for the report
			$disabledUsersDestination = Join-Path -Path $Path -ChildPath "DisabledUsers-$csvDate.csv"
			Add-Content $disabledUsersDestination "Name,LastLogon" | Wait-Job
			# Add information generated to the CSV
			ForEach($user in $disabledUsers) {
				$DisabledUserName = $user.Name
				$userDisabled = $user.LastLogon
				if (-not ($userDisabled)) { $userDisabled = 0 }
				if (($userDisabled -eq 0) -or ($userDisabled -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$userDisabled
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $disabledUsersDestination "$DisabledUserName,$LastLogon" | Wait-Job
				# If Testing is disabled, then go ahead and move the object
				if ($Testing -eq "NO") {
					Set-ADUser -Identity ($user) -Description "Auto-Cleanup on $csvDate" -Enabled $false
					Move-ADObject -Identity ($user) -TargetPath $disableRoot
				}
			}
		}
		$disabledUsersCount = $disabledUsers | Measure-Object
		Write-Host -ForegroundColor Red $disabledUsersCount.Count
		$body += "Disabled User Objects Moved: <strong>$($disabledUsersCount.Count)</strong><br />"
	}


### Delete Old Computers ########################
	if ($enableDeleteComps -eq "YES") {
		Write-Host -ForegroundColor Green "Deleting old computer accounts..." -NoNewLine
		$deleteCompsSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADComputer -Server $_.Hostname -SearchBase $disableRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $deleteTime) -AND (PasswordLastSet -le $deleteTime)) -AND (Enabled -eq $False)} -Properties LastLogon,Description | Where-Object {$_.description -notmatch $deleteExempt} | Sort-Object Name }
		$deleteComputers = ForEach ($searchEntry in $deleteCompsSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($deleteComputers -ne $Null) {
			$deleteCompsDestination = Join-Path -Path $Path -ChildPath "DeletedComputers-$csvDate.csv"
			Add-Content $deleteCompsDestination "Name,LastLogon" | Wait-Job
			ForEach($computer in $deleteComputers){
				$deletedCompName = $computer.Name
				$deletedComp = $computer.LastLogon
				if (-not ($deletedComp)) { $deletedComp = 0 }
				if (($deletedComp -eq 0) -or ($deletedComp -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$deletedComp
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $deleteCompsDestination "$deletedCompName,$LastLogon" | Wait-Job
				if ($Testing -eq "NO") {
					Remove-ADObject -Identity ($computer)
				}
			}
		}
		$deleteComputersCount = $deleteComputers | Measure-Object
		Write-Host -ForegroundColor Red $deleteComputersCount.Count
		$body += "Computer Objects Deleted: <strong>$($deleteComputersCount.Count)</strong><br />"
	}


### Delete Old Users ############################
	if ($enableDeleteUsers -eq "YES") {
		Write-Host -ForegroundColor Green "Deleting old user accounts..." -NoNewLine
		$deleteUsersSearch = Get-ADDomainController -Server $domain -Filter * | ForEach-Object { Get-ADUser -Server $_.Hostname -SearchBase $disableRoot -Filter {((LastLogon -notlike "*" -OR LastLogon -le $deleteTime) -AND (PasswordLastSet -le $deleteTime)) -AND (Enabled -eq $False)} -Properties LastLogon,Description | Where-Object {$_.description -notmatch $deleteExempt} | Sort-Object Name }
		$deleteUsers = ForEach ($searchEntry in $deleteUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogon -Descending | Select-Object -First 1 }
		if ($deleteUsers -ne $Null) {
			$deleteUsersDestination = Join-Path -Path $Path -ChildPath "DeletedUsers-$csvDate.csv"
			Add-Content $deleteUsersDestination "Name,LastLogon"  | Wait-Job
			ForEach($user in $deleteUsers) {
				$deletedUserName = $user.Name
				$deletedUser = $user.LastLogon
				if (-not ($deletedUser)) { $deletedUser = 0 }
				if (($deletedUser -eq 0) -or ($deletedUser -gt [DateTime]::MaxValue.Ticks)) {
					$LastLogon = "<Never>"
				}
				Else {
					$Date = [DateTime]$deletedUser
					$LastLogon = $Date.AddYears(1600).ToLocalTime()
				}
				Add-Content $deleteUsersDestination "$deletedUserName,$LastLogon" | Wait-Job
				if ($Testing -eq "NO") {
					Remove-ADObject -Identity ($user)
				}
			}
		}
		$deleteUsersCount = $deleteUsers | Measure-Object
		Write-Host -ForegroundColor Red $deleteUsersCount.Count
		$body += "User Objects Deleted: <strong>$($deleteUsersCount.Count)</strong><br />"
	}


### Domain Password Notice ######################
	if ($enableDomainPasswords -eq "YES") {
		Write-Host -ForegroundColor Green "Looking up user accounts that need to change their password..." -NoNewLine
		$textEncoding = [System.Text.Encoding]::UTF8
		# Create the CSV for the report
		$domainPwdDestination = Join-Path -Path $Path -ChildPath "DomainPwdExpiration-$csvDate.csv"
		Add-Content $domainPwdDestination "Date,Name,EmailAddress,DaystoExpire,ExpiresOn,Notified" | Wait-Job
		# Get list of all enabled users whose passwords will expire
		$users = Get-ADUser -Server $domain -Filter * -Properties * | Where-Object {$_.Enabled -eq "True"} | Where-Object { $_.PasswordNeverExpires -eq $false } | Where-Object { $_.passwordexpired -eq $false }
		$DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
		# Check the users to see which ones fall within the notification window
		ForEach ($user in $users) {
			$Name = $user.Name
			$GivenName = $user.GivenName
			$EmailAddress = $user.EmailAddress
			$samName = $user.SamAccountName
			$UPN = $user.UserPrincipalName
			$passwordSetDate = $user.PasswordLastSet
			$PasswordPol = (Get-AduserResultantPasswordPolicy $user)
			$sent = ""
			# Check for Fine Grained Password or set to Domain Default
			if (($PasswordPol) -ne $Null) { $maxPasswordAge = ($PasswordPol).MaxPasswordAge }
			Else { $maxPasswordAge = $DefaultmaxPasswordAge }
			# Set Greeting based on Number of Days to Expiry.
			$expireson = $passwordsetdate + $maxPasswordAge
			$today = (Get-Date)
			$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
			$messageDays = $daystoexpire
			if (($messageDays) -gt "1") { $messageDays = "in " + "$daystoexpire" + " days" }
			Else { $messageDays = "today" }

			# Email Content
			$pwdSubject="Your password will expire $messageDays"
			$pwdBody ="
			<span style='font-family:calibri;'>
			<p>Hello $GivenName,</p>
			<p>Your Windows login password for <strong>$UPN</strong> will expire <font color=red><strong>$messageDays</strong></font>.</p>
			<p>Please make sure you are connected to the $companyName network or on VPN before proceeding.</p>
			<p>Press CTRL-ALT-DEL and select 'Change a password...'</p>
			<p>Enter your old password in the top box, then your new password twice to confirm it.</p>
			<br />
			<p>Requirements for the password are as follows:</p>
			<ul>
			<li>Must be at least $pwdMinimum characters long.</li>
			<li>Must not contain the user's account name or parts of the user's full name that exceed two consecutive characters</li>
			<li>Must not be one of your last 24 passwords</li>
			<li>Contain characters from three of the following four categories:</li>
			<ul>
			<li>English UPPERCASE characters (A through Z)</li>
			<li>English lowercase characters (a through z)</li>
			<li>Base 10 digits (0 through 9)</li>
			<li>Non-alphabetic characters (for example, !, $, #, %)</li>
			</ul></ul>
			<p>For any assistance, please create a ticket at $serviceDesk and someone will assist you as soon as possible.</p></span>"
			if ($enableQuotes -eq "YES") {
				$domainPwdQuote = Get-Quote
				$pwdBody += "<br /><br /><span style='font-family:arial;color:green;'><blockquote>$domainPwdQuote</blockquote></span>"
			}
			# If Testing Is Enabled - Email Administrator
			if (($Testing) -eq "YES") { $EmailAddress = $toAddr }
			# If a user has no email address listed - Email Administrator
			if (($EmailAddress) -eq $Null) { $EmailAddress = $toAddr }
			# Add CSV Content and Send Email Message
			if (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireInDays)) {
				$sent = "Yes"
				Add-Content $domainPwdDestination "$csvDate,$Name,$EmailAddress,$daystoExpire,$expireson,$sent"  | Wait-Job
				Send-Mailmessage -SmtpServer $smtpServer -From $fromAddr -To $EmailAddress -Subject $pwdSubject -Body $pwdBody -BodyAsHTML -Priority High -Encoding $textEncoding
			}
		}
		$userCount = Import-Csv $domainPwdDestination | Measure-Object | Select-Object -Expand Count
		Write-Host -ForegroundColor Red $userCount
		$body += "Password Expiration Notifications Sent: <strong>$($userCount)</strong><br />"
	}


### AD Health Checks ############################
	if ($enableADHealthChecks -eq "YES") {
		if ($adHealthDayOfWeek -eq $adHealthDay -OR $adHealthDayOfWeek -eq "ALL") {
			Write-Host -ForegroundColor Green "Running AD Diagnostics and Replication Checks..." -NoNewLine
			# Create the Replication Summary report
			$repFile = Join-Path -Path $Path -ChildPath "replsummary-$csvDate.txt"
			CMD /C "repadmin /replsummary 2>&1" | Out-File $repFile | Wait-Job
			# Convoluted way to use a variable entry to run the DCDIAG Enterprise report
			$diagFile = Join-Path -Path $Path -ChildPath "dcdiag-$csvDate.txt"
			$diag = "CMD"
			$diagDom = "/n:" + $domain
			[Array]$diagParams = "/c", "dcdiag", "/e", "/v", "/c", $diagDom, "2>&1"
			& $diag $diagParams | Out-File $diagFile | Wait-Job
			Write-Host -ForegroundColor Red "Done"
		}
	}


### DFS Report ##################################
	if ($enableDFSExport -eq "YES") {
		Write-Host -ForegroundColor Green "Looking up DFS folders and locations..." -NoNewLine
		function Get-DfsnAllFolderTargets () {
			#Get a list of all Namespaces in the Domain
			Write-Progress -Activity "1/3 - Getting List of Domain NameSpaces"
			$RootList = Get-DfsnRoot

			#Get a list of all FolderPaths in the Namespaces
			Write-Progress -Activity "2/3 - Getting List of Domain Folder Paths"
			$FolderPaths = ForEach ($item in $RootList) {
				Get-DfsnFolder -Path "$($item.path)\*"
			}

			#Get a list of all Folder Targets in the Folder Paths, in the Namespaces"
			Write-Progress -Activity "3/3 - Getting List of Folder Targets"
			$FolderTargets = ForEach ($item in $FolderPaths) {
				Get-DfsnFolderTarget -Path $item.Path
			}
			return $FolderTargets
		}
		$dfsDestination = Join-Path -Path $Path -ChildPath "DFSExport-$csvDate.csv"
		Get-DfsnAllFolderTargets | Export-Csv -Path $dfsDestination | Wait-Job
		$body += "DFS Folder Locations Attached as <strong>DFSExport-$csvDate.csv</strong><br />"
		Write-Host -ForegroundColor Red "Done"
	}


### Email The Lot ###############################
	if ($emailReports -eq "YES") {
		$subject = "Daily Reports for $domain on $csvDate"
		$body += "Reports Run: $reportDate<br />"
		if ($enableQuotes -eq "YES") {
			$reportQuote = Get-Quote
			$body += "<br /><br /><span style='font-family:arial;color:green;'><blockquote>$reportQuote</blockquote></span>"
		}
		Write-Host -ForegroundColor Cyan "Gathering the reports and sending the results..." -NoNewLine
		Start-Sleep -Seconds 5
		# Gets all files in the Path folder except the script and attaches them to the email.
		Get-ChildItem -Path $Path -Exclude *.ps1 | ForEach-Object {$_.FullName} | Send-MailMessage -To $toAddr -From $fromAddr -Subject $subject -Body "$body" -SmtpServer $smtpServer -BodyAsHtml
		Write-Host -ForegroundColor Red "Complete"
	}
