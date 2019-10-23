<#
	.SYNOPSIS
	Bill Lumbergh: So, Peter, what's happening? Aahh, now, are you going to go ahead and have those TPS reports for us this afternoon?
	Peter Gibbons: No.
	Bill Lumbergh: Ah. Yeah. So I guess we should probably go ahead and have a little talk. Hmm?
	Peter Gibbons: Not right now, Lumbergh, I'm kinda busy. In fact, look, I'm gonna have to ask you to just go ahead and come back another time. I got a meeting with the Bobs in a couple of minutes.

	.PARAMETER Email
	Email YES or NO. Do you want the results of this script to go to an email report? If no, output will be put on screen.

	.DESCRIPTION
	Usage: .\ThirtySevenFlairs.ps1 -Email NO
#>
#################################################
<#
Required OS Version is 2012 or better

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: February 18, 2017
Modified: October 22, 2019
#>
#################################################

	Param (
        [Parameter(Mandatory=$True)]
		$Email
	)

$fromAddr = "MyCompany IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
$toAddr = "itadmins@MyDomain.TLD" # Enter the TO address for the e-mail alert - Must be inside quotes.
$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.
Import-Module ActiveDirectory

do {
### Script Start ###
$today = Get-Date
$Email = $Email.ToUpper()

Write-Host -ForegroundColor Green "Looking Up DFS Replication Groups..."
Write-Host
Write-Host -ForegroundColor Yellow $today
Write-Host

$ListOfDFSConnections = Get-DfsReplicatedFolder 
$replicationGroups = dfsradmin rg list
$body = @("
<table border=1 width=50% cellspacing=0 cellpadding=8 bgcolor=Black cols=3>
<tr bgcolor=White><td>Group</td><td>Source</td><td>Destination</td><td>Backlog</td></tr>")
$i = 0

    foreach ($item in $ListOfDFSConnections)
        {
        $Group = Get-DfsrConnection -GroupName $item.GroupName
        foreach ($Connection in $Group) {
                #Check OS version of source and destination as cmdlet Get-DFSRBacklog does not work lower than 2012 
                $SourceComputerNameOS = (get-adcomputer -identity $Connection.SourceComputerName -properties *).operatingsystem
                $DestinationComputerNameOS = (get-adcomputer -identity $Connection.DestinationComputerName -properties *).operatingsystem
                               
                $QueueLength = (Get-DfsrBacklog -Groupname $item.GroupName -FolderName $item.FolderName  -SourceComputerName $Connection.SourceComputerName   -DestinationComputerName $Connection.DestinationComputerName  -verbose 4>&1).Message.Split(':')[2]   
					if ($QueueLength -eq $null) {
					$Backlog = 0
						}
					else {
					$Backlog = $QueueLength
						}
					if ($Email -eq "NO") {
					write-host $item.GroupName "-" $Connection.SourceComputerName "to" $Connection.DestinationComputerName "Backlog:" $Backlog
						}
					else {
						do {
						if($i % 2){$body += "<tr bgcolor=#D2CFCF><td>$($item.GroupName)</td><td>$($Connection.SourceComputerName)</td><td>$($Connection.DestinationComputerName)</td><td>$($Backlog)</td></tr>";$i++}
						else {$body += "<tr bgcolor=#EFEFEF><td>$($item.GroupName)</td><td>$($Connection.SourceComputerName)</td><td>$($Connection.DestinationComputerName)</td><td>$($Backlog)</td></tr>";$i++}
							}
						while ($ListOfDFSConnections[$i] -ne $null)
							  }
					}
        }

$body += "</table>"

	if ($Email -eq "YES") { Send-MailMessage -SmtpServer $smtpServer -To $toAddr -From $fromAddr -Subject "DFS Replication Report" -BodyAsHtml -body "$body" }
	else {
		Write-Host
		Write-Host -ForegroundColor Red "Run Again? (y/N) " -NoNewLine
		$response = Read-Host
		$response = $response.ToLower()
		Write-Host
	}
}
While ($response -eq "y")