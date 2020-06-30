    <#
        .SYNOPSIS
        Get the public IP address of the machine this is run on and send it as an email or place a text file somewhere.

        .PARAMETER Output
        Output to EMAIL or TEXT file.
		
		.PARAMETER Path
		Default location for text is the user's desktop.

		.USAGE
		.\IpInfo.ps1 -Output TEXT (-Path C:\Some\Location)
    #>
#################################################
<#
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    https://github.com/dlindridge/UsefulPowershell
    Created: September 23, 2019
    Modified: October 22, 2019
#>
#################################################

	Param (
        [Parameter(Mandatory=$True)]
		$Output,
        [Parameter(Mandatory=$False)]
		$Path
	)

$fromAddr = "MyCompany IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
$toAddr = "itadmins@MyDomain.TLD" # Enter the TO address for the e-mail alert - Must be inside quotes.
$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.

$date = Get-Date -DisplayHint Date
$ipinfo = Invoke-RestMethod http://ipinfo.io/json
$server = Get-Content env:COMPUTERNAME
$ip = $ipinfo.ip
$hostname = $ipinfo.hostname
$city = $ipinfo.city
$region = $ipinfo.region
$country = $ipinfo.country
$loc = $ipinfo.loc
$org = $ipinfo.org
$Profile = $env:USERPROFILE
if ($Path -eq "" -OR $Path -eq $Null) { $Path = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\$server.txt" }
else { $Path = Join-Path -Path $Path -ChildPath "$server.txt" }

$Output = $Output.ToUpper()

if ($Output -eq "EMAIL") {
	$body = "$ip <br />$hostname <br />$city <br />$region <br />$country <br />$loc <br />$org <br />"
	Send-MailMessage -To $toAddr -From $fromAddr -Subject "$server" -Body "$body" -SmtpServer $smtpsrv -BodyAsHtml
}

if ($Output -eq "TEXT") {
	Get-ChildItem -Path $Path | Remove-Item -Force
	$textOut = $Path
	Add-Content $textOut $server
	Add-Content $textOut $date
	Add-Content $textOut $ip
	Add-Content $textOut $hostname
	Add-Content $textOut $city
	Add-Content $textOut $region
	Add-Content $textOut $country
	Add-Content $textOut $loc
	Add-Content $textOut $org
}
