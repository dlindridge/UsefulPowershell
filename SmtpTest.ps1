    <#
        .SYNOPSIS
        Sends a test email to address specified.

        .PARAMETER Server
        The SMTP server name - FQDN recommended.

        .PARAMETER Recipient
        The destination email address.

        .DESCRIPTION
        Usage: .\SmtpTest.ps1 (-Server smtprelay.somedomain.net)
    #>

#################################################
<#
Simple SMTP relay test to validate your relay settings.

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: September 7, 2017
Modified: October 6, 2022
#>
#################################################

    Param (
        $Server = "smtp.MyDomain.TLD",
        [Parameter(Mandatory=$True)]
        $Recipient
    )

######################

$LocalHost = Get-Content ENV:ComputerName

$From = "SMTP Test <noreply@mydomain.tld>" # Enter the FROM address for the e-mail alert - Must be inside quotes.

$body = "<p style='font-family:calibri'>This is a test of your SMTP relay settings and authorizations on $Server.</p>
         <p style='font-family:calibri'>If you are reading this, it worked.</p>
         <p style='font-family:calibri'>This test was run from $LocalHost.</p>"

Send-MailMessage -To $Recipient -From $From -Subject "Info: SMTP Relay Test" -Body "$body" -SmtpServer $Server -BodyAsHtml

######
