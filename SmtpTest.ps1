    <#
        .SYNOPSIS
        Sends a test email to address specified.

        .PARAMETER Server
        The SMTP server name - FQDN recommended.

        .PARAMETER Recipient
        The destination email address.
    #>

#################################################
<#
Just a simple SMTP relay test to validate your relay settings.

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: September 7, 2017
Modified: October 9, 2019
#>
#################################################

    Param (
        $Server = "smtp.MyDomain.TLD",
        [Parameter(Mandatory=$True)]
        $Recipient
    )

######################

$From = "SMTP Test <noreply@mydomain.tld>" # Enter the FROM address for the e-mail alert - Must be inside quotes.

$body = "<p style='font-family:calibri'>This is a test of your SMTP relay settings and authorizations on $Server.</p>
		<p style='font-family:calibri'>If you are reading this, it worked.</p>"

Send-MailMessage -To $Recipient -From $From -Subject "Info: SMTP Relay Test" -Body "$body" -SmtpServer $Server -BodyAsHtml

######
