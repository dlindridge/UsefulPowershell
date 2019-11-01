<#
    .SYNOPSIS
    Scotty: Admiral, there be whales here!

    .PARAMETER Active
    YES or NO: Set the created signature blocks as active. Default is "YES".

    .DESCRIPTION
    Using an HTML template for an email signature, customize it for each user in a specified source CSV to use as a
    signature block with their local Outlook client. Ideal for running as a login script to keep signature information
    current from Active Directory.
    Usage:  powershell.exe -file (path)\VoyageHome.ps1 as part of a logon script.
            *OR* Set PowerShell option for logon script to call the ps1 directly. Don't forget to change the logon
            script wait time, MS default is 5 minutes (!)
#>
#################################################
<#
    TO-DO:  Input parameters for signature template filenames.

    Significant help came from Chris Meeuwen with a source script for updating OWA signatures which morphed into this.
        -Chris, thanks for the source info, it gave me places to go!
    
   	When building your signature template use the placeholders below. If the mobile/cell number is blank in AD the line
    that contains it's placeholder is skipped. Write your signature template accordingly.
        PS-USER-NAME = User's Full Name
        PS-TITLE-NAME = User's Title
        PS-OFFICE-PHONE = User's Office Phone Number
        PS-MOBILE-PHONE = User's Cell/Mobile Phone Number
        PS-EMAIL-ADDR = User's Email Address

    See here for logon scripting with Powershell: http://woshub.com/running-powershell-startup-scripts-using-gpo/
    *NOTE:  Even though this is a USER policy in Group Policy, still follow the directions to add Domain Computers to 
            security Read & Execute on that folder.

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 30, 2019
    Modified: October 31, 2019
#>
#################################################

### Build some organization specific variables
$MainName = "Main" # Final name of the New email signature block
$ReplyName = "Reply" # Final name of the Reply email signature block
$MainSignatureFile = "\\MyDomain.org\NETLOGON\EmailSig_Main.html" # Shared file location for the New/Main signature template
$ReplySignatureFile = "\\MyDomain.org\NETLOGON\EmailSig_Reply.html" # Shared file location for the Reply signature template

### There Be Script Here! #######################

Param (
    $Active = "YES"
)
# Check to see if we can talk to the Main signature file - skip everything if we can't
If (Test-Path $MainSignatureFile) {
    ### Path to user's AppData folder and default signature folder for Outlook
    $AppData = (Get-Item ENV:AppData).Value
    $LocalSignatureFolder = $AppData+'\Microsoft\Signatures'
    If (Test-Path $LocalSignatureFolder) { Write-Output "Signature folder already exists" } Else { New-Item -Path $LocalSignatureFolder -ItemType "Directory" -Force }
    
    ### Get the current logged in username
    $UserName = (Get-Item ENV:UserName).Value

    ### The Query AD and get an ADUser Object
    $Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
    $Searcher = New-Object System.DirectoryServices.DirectorySearcher
    $Searcher.Filter = $Filter
    $ADUserPath = $Searcher.FindOne()
    $SignatureUser = $ADUserPath.GetDirectoryEntry()
    If ($SignatureUser.mobile) { $Mobile = $SignatureUser.mobile } Else { $Mobile = $Null }

    ### Build the New email signature
    $MainSignatureTemplate = Get-Content $MainSignatureFile

    ### Generate Signature
    $MainSignature= @()

    ### Find the PS-*-* fields and replace those with the appropriate details.
    ForEach ($line in $MainSignatureTemplate) {
        If ($line -like "*PS-USER-NAME*") {
            $MainSignature += $line.Replace("PS-USER-NAME","$($signatureUser.Name)")
        }
        ElseIf ($line -like "*PS-TITLE-NAME*") {
            $MainSignature += $line.Replace("PS-TITLE-NAME","$($signatureUser.Title)")
        }
        ElseIf ($line -like "*PS-OFFICE-PHONE*") {
            $MainSignature += $line.Replace("PS-OFFICE-PHONE","$($signatureUser.telephoneNumber)")
        }
        ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -ne "" -AND $Mobile -ne $Null)) {
            $MainSignature += $line.Replace("PS-MOBILE-PHONE",$Mobile)
        }
        ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -eq "" -OR $Mobile -eq $Null)) {
            Write-Output "No Mobile Phone Number"
        }
        ElseIf ($line -like "*PS-EMAIL-ADDR*") {
            $MainSignature += $line.Replace("PS-EMAIL-ADDR","$($signatureUser.emailaddress)")
        }
        Else { $MainSignature += $line }
    }

    ### Save it as Main
    $MainSignatureDest = $LocalSignatureFolder + "\" + $MainName + ".htm"
    $MainSignature > $MainSignatureDest

    ### Do it all over again for the Reply signature
    $ReplySignatureTemplate = Get-Content $ReplySignatureFile
    $ReplySignature= @()

    ForEach ($line in $ReplySignatureTemplate) {
        If ($line -like "*PS-USER-NAME*") {
            $ReplySignature += $line.Replace("PS-USER-NAME","$($signatureUser.Name)")
        }
        ElseIf ($line -like "*PS-TITLE-NAME*") {
            $ReplySignature += $line.Replace("PS-TITLE-NAME","$($signatureUser.Title)")
        }
        ElseIf ($line -like "*PS-OFFICE-PHONE*") {
            $ReplySignature += $line.Replace("PS-OFFICE-PHONE","$($signatureUser.telephoneNumber)")
        }
        ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -ne "" -AND $Mobile -ne $Null)) {
            $ReplySignature += $line.Replace("PS-MOBILE-PHONE",$Mobile)
        }
        ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -eq "" -OR $Mobile -eq $Null)) {
            Write-Output "No Mobile Phone Number"
        }
        Else { $ReplySignature += $line }
    }

    ### Save it as Reply
    $ReplySignatureDest = $LocalSignatureFolder + "\" + $ReplyName + ".htm"
    $ReplySignature > $ReplySignatureDest

    ### Set New Sigs as Active
    $Active = $Active.ToUpper()
    If ($Active -eq "YES") {
        $MSWord = New-Object -ComObject word.application
        $EmailOptions = $MSWord.EmailOptions
        $EmailSignature = $EmailOptions.EmailSignature
        $EmailSignature.NewMessageSignature = $MainName
        $EmailSignature.ReplyMessageSignature = $ReplyName
        $MSWord.Quit()
    }
}
