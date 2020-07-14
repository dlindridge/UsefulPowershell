<#
    .SYNOPSIS
	Builds a signature file for use in Outlook based on Active Directory information for the current user.

    .PARAMETER Active
    YES or NO: Set the created signature blocks as active. Default is "YES".

    .DESCRIPTION
    Usage:  powershell.exe -file (path)\OutlookSigUpdate.ps1 as part of a logon script.
            *OR* Set PowerShell option for logon script to call the ps1 directly. Don't forget to change the logon
            script wait time, MS default is 5 minutes (!)
#>
#################################################
<#
    Significant help came from Chris Meeuwen with a source script for updating OWA signatures which morphed into this.
        -Chris, thanks for the source info, it gave me places to go!
    
   	The HTML file used for the signature is built inline in the script below. Variables at the start of the script
	will allow you to cusomize information for your organization and use the built-in formatting to build a nice
	but simple signature. If you want to change the HTML code for a different look, I recommend building the signature
	file in HTML independently then piecing lines into the code below. Be sure to pay attention to formatting calls
	used by HTML as they will carry over line to line unless properly terminated each time.
	
	As written this script requires a custom AD Attribute in the user object for AWS Chime information. Modify this as
	appropriate for your needs. See here for how to create this attribute:
	https://social.technet.microsoft.com/wiki/contents/articles/20319.how-to-create-a-custom-attribute-in-active-directory.aspx

    See here for logon scripting with Powershell: http://woshub.com/running-powershell-startup-scripts-using-gpo/
    *NOTE:  Even though this is a USER policy in Group Policy, still follow the directions to add Domain Computers to 
            security Read & Execute on that folder.

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    https://github.com/dlindridge/UsefulPowershell
    Created: October 30, 2019
    Modified: July 14, 2020
#>
#################################################

Param (
    $Active = "YES"
)

### Company Specific Information ################
$MainName = "Main-Signature" # Final name of the New email signature block
$ReplyName = "Reply-Signature" # Final name of the Reply email signature block
$SocialMediaTitle = "Find us on LinkedIn" # Do you have a social media presence?
$SocialMediaURL = "https://www.linkedin.com/company/MyCompany/" # What is your social media URL?
$WebsiteDomain = "MyDomain.TLD" # Company public website domain
$WebsiteURL = "https://www.MyDomain.TLD/" # URL to the public website
$Motto = "Be the person Mr. Rogers knew you could be" # Tagline or motto
$CompanyName = "MyCompany" # Company name
$CompanyLogo = "https://MyDomain.TLD/email-logo.gif" # Public URL to company logo


### There Be Script Here! #######################

### Get Local Info ##############################
$AppData = (Get-Item ENV:AppData).Value
$LocalSignatureFolder = $AppData+'\Microsoft\Signatures'
$UserName = (Get-Item ENV:UserName).Value
If (Test-Path $LocalSignatureFolder) { Write-Output "Signature folder already exists" } Else { New-Item -Path $LocalSignatureFolder -ItemType "Directory" -Force }


### Query AD for User Object ####################
$Filter = "(&(objectCategory=User)(samAccountName=$UserName))"
$Searcher = New-Object System.DirectoryServices.DirectorySearcher
$Searcher.Filter = $Filter
$ADUserPath = $Searcher.FindOne()
$SignatureUser = $ADUserPath.GetDirectoryEntry()


### User Information ############################
$UserFullName = $SignatureUser.Name
$UserTitle = $SignatureUser.Title
If ($SignatureUser.telephoneNumber) { $OfficePhone = $SignatureUser.telephoneNumber } Else { $OfficePhone = $Null }
If ($SignatureUser.mobile) { $Mobile = $SignatureUser.mobile } Else { $Mobile = $Null }
# Requires custom AD attribute called 'awsChime' in the user attributes
If ($SignatureUser.awsChime) { $AwsChime = $SignatureUser.awsChime } Else { $AwsChime = $Null }
$UserEmail = $SignatureUser.emailAddress


### Signature Locations #########################
$MainSignatureDest = $LocalSignatureFolder + "\" + $MainName + ".htm"
$ReplySignatureDest = $LocalSignatureFolder + "\" + $ReplyName + ".htm"


### Build the Main Signature ####################
Set-Content $MainSignatureDest "<br>"
Add-Content $MainSignatureDest "<hr size=`"3`" width=`"33%`" align=`"left`">"
Add-Content $MainSignatureDest "<div style=`"font-size:10pt;  font-family: 'Calibri',sans-serif;`">"
Add-Content $MainSignatureDest "<b><span style='font-size:14.0pt;color:#D34727'>$UserFullName</span></b><br>"
Add-Content $MainSignatureDest "<span style='text-transform:uppercase'>$UserTitle</span><br>"
Add-Content $MainSignatureDest "<div><img alt=`"$CompanyName`"  src=`"$CompanyLogo`"></div>"
If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null)) {
	Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Office - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$OfficePhone</span>"
	}
If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null) -AND ($Mobile -ne "") -AND ($Mobile -ne $Null)) {
	Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
	}
If (($Mobile -ne "") -AND ($Mobile -ne $Null)) {
	Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Cell - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$Mobile</span></span>"
	}
If (($AwsChime -ne "") -AND ($AwsChime -ne $Null)) {
	Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
	Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Chime - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$AwsChime</span></span>"
	}
Add-Content $MainSignatureDest "<br><span style='font-size:10.0pt;text-transform:lowercase'>$UserEmail | <a href=`"$SocialMediaURL`"><span style='text-transform:uppercase;color:#000001'>$SocialMediaTitle</span></a> | <a href=`"$WebsiteURL`"><b>$WebsiteDomain</b></a></span><br>"
Add-Content $MainSignatureDest "<br>"
Add-Content $MainSignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>$Motto</span></i> <o:p></o:p></p>"
Add-Content $MainSignatureDest "</div>"


### Build the Reply Signature ###################
Set-Content $ReplySignatureDest "<br>"
Add-Content $ReplySignatureDest "<hr size=`"3`" width=`"33%`" align=`"left`">"
Add-Content $ReplySignatureDest "<div style=`"font-size:10pt;  font-family: 'Calibri',sans-serif;`">"
Add-Content $ReplySignatureDest "<b><span style='font-size:14.0pt;color:#D34727'>$UserFullName</span></b><br>"
Add-Content $ReplySignatureDest "<span style='text-transform:uppercase'>$UserTitle</span><br>"
If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null)) {
	Add-Content $ReplySignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Office - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$OfficePhone</span>"
	}
If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null) -AND ($Mobile -ne "") -AND ($Mobile -ne $Null)) {
	Add-Content $ReplySignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
	}
If (($Mobile -ne "") -AND ($Mobile -ne $Null)) {
	Add-Content $ReplySignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Cell - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$Mobile</span></span>"
	}
If (($AwsChime -ne "") -AND ($AwsChime -ne $Null)) {
	Add-Content $ReplySignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
	Add-Content $ReplySignatureDest "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Chime - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$AwsChime</span></span>"
	}
Add-Content $ReplySignatureDest "<br>"
Add-Content $ReplySignatureDest "<br><i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>$Motto</span></i> <o:p></o:p></p>"
Add-Content $ReplySignatureDest "</div>"


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

