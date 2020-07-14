<#
    .SYNOPSIS
    Using an HTML template for an email signature, customize it for each user in a specified source CSV to use as a
	signature block with OWA. Pulls data directly from Active Directory and connects to Office 365/Exchange Online
	to apply the signature and set it as default.

	.PARAMETER MFA
	Is MFA authentication required for your connection to Office 365/Exchange Online? Default: NO

    .DESCRIPTION
    Usage: .\OutlookSigUpdate.ps1 (-MFA YES).
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

	The username in the CSV file should be the user's SamAccountName in AD with the column heading "UserName"

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 30, 2019
    Modified: July 14, 2020
#>
#################################################

### Build some organization specific variables
$Script:Users = Import-Csv "C:\Scripts\ADUserInfo\ADUserInfo.csv" # Name and Location of user CSV list
$Script:SocialMediaTitle = "Find us on LinkedIn" # Do you have a social media presence?
$Script:SocialMediaURL = "https://www.linkedin.com/company/MyCompany/" # What is your social media URL?
$Script:WebsiteDomain = "MyDomain.TLD" # Company public website domain
$Script:WebsiteURL = "https://www.MyDomain.TLD/" # URL to the public website
$Script:Motto = "Be the person Mr. Rogers knew you could be" # Tagline or motto
$Script:CompanyName = "MyCompany" # Company name
$Script:CompanyLogo = "https://MyDomain.TLD/email-logo.gif" # Public URL to company logo


### Script Starts - Go No Further ###############
Function Connect-Office365 {
	### This is copy-pasta from another script trimmed to make it work only with Exchange (for simplicity sake) if you have MFA enabled. DO NOT TOUCH!!
	### Seriously, don't mess with this and just go straight to the next function. That is where the heavy lifting happens.
	[OutputType()]
	[CmdletBinding(DefaultParameterSetName)]
	Param (
		[ValidateSet('Exchange')]
		[string[]]$Service,
		[Parameter(Mandatory = $False, ParameterSetName = 'MFA')]
		[Switch]$MFA
	)

	If ($MFA -ne $True) {
		Write-Verbose "Gathering PSCredentials object for non MFA sign on"
		$Credential = Get-Credential -Message "Please enter your Office 365 credentials"
	}

	ForEach ($Item in $PSBoundParameters.Service) {
		Write-Verbose "Attempting connection to $Item"
		Switch ($Item) {
			Exchange {
				If ($MFA -eq $True) {
					$getChildItemSplat = @{
						Path = "$Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1"
						Recurse = $true
						ErrorAction = 'SilentlyContinue'
						Verbose = $false
					}
					$MFAExchangeModule = ((Get-ChildItem @getChildItemSplat | Select-Object -ExpandProperty Target -First 1).Replace("CreateExoPSSession.ps1", ""))
					
					If ($null -eq $MFAExchangeModule) {
						Write-Error "The Exchange Online MFA Module was not found! https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
						continue
					}
					Else {
						Write-Verbose "Importing Exchange MFA Module"
						. "$MFAExchangeModule\CreateExoPSSession.ps1"
						
						Write-Verbose "Connecting to Exchange Online"
						Connect-EXOPSSession
						If ($Null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" })) {
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*") {
								$host.ui.RawUI.WindowTitle += " - Connected To: Exchange"
							}
							Else {
								$host.ui.RawUI.WindowTitle += " - Exchange"
							}
						}
					}
				}
				Else {
					$newPSSessionSplat = @{
						ConfigurationName = 'Microsoft.Exchange'
						ConnectionUri	  = "https://ps.outlook.com/powershell/"
						Authentication    = 'Basic'
						Credential	      = $Credential
						AllowRedirection  = $true
					}
					$Session = New-PSSession @newPSSessionSplat
					Write-Verbose "Connecting to Exchange Online"
					Import-PSSession $Session -AllowClobber
					If ($Null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" })) {
						If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*") {
							$host.ui.RawUI.WindowTitle += " - Connected To: Exchange"
						}
						Else {
							$host.ui.RawUI.WindowTitle += " - Exchange"
						}
					}
				}
				continue
			}
			Default { }
		}
	}
}

### Input Parameters and Transforms #############
Param (
    $MFA = "NO"
)
$Script:MFA = $MFA.ToUpper()

### Apply The Signatures ########################
Function OWASigs {
	# Check for Office 365 connetion and connect
	$ConnectToO365 = Get-PSSession | Where-Object {$_.ConfigurationName -eq "Microsoft.exchange"}
	If ($ConnectToO365 -eq $Null -or $ConnectToO365.State -eq "Closed") {
		# If you do not require MFA for Office 365 remove the -MFA switch below
		If ($MFA -eq "YES") {
			Connect-Office365 -Service Exchange -MFA
		}
		Else {
			Connect-Office365 -Service Exchange
		}
	}

	$SignatureTemplate = Get-Content $SignatureFile

	ForEach ($User in $Users) {
		$UPN = $User.UserName
		$UserDetails = Get-ADUser -Filter * -Properties * | Where-Object {$_.SamAccountName -eq $UPN} | Select-Object name,title,emailaddress,telephoneNumber,mobile,userPrincipalName,awsChime

		ForEach ($SignatureUser in $UserDetails) {
			#Generated Signature
			$Signature = @()
			$UserFullName = $SignatureUser.Name
			$UserTitle = $SignatureUser.Title
			If ($SignatureUser.telephoneNumber) { $OfficePhone = $SignatureUser.telephoneNumber } Else { $OfficePhone = $Null }
			If ($SignatureUser.mobile) { $Mobile = $SignatureUser.mobile } Else { $Mobile = $Null }
			# Requires custom AD attribute called 'awsChime' in the user attributes
			If ($SignatureUser.awsChime) { $AwsChime = $SignatureUser.awsChime } Else { $AwsChime = $Null }
			$UserEmail = $SignatureUser.emailAddress

			$Signature += "<br>"
			$Signature += "<hr size=`"3`" width=`"33%`" align=`"left`">"
			$Signature += "<div style=`"font-size:10pt;  font-family: 'Calibri',sans-serif;`">"
			$Signature += "<b><span style='font-size:14.0pt;color:#D34727'>$UserFullName</span></b><br>"
			$Signature += "<span style='text-transform:uppercase'>$UserTitle</span><br>"
			$Signature += "<div><img alt=`"$CompanyName`"  src=`"$CompanyLogo`"></div>"
			If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null)) {
				$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Office - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$OfficePhone</span>"
				}
			If (($OfficePhone -ne "") -AND ($OfficePhone -ne $Null) -AND ($Mobile -ne "") -AND ($Mobile -ne $Null)) {
				$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
				}
			If (($Mobile -ne "") -AND ($Mobile -ne $Null)) {
				$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Cell - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$Mobile</span></span>"
				}
			If (($AwsChime -ne "") -AND ($AwsChime -ne $Null)) {
				$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'> | </span></i>"
				$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>Cell - </span></i><span style='font-size:10.0pt'><span style='color:#000001; text-decoration:none;text-underline:none'>$AwsChime</span></span>"
				}
			$Signature += "<br><span style='font-size:10.0pt;text-transform:lowercase'>$UserEmail | <a href=`"$SocialMediaURL`"><span style='text-transform:uppercase;color:#000001'>$SocialMediaTitle</span></a> | <a href=`"$WebsiteURL`"><b>$WebsiteDomain</b></a></span><br>"
			$Signature += "<br>"
			$Signature += "<i><span style='font-size:8.0pt;font-family:`"Georgia`",serif'>$Motto</span></i> <o:p></o:p></p>"
			$Signature += "</div>"
			}
			
			Set-MailboxMessageConfiguration $SignatureUser.userPrincipalName -SignatureHtml $Signature -AutoAddSignature $True -AutoAddSignatureOnMobile $True -AutoAddSignatureOnReply $True -DefaultFontColor "#000000" -UseDefaultSignatureOnMobile $True
			Write-Host "Signatures should be created for $($SignatureUser.userPrincipalName). Please check OWA to verify." -ForegroundColor Green
		}
	}

	Get-PSSession | Remove-PSSession
}

OWASigs
