<#
    .SYNOPSIS
    Using an HTML template for an email signature, customize it for each user in a specified source CSV to use as a
	signature block with OWA. Pulls data directly from Active Directory and connects to Office 365/Exchange Online
	to apply the signature and set it as default.

	.PARAMETER MFA
	Is MFA authentication required for your connection to Office 365/Exchange Online? Default: NO

    .DESCRIPTION
    Usage: .\OWASigUpdate.ps1 (-MFA YES).
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

	The username in the CSV file should be the user's SamAccountName in AD with the column heading "UserName"

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 30, 2019
    Modified: October 31, 2019
#>
#################################################

### Build some organization specific variables
$Script:SignatureFile = "\\MyDomain.org\NETLOGON\EmailSig_Main.html" # Name and Location of signature template
$Script:Users = Import-Csv "C:\Scripts\ADUserInfo\ADUserInfo.csv" # Name and Location of user CSV list


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
		$UserDetails = Get-ADUser -Filter * -Properties * | Where-Object {$_.SamAccountName -eq $UPN} | Select-Object name,title,emailaddress,telephoneNumber,mobile,userPrincipalName

		ForEach ($SignatureUser in $UserDetails) {
			#Generated Signature
			$Signature = @()
			If ($SignatureUser.mobile) { $Mobile = $SignatureUser.mobile } Else { $Mobile = $Null }

			# Find the PS-*-* fields and replace those with the appropriate details.
			ForEach ($line in $SignatureTemplate) {
				If ($line -like "*PS-USER-NAME*") {
					$Signature += $line.Replace("PS-USER-NAME","$($SignatureUser.Name)")
				}
				ElseIf ($line -like "*PS-TITLE-NAME*") {
					$Signature += $line.Replace("PS-TITLE-NAME","$($SignatureUser.Title)")
				}
				ElseIf ($line -like "*PS-OFFICE-PHONE*") {
					$Signature += $line.Replace("PS-OFFICE-PHONE","$($SignatureUser.telephoneNumber)")
				}
				ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -ne "" -AND $Mobile -ne $Null)) {
					$Signature += $line.Replace("PS-MOBILE-PHONE",$Mobile)
				}
				ElseIf ($line -like "*PS-MOBILE-PHONE*" -AND ($Mobile -eq "" -OR $Mobile -eq $Null)) {
					Write-Output "$($SignatureUser.Name) - No Mobile Phone Number"
				}
				ElseIf ($line -like "*PS-EMAIL-ADDR*") {
					$Signature += $line.Replace("PS-EMAIL-ADDR","$($SignatureUser.emailaddress)")
				}
				Else { $Signature += $line }
			}
			Set-MailboxMessageConfiguration $SignatureUser.userPrincipalName -SignatureHtml $Signature -AutoAddSignature $True -AutoAddSignatureOnMobile $True -AutoAddSignatureOnReply $True -DefaultFontColor "#000000" -UseDefaultSignatureOnMobile $True
			Write-Host "Signatures should be created for $($SignatureUser.userPrincipalName). Please check OWA to verify." -ForegroundColor Green
		}
	}

	Get-PSSession | Remove-PSSession
}

OWASigs
