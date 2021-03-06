<#
    .SYNOPSIS
    Lock an AD account by attempting to login with/use it one more than the Domain Lockout Policy allows.

    .PARAMETER SamName
    SamAccountName of the user to lock

    .DESCRIPTION
    Usage: .\LockADAccount.ps1 -SamName john.doe

#>
#################################################
<#
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    https://github.com/dlindridge/UsefulPowershell
    Created: October 23, 2019
    Modified: October 23, 2019
#>
#################################################

Param (
    [Parameter(Mandatory=$True)]
    $SamName
)

### Random Password Generator ###################
Function Get-RandomCharacters($Length) { 
	$characters = "1234567890abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ"
    $random = 1..$Length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    Return [String]$characters[$random]
}

### Check If Account Locked #####################
$AccountLocked = Get-ADUser $SamName -Properties * | Select-Object LockedOut
If ($AccountLocked.LockedOut -eq $True) {
    Write-Output "$SamName is already locked. No action necessary"
    Break
}

### Lock The Account ############################
If ($LockoutBadCount = ((([xml](Get-GPOReport -Name "Default Domain Policy" -ReportType Xml)).GPO.Computer.ExtensionData.Extension.Account | Where-Object name -eq LockoutBadCount).SettingNumber)) {
    $RandomPwd = Get-RandomCharacters -Length 33
    $Password = ConvertTo-SecureString $RandomPwd -AsPlainText -Force
    $PickDC = Get-ADDomainController -filter * | Select-Object -First 1 | Select-Object Name
    Get-ADUser $SamName -Properties SamAccountName, UserPrincipalName, LockedOut | ForEach-Object {
        for ($i = 0; $i -le $LockoutBadCount; $i++) { Invoke-Command -ComputerName $PickDC.Name { Get-Process } -Credential (New-Object System.Management.Automation.PSCredential ($($_.UserPrincipalName), $Password)) -ErrorAction SilentlyContinue }
        Write-Output "$($_.SamAccountName) has been locked out: $((Get-ADUser -Identity $_.SamAccountName -Properties LockedOut).LockedOut)"
    }
}
