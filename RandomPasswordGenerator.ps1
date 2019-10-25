#################################################
<# 
	Gets all enabled accounts in the domain (at the Search OU Root) and
	generates a random 15 character password that meets Microsoft AD
	password requirements. Generates and output CSV with the passwords
	assigned to each user.
 #>
#################################################

$domain = "MyDomain.TLD"
$searchRoot = "OU=UserRoot,DC=MyDomain,DC=TLD"
$csvPath = "C:\Path\To\Directory"

### Gets all enabled & expiring users
$userPath = Join-Path -Path $csvPath -ChildPath "users.csv"
$starters = (Get-ADUser -Server $domain -SearchBase $searchRoot -Properties * -Filter {(Enabled -eq $True -AND PasswordNeverExpires -eq $False)} | Select-Object SamAccountName)
ForEach ($starter in $starters) { Add-Content $userPath $starter.SamAccountName | Wait-Job }

### Function for creating a random password
Function Get-RandomCharacters($length) { 
	$characters = "1234567890abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ"
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    Return [String]$characters[$random]
}

### Opens the initial output for you to edit who gets reset
do {
	Write-Host -ForegroundColor Green "Imma gonna pause here for you to edit the output CSV to remove users to exclude before"
	Write-Host -ForegroundColor Green "proceeding. You should probably remove Domain and Enterprise Admins at the least..."
	Start-Process $userPath
	Write-Host -ForegroundColor Red -NoNewLine "Type 'GO' to continue: "
	$Continue = Read-Host
	$Continue = $Continue.ToUpper()
}
while ($Continue -ne "GO")

### Creates random password for users in first CSV and creates a new combined CSV
$users = Import-CSV $userPath
$outPath = Join-Path -Path $csvPath -ChildPath "newlist.csv"
Add-Content $outPath "User,Pass" | Wait-Job
ForEach ($user in $users) {
    $password = (Get-RandomCharacters -length 15)
    Set-ADAccountPassword -Identity $user -NewPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Reset | Wait-Job
	Set-ADUser -Identity $user -ChangePasswordAtLogon $True | Wait-Job
	Add-Content $outPath "$user,$password" | Wait-Job
}
