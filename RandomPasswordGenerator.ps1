#################################################
<# 
	Gets all enabled accounts in the domain (at the Search OU Root) and
	generates a random 15 character password that meets Microsoft AD
	password requirements. Generates and output CSV with the passwords
	assigned to each user.
 #>
#################################################

$domain = "MyDomain.TLD"
$searchRoot = "OU=Root,DC=MyDomain,DC=TLD"
$csvPath = "C:\Output\Path"

### Gets all enabled & expiring users
$userPath = Join-Path -Path $csvPath -ChildPath "users.csv"
$starters = (Get-ADUser -Server $domain -SearchBase $searchRoot -Properties * -Filter {(Enabled -eq $True -AND PasswordNeverExpires -eq $False)} | Select SamAccountName)
ForEach ($starter in $starters) { Add-Content $userPath $starter.SamAccountName | Wait-Job }

### Function for creating a random password
Function Get-RandomCharacters($length) { 
	$characters = "1234567890abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ"
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    Return [String]$characters[$random]
}

### Creates random password for users in first CSV and creates a new combined CSV
$users = Import-CSV $userPath
$outPath = Join-Path -Path $csvPath -ChildPath "newlist.csv"
Add-Content $outPath "User,Pass" | Wait-Job
ForEach ($user in $users) {
    $password = (Get-RandomCharacters -length 15)
    Set-ADAccountPassword -Identity $user -NewPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Reset
	Set-ADUser -Identity $user -ChangePasswordAtLogon $True
	Add-Content $outPath "$user,$password" | Wait-Job
}
