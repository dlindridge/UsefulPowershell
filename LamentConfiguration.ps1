<#
    .SYNOPSIS
    Lead Cenobite: The box, you opened it, we came.
    Kirsty Cotton: [Kirsty screams] It's just a puzzle box!
    Lead Cenobite: Oh no, it is a means to summon us.
    Lead Cenobite: We have such sights to show you!

    .DESCRIPTION
    Queries your domain to find locked users and the Domain Controller that registered the most recent logon (this should in theory
    put it closest to the user) then presents you with options to unlock the account and optionally reset the password.
    Usage: Create a shortcut pointing to (powershell.exe -file "C:\Scripts\UserStuff\LamentConfiguration.ps1")
#>
#################################################
<#
    TO-DO:
    Progress Bar on user lookup.
    Let you choose the DC to unlock on.

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 23, 2019
    Modified: October 24, 2019
#>
#################################################

$Path = "C:\Users\dlindridge\Desktop\Unlock" # Path to image files (probably the same as the script location)
$IconImage = "symbol.ico" # Filename for the application icon - blank for none ("")
$BackgroundImage = "background.jpg" # Filename for the background image - blank for none ("")

### Show or Hide Console ########################
Add-Type -Name Window -Namespace Console -MemberDefinition '
[DllImport("Kernel32.dll")]
public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")]
public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);
'

Function ShowConsole {
    $PSConsole = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($PSConsole, 5)
 }
 
 Function HideConsole {
    $PSConsole = [Console.Window]::GetConsoleWindow()
    [Console.Window]::ShowWindow($PSConsole, 0)
 }

ShowConsole


### Cancel Form #################################
Function CancelForm{
    $Form.Close()
    $Form.Dispose()
}


### Build Form ##################################
Function MakeForm {
    ### Set form parameters
    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Create a New Active Directory User"
    $Form.Font = New-Object System.Drawing.Font("Times New Roman",10,[System.Drawing.FontStyle]::Bold)
    $Form.AutoSize = $True
    $Form.AutoSizeMode = "GrowOnly" # GrowAndShrink, GrowOnly
    $Form.MaximizeBox = $False
    $Form.AutoScroll = $False
    $Form.WindowState = "Normal" # Maximized, Minimized, Normal
    $Form.SizeGripStyle = "Hide" # Auto, Hide, Show
    $Form.StartPosition = "WindowsDefaultLocation" # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
    $Form.BackgroundImageLayout = "Tile" # None, Tile, Center, Stretch, Zoom
    $Icon = Join-Path -Path $Path -ChildPath $IconImage
	$Background = Join-Path -Path $Path -ChildPath $BackgroundImage
    If ($IconImage -ne "" -OR $IconImage -ne $Null) { $Form.Icon = New-Object system.drawing.icon ($Icon) }
    If ($BackgroundImage -ne "" -OR $BackgroundImage -ne $Null) { $Form.BackgroundImage = [system.drawing.image]::FromFile($Background) }
    $ObjFont = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)
    $ObjFontBold = New-Object System.Drawing.Font("Microsoft Sans Serif",14,[System.Drawing.FontStyle]::Bold)

    ### Set and add form objects
    $adDomainsDropDownBoxLabel = New-Object System.Windows.Forms.Label
    $adDomainsDropDownBoxLabel.Location = New-Object System.Drawing.Size(2,10)
    $adDomainsDropDownBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $adDomainsDropDownBoxLabel.TextAlign = "MiddleLeft"
    $adDomainsDropDownBoxLabel.Text = "Domain:"
    $adDomainsDropDownBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($adDomainsDropDownBoxLabel)
    
    [array]$adDomains = (Get-ADForest).domains
    $adDomainsDropDownBox = New-Object System.Windows.Forms.ComboBox
    $adDomainsDropDownBox.Location = New-Object System.Drawing.Size(80,10)
    $adDomainsDropDownBox.Size = New-Object System.Drawing.Size(222,25)
    $adDomainsDropDownBox.Font = $ObjFont
    $adDomainsDropDownBox.DropDownStyle = "DropDownList"
    ForEach ($adDomain in $adDomains) { $adDomainsDropDownBox.Items.Add($adDomain) }
    $adDomainsDropDownBox.SelectedIndex = 0
    $Form.Controls.Add($adDomainsDropDownBox)

    $UserListBox = New-Object System.Windows.Forms.ListBox 
    $UserListBox.Location = New-Object System.Drawing.Size(2,40) 
    $UserListBox.Size = New-Object System.Drawing.Size(420,420) 
    $UserListBox.Font = $ObjFont
    $UserListBox.Sorted = $True
    $UserListBox.Enabled = $False
    $UserListBox.SelectionMode = "One" # MultiExtended, MultiSimple, None, One
    $Form.Controls.Add($UserListBox)

    $SelectDomainButton = New-Object System.Windows.Forms.Button 
    $SelectDomainButton.Location = New-Object System.Drawing.Size(310,10)
    $SelectDomainButton.Size = New-Object System.Drawing.Size(111,25)
    $SelectDomainButton.Text = "Select Domain"
    $Form.Controls.Add($SelectDomainButton)

    $UnlockUsersButton = New-Object System.Windows.Forms.Button 
    $UnlockUsersButton.Location = New-Object System.Drawing.Size(440,150)
    $UnlockUsersButton.Size = New-Object System.Drawing.Size(111,75)
    $UnlockUsersButton.Text = "Unlock Account Only"
    $UnlockUsersButton.Enabled = $False
    $Form.Controls.Add($UnlockUsersButton)

    $ResetUsersButton = New-Object System.Windows.Forms.Button 
    $ResetUsersButton.Location = New-Object System.Drawing.Size(440,240)
    $ResetUsersButton.Size = New-Object System.Drawing.Size(111,75)
    $ResetUsersButton.Text = "Reset Password and Unlock"
    $ResetUsersButton.Enabled = $False
    $Form.Controls.Add($ResetUsersButton)

    $CancelButton = New-Object System.Windows.Forms.Button 
    $CancelButton.Location = New-Object System.Drawing.Size(440,410)
    $CancelButton.Size = New-Object System.Drawing.Size(180,50)
    $CancelButton.Text = "Nevermind"
    $CancelButton.Add_Click({ CancelForm })
    $Form.Controls.Add($CancelButton)

    $PopulatingLabel = New-Object System.Windows.Forms.Label
    $PopulatingLabel.Location = New-Object System.Drawing.Size(430,20)
    $PopulatingLabel.Size = New-Object System.Drawing.Size(175,75)
    $PopulatingLabel.TextAlign = "MiddleLeft"
    $PopulatingLabel.Text = "Please Standby:`n Populating From"
    $PopulatingLabel.Font = $ObjFontBold
    $PopulatingLabel.ForeColor = "Red"
    $PopulatingLabel.BackColor = "Transparent"

    ### Action when Select Domain button is clicked
    $SelectDomainButton.Add_Click({
        $Form.Controls.Add($PopulatingLabel)
        $Script:Domain = $adDomainsDropDownBox.SelectedItem
        $UserListBox.DataSource = $null
        $UserListBox.Items.Clear()
        PopulateListbox
    })

    ### Action when Unlock User button is clicked
    $UnlockUsersButton.Add_Click({
        $UserListBox.Enabled = $False
        $SelectedUsers = $UserListBox.SelectedItems
        UnlockUsers
    })

    ### Action when Reset Password button is clicked
    $ResetUsersButton.Add_Click({
        $UserListBox.Enabled = $False
        $SelectedUsers = $UserListBox.SelectedItems
        ResetUsers
    })


    ### Launch Form
    $Form.ShowDialog()
}


### Populate Listbox ############################
Function PopulateListbox {
    ### Fill in list box with locked users (if any)
    $UserListBox.Enabled = $True
    $DN = Get-ADDomain -Server $Domain | Select-Object DistinguishedName
    ### But skip the Builtin and default Users OUs (because you shouldn't be using them for normal stuff!)
    $BuiltinOU = "*CN=Builtin,"
    $UsersOU = "*CN=Users,"
    $BuiltinOU = $BuiltinOU + $DN.DistinguishedName
    $UsersOU = $UsersOU + $DN.DistinguishedName
    $BuiltinOU = $BuiltinOU.ToString()
    $UsersOU = $UsersOU.ToString()
    $LockedUsersSearch = Get-ADDomainController -Server $Domain -Filter * | ForEach-Object { Get-ADUser -Server $Domain -Filter * -Properties * | Where-Object {($_.DistinguishedName -NotLike $BuiltinOU)} | Where-Object {($_.DistinguishedName -NotLike $UsersOU)} | Where-Object {($_.LockedOut -eq $True)}} | Select-Object Name,SamAccountName
    $LockedUsers = ForEach ($searchEntry in $LockedUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogonTimestamp -Descending | Select-Object -First 1 }
    $Form.Controls.Remove($PopulatingLabel)
    ### Are there locked users? If no, do the first
    If ($LockedUsers -eq $Null -OR $LockedUsers -eq "") {
        $UserListBox.Items.Add("No locked users in the $Domain domain.")
        $Form.Controls.Remove($PopulatingLabel)
        $UnlockUsersButton.Enabled = $False
        $ResetUsersButton.Enabled = $False
        $SelectDomainButton.Enabled = $True
    }
    ### Otherwise put them in the box
    Else {
        $FoundUsers = [System.Collections.Generic.List[object]]($LockedUsers)
        $UserListBox.DataSource = $FoundUsers
        $UserListBox.ValueMember = "SamAccountName"
        $UserListBox.DisplayMember = "Name"
        $Form.Controls.Remove($PopulatingLabel)
        $UnlockUsersButton.Enabled = $True
        $ResetUsersButton.Enabled = $True
        $SelectDomainButton.Enabled = $True
    }
}


### Random Password Generator ###################
Function Get-RandomCharacters($length) { 
	$characters = "1234567890abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ"
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    Return [String]$characters[$random]
}


### Unlock Users ################################
Function UnlockUsers {
    ### Unlock the user account then rescan for locked users
    ForEach ($User in $SelectedUsers) {
        Unlock-ADAccount -Server $Domain -Identity $User.SamAccountName
    }
    ### Popup window with the resulting action
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "User Unlocked"
    $MessageboxBody = "The user $($SelectedUsers.Name) has been unlocked"
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
    $Script:Domain = $adDomainsDropDownBox.SelectedItem
    ### Reset the form
    $Form.Controls.Add($PopulatingLabel)
    $UserListBox.DataSource = $null
    $UserListBox.Items.Clear()
    $Form.Controls.Add($PopulatingLabel)
    PopulateListbox
}


### Reset Users #################################
Function ResetUsers {
    ### Unlock the user account, reset the password, then rescan for locked users
    ForEach ($User in $SelectedUsers) {
        Unlock-ADAccount -Server $Domain -Identity $User.SamAccountName
        $Password = Get-RandomCharacters -length 15
        Set-ADAccountPassword -Server $Domain -Identity $user.SamAccountName -NewPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -Reset | Wait-Job
        Set-ADUser -Server $Domain -Identity $user.SamAccountName -ChangePasswordAtLogon $True | Wait-Job
    }
    ### Popup window with the resulting action
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "User Unlocked"
    $MessageboxBody = "The user $($SelectedUsers.Name) has been unlocked.`nTheir new password is $Password"
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
    ### Reset the form
    $Form.Controls.Add($PopulatingLabel)
    $UserListBox.DataSource = $null
    $UserListBox.Items.Clear()
    $Form.Controls.Add($PopulatingLabel)
    PopulateListbox

}

### Start Form As Function ######################
MakeForm
