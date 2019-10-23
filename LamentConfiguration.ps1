<#
    .SYNOPSIS
    Lead Cenobite: The box, you opened it, we came.
    Kirsty Cotton: [Kirsty screams] It's just a puzzle box!
    Lead Cenobite: Oh no, it is a means to summon us.
    Lead Cenobite: We have such sights to show you!

    .DESCRIPTION
    Queries your domain to find locked users and the Domain Controller that registered the most recent logon (this should in theory put it
    closest to the user) then presents you with options to unlock the account.
#>
#################################################
<#
    TO-DO:
    Let you choose the DC to unlock on.
    Provide the option to reset the password at the same time.
	
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 23, 2019
    Modified: October 23, 2019
#>
#################################################

$Path = "C:\Scripts\Unlock" # Path to image files (probably the same as the script location)
$IconImage = "LC_symbol.ico" # Filename for the application icon - blank for none ("")
$BackgroundImage = "LC_background.jpg" # Filename for the background image - blank for none ("")

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

HideConsole


### Reset Form ##################################
Function MakeNewForm {
    $Form.Close()
    $Form.Dispose()
    MakeForm
}


### Cancel Form #################################
Function CancelForm{
    $Form.Close()
    $Form.Dispose()
}


### Build Form ##################################
Function MakeForm {
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
    $ObjFontBold = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)

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
     
    $SelectDomainButton = New-Object System.Windows.Forms.Button 
    $SelectDomainButton.Location = New-Object System.Drawing.Size(310,10)
    $SelectDomainButton.Size = New-Object System.Drawing.Size(111,25)
    $SelectDomainButton.Text = "Select Domain"
    $Form.Controls.Add($SelectDomainButton)

    $CancelButton1 = New-Object System.Windows.Forms.Button 
    $CancelButton1.Location = New-Object System.Drawing.Size(440,410)
    $CancelButton1.Size = New-Object System.Drawing.Size(180,50)
    $CancelButton1.Text = "Nevermind"
    $CancelButton1.Add_Click({ CancelForm })
    $Form.Controls.Add($CancelButton1)

    $UserListBox = New-Object System.Windows.Forms.ListBox 
    $UserListBox.Location = New-Object System.Drawing.Size(2,40) 
    $UserListBox.Size = New-Object System.Drawing.Size(420,420) 
    $UserListBox.Font = $ObjFont
    $UserListBox.Sorted = $True
    $UserListBox.Enabled = $False
    $UserListBox.SelectionMode = "MultiExtended" # MultiExtended, MultiSimple, None, One
    $Form.Controls.Add($UserListBox)

    $SelectUsersButton = New-Object System.Windows.Forms.Button 
    $SelectUsersButton.Location = New-Object System.Drawing.Size(440,195)
    $SelectUsersButton.Size = New-Object System.Drawing.Size(111,75)
    $SelectUsersButton.Text = "Unlock Users"
    $SelectUsersButton.Enabled = $False
    $Form.Controls.Add($SelectUsersButton)


    $SelectDomainButton.Add_Click({
        $Script:Domain = $adDomainsDropDownBox.SelectedItem
        $UserListBoxItems = $UserListBox.Items
        while ($UserListBoxItems -ne $Null) {
            $UserListBox.Items.Remove($UserListBoxItems[0])
            $UserListBoxItems = $UserListBox.Items
        }
        $UserListBox.Enabled = $True
        GetUsers
    })
    ### Launch Form
    $Form.ShowDialog()
}


### Get Locked Users ############################
Function GetUsers {
    $DN = Get-ADDomain -Server $Domain | Select-Object DistinguishedName
    $BuiltinOU = "*CN=Builtin,"
    $UsersOU = "*CN=Users,"
    $BuiltinOU = $BuiltinOU + $DN.DistinguishedName
    $UsersOU = $UsersOU + $DN.DistinguishedName
    $BuiltinOU = $BuiltinOU.ToString()
    $UsersOU = $UsersOU.ToString()

    $LockedUsersSearch = Get-ADDomainController -Server $Domain -Filter * | ForEach-Object { Get-ADUser -Server $Domain -Filter * -Properties * | Where-Object {($_.DistinguishedName -NotLike $BuiltinOU)} | Where-Object {($_.DistinguishedName -NotLike $UsersOU)} | Where-Object {($_.LockedOut -eq $True)}} | Select-Object Name,SamAccountName
    $LockedUsers = ForEach ($searchEntry in $LockedUsersSearch | Group-Object SamAccountName){ $searchEntry.Group | Sort-Object -Property LastLogonTimestamp -Descending | Select-Object -First 1 }
    If ($LockedUsers -eq $Null -OR $LockedUsers -eq "") {
        $UserListBox.Items.Add("No locked users in the $Domain domain.")
        $SelectUsersButton.Enabled = $False
    }
    Else {
        $FoundUsers = [System.Collections.Generic.List[object]]($LockedUsers)
        $UserListBox.DataSource = $FoundUsers
        $UserListBox.ValueMember = "SamAccountName"
        $UserListBox.DisplayMember = "Name"
        $SelectUsersButton.Enabled = $True
    }

    $SelectUsersButton.Add_Click({
        $SelectedUsers = $UserListBox.SelectedItems
        UnlockUsers
    })
}


### Unlock Users ################################
Function UnlockUsers {
    ForEach ($User in $SelectedUsers) {
        Unlock-ADAccount -Server $Domain -Identity $User.SamAccountName
    }
    Add-Type -AssemblyName PresentationCore,PresentationFramework
    $ButtonType = [System.Windows.MessageBoxButton]::OK
    $MessageboxTitle = "User Unlocked"
    $MessageboxBody = "The following users have been unlocked: $($SelectedUsers.Name)"
    $MessageIcon = [System.Windows.MessageBoxImage]::Information
    [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
    MakeNewForm
}


MakeForm
