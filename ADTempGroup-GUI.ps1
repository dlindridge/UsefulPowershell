#################################################
<#
	Scans for any domains in your Forest and creates a group in a specified domain that will expire and delete at
	a set time. Server2019+ AD Domains support natively temporary groups to allow users to get access to resources
	they won't need long term. Server 2008R2+ AD Domains allow dynamic groups with a TTL, but do not provide and
	easy way to create them. This script bridges that gap.
	
	The script will scan for your Default User Container (where user accounts are created if you script user creation
	without specifying a destination path) and create the temporary group in that location. This script will also 
	verify that the chosen group name is not otherwise in use in the domain.
	
	WARNING: Dynamic Groups with a TTL *SKIP* the tombstone and recovery phase. Once they expire, they are *GONE*.
	
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: January 2, 2021
    Modified: February 25,2021
#>
#################################################


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


### Create Base Form ############################
Function MakeForm {
    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Create a Temporary AD Group"
    $Form.Font = New-Object System.Drawing.Font("Times New Roman",10,[System.Drawing.FontStyle]::Bold)
    $Form.AutoSize = $True
    $Form.AutoSizeMode = "GrowOnly" # GrowAndShrink, GrowOnly
    $Form.MaximizeBox = $False
    $Form.AutoScroll = $False
    $Form.WindowState = "Normal" # Maximized, Minimized, Normal
    $Form.SizeGripStyle = "Hide" # Auto, Hide, Show
    $Form.StartPosition = "WindowsDefaultLocation" # CenterScreen, Manual, WindowsDefaultLocation, WindowsDefaultBounds, CenterParent
    $Form.BackgroundImageLayout = "Tile" # None, Tile, Center, Stretch, Zoom
    $Form.BackColor = "lightgray"
	$ObjFont = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)
    $ObjFontBold = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)

    $GroupNameTextBoxLabel = New-Object System.Windows.Forms.Label
    $GroupNameTextBoxLabel.Location = New-Object System.Drawing.Size(2,10)
    $GroupNameTextBoxLabel.Size = New-Object System.Drawing.Size(90,25)
    $GroupNameTextBoxLabel.TextAlign = "MiddleLeft"
    $GroupNameTextBoxLabel.Text = "Group Name:"
    $GroupNameTextBoxLabel.BackColor = "Transparent"

    $GroupNameTextBox = New-Object System.Windows.Forms.TextBox
    $GroupNameTextBox.Location = New-Object System.Drawing.Size(100,10)
    $GroupNameTextBox.Size = New-Object System.Drawing.Size(222,25)
    $GroupNameTextBox.Font = $ObjFont
	
    $adDomainsDropDownBoxLabel = New-Object System.Windows.Forms.Label
    $adDomainsDropDownBoxLabel.Location = New-Object System.Drawing.Size(2,45)
    $adDomainsDropDownBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $adDomainsDropDownBoxLabel.TextAlign = "MiddleLeft"
    $adDomainsDropDownBoxLabel.Text = "Domain:"
    $adDomainsDropDownBoxLabel.BackColor = "Transparent"

    [array]$adDomains = (Get-ADForest).domains
    $adDomainsDropDownBox = New-Object System.Windows.Forms.ComboBox
    $adDomainsDropDownBox.Location = New-Object System.Drawing.Size(80,45)
    $adDomainsDropDownBox.Size = New-Object System.Drawing.Size(243,25)
    $adDomainsDropDownBox.Font = $ObjFont
    $adDomainsDropDownBox.DropDownStyle = "DropDownList"
    ForEach ($adDomain in $adDomains) { $adDomainsDropDownBox.Items.Add($adDomain) }
    $adDomainsDropDownBox.SelectedIndex = 0

    $adDomainsSelectButton = New-Object System.Windows.Forms.Button 
    $adDomainsSelectButton.BackColor = "white" 
    $adDomainsSelectButton.Location = New-Object System.Drawing.Size(330,8)
    $adDomainsSelectButton.Size = New-Object System.Drawing.Size(110,60)
    $adDomainsSelectButton.Text = "Check Groupname and Domain"

	$adGroupListTextBox = New-Object System.Windows.Forms.Label
	$adGroupListTextBox.Location = New-Object System.Drawing.Size(2,80)
	$adGroupListTextBox.Size = New-Object System.Drawing.Size(200,25)
	$adGroupListTextBox.Text = "Group Is A Member Of:"

    $adGroupListBox = New-Object System.Windows.Forms.ListBox 
    $adGroupListBox.Location = New-Object System.Drawing.Size(2,105) 
    $adGroupListBox.Size = New-Object System.Drawing.Size(200,420) 
    $adGroupListBox.Font = $ObjFont
    $adGroupListBox.Sorted = $True
	$adGroupListBox.Enabled = $False

	$userListTextBox = New-Object System.Windows.Forms.Label
	$userListTextBox.Location = New-Object System.Drawing.Size(230,80)
	$userListTextBox.Size = New-Object System.Drawing.Size(200,25)
	$userListTextBox.Text = "Users Part Of This Group:"

    $userListBox = New-Object System.Windows.Forms.ListBox 
    $userListBox.Location = New-Object System.Drawing.Size(230,105) 
    $userListBox.Size = New-Object System.Drawing.Size(200,420) 
    $userListBox.Font = $ObjFont
    $userListBox.Sorted = $True
	$userListBox.Enabled = $False

	$timeTextBox = New-Object System.Windows.Forms.Label
	$timeTextBox.Location = New-Object System.Drawing.Size(2,540)
	$timeTextBox.Size = New-Object System.Drawing.Size(100,25)
	$timeTextBox.Text = "Time To Live:"

    $timeDropDownBox = New-Object System.Windows.Forms.ComboBox 
    $timeDropDownBox.Location = New-Object System.Drawing.Size(110,540) 
    $timeDropDownBox.Size = New-Object System.Drawing.Size(200,25) 
	[array]$TTLs = @('15 Minutes','30 Minutes','1 Hour','90 Minutes','2 Hours','4 Hours','6 Hours','8 Hours','1 Day','2 Days','3 Days','4 Days','5 Days','6 Days','7 Days','14 Days','30 Days','90 Days')
	ForEach ($TTL in $TTLs) { $timeDropDownBox.Items.Add($TTL) }
    $timeDropDownBox.Font = $ObjFont
    $timeDropDownBox.SelectedIndex = 0
	$timeDropDownBox.Enabled = $False

    $createGroupButton = New-Object System.Windows.Forms.Button 
    $createGroupButton.BackColor = "white" 
    $createGroupButton.Location = New-Object System.Drawing.Size(330,540)
    $createGroupButton.Size = New-Object System.Drawing.Size(110,40)
    $createGroupButton.Text = "Create Group"
	$createGroupButton.Enabled = $False


    ### Add controls to the form
	$Form.Controls.Add($GroupNameTextBoxLabel)
	$Form.Controls.Add($GroupNameTextBox)
	$Form.Controls.Add($adDomainsDropDownBox)
    $Form.Controls.Add($adDomainsDropDownBoxLabel)
    $Form.Controls.Add($adDomainsDropDownBox)
    $Form.Controls.Add($adDomainsSelectButton)
    $Form.Controls.Add($adGroupListTextBox)
    $Form.Controls.Add($adGroupListBox)
	$Form.Controls.Add($userListTextBox)
	$Form.Controls.Add($userListBox)
	$Form.Controls.Add($timeTextBox)
	$Form.Controls.Add($timeDropDownBox)
	$Form.Controls.Add($createGroupButton)


    ### Populate the group & user selection box
    $adDomainsSelectButton.Add_Click({
        $Script:Domain = $adDomainsDropDownBox.SelectedItem.ToString()
		$Script:ProposedName = $GroupNameTextBox.Text
		GroupNameValidation
	})
	
	
	### Validate Group Name
	Function GroupNameValidation {
		If ($ProposedName -eq " " -OR $ProposedName -eq "" -OR $ProposedName -eq $Null) {
			Add-Type -AssemblyName PresentationCore,PresentationFramework
			$ButtonType = [System.Windows.MessageBoxButton]::OK
			$MessageboxTitle = "Group Name Error"
			$MessageboxBody = "Please pick a valid Group Name. This field cannot be blank."
			$MessageIcon = [System.Windows.MessageBoxImage]::Error
			[System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
			ClearForm
		}
		$CheckGroupName = Get-ADGroup -Server $Domain -LDAPFilter "(SamAccountName=$ProposedName)"
		If ($CheckGroupName -eq $Null) { PopulateGroupInformation }
		Else {
			Add-Type -AssemblyName PresentationCore,PresentationFramework
			$ButtonType = [System.Windows.MessageBoxButton]::OK
			$MessageboxTitle = "Group Name Error"
			$MessageboxBody = "Group Name '$ProposedName' already exists in the $Domain domain. Please try again or contact your administrator for help."
			$MessageIcon = [System.Windows.MessageBoxImage]::Error
			[System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
			ClearForm
		}
	}

	
	### Populate Group Information
	Function PopulateGroupInformation {
		$DN = Get-ADDomain -Server $Domain | Select-Object DistinguishedName
		$BuiltinOU = "*CN=Builtin,"
		$UsersOU = "*CN=Users,"
		$BuiltinOU = $BuiltinOU + $DN.DistinguishedName
		$UsersOU = $UsersOU + $DN.DistinguishedName
		$BuiltinOU = $BuiltinOU.ToString()
		$UsersOU = $UsersOU.ToString()
		$adGroupItems = [System.Collections.Generic.List[object]](Get-ADGroup -Server $Domain -Filter * | Where-Object {($_.DistinguishedName -NotLike $BuiltinOU)} | Where-Object {($_.DistinguishedName -NotLike $UsersOU)} | Select-Object Name,sAMAccountName | Sort-Object Name)
		$adUserItems = [System.Collections.Generic.List[object]](Get-ADUser -Server $Domain -Filter * | Where-Object {($_.DistinguishedName -NotLike $BuiltinOU)} | Where-Object {($_.DistinguishedName -NotLike $UsersOU)} | Select-Object Name,sAMAccountName | Sort-Object Name)
		$adGroupListBox.SelectionMode = "MultiExtended" # MultiExtended, MultiSimple, None, One
		$userListBox.SelectionMode = "MultiExtended" # MultiExtended, MultiSimple, None, One
		If ($adGroupItems -ne "" -OR $adGroupItems -ne $Null) {
			$adGroupListBox.ValueMember = "SamAccountName"
			$adGroupListBox.DisplayMember = "Name"
			$adGroupListBox.DataSource = $adGroupItems
		}
		If ($adUserItems -ne "" -OR $adUserItems -ne $Null) {
			$userListBox.ValueMember = "SamAccountName"
			$userListBox.DisplayMember = "Name"
			$userListBox.DataSource = $adUserItems
		}
		$adGroupListBox.TopIndex = 0
		$adGroupListBox.Enabled = $True
		$userListBox.TopIndex = 0
		$userListBox.Enabled = $True
		$timeDropDownBox.Enabled = $True
		$adDomainsSelectButton.Enabled = $False
		$createGroupButton.Enabled = $True
	}
	
	
	### Clear Form
	Function ClearForm {
		$Form.Close()
		$Form.Dispose()
		MakeForm
	}
	
	
	### Create Group
	$createGroupButton.Add_Click({
		$TimeToLive = $timeDropDownBox.SelectedItem.ToString()
		$Groups = ($adGroupListBox.SelectedItems).SamAccountName
		$Users = ($userListBox.SelectedItems).SamAccountName
		Switch ($TimeToLive) {
            { $TimeToLive -eq '15 Minutes' } { $TTLMinutes = 15 }
            { $TimeToLive -eq '30 Minutes' } { $TTLMinutes = 30 }
            { $TimeToLive -eq '1 Hour' } { $TTLMinutes = 60 }
            { $TimeToLive -eq '90 Minutes' } { $TTLMinutes = 90 }
            { $TimeToLive -eq '2 Hours' } { $TTLMinutes = 2*60 }
            { $TimeToLive -eq '4 Hours' } { $TTLMinutes = 4*60 }
            { $TimeToLive -eq '6 Hours' } { $TTLMinutes = 6*60 }
            { $TimeToLive -eq '8 Hours' } { $TTLMinutes = 8*60 }
            { $TimeToLive -eq '1 Day' } { $TTLMinutes = 24*60 }
            { $TimeToLive -eq '2 Days' } { $TTLMinutes = 24*60*2 }
            { $TimeToLive -eq '3 Days' } { $TTLMinutes = 24*60*3 }
            { $TimeToLive -eq '4 Days' } { $TTLMinutes = 24*60*4 }
            { $TimeToLive -eq '5 Days' } { $TTLMinutes = 24*60*5 }
            { $TimeToLive -eq '6 Days' } { $TTLMinutes = 24*60*6 }
            { $TimeToLive -eq '7 Days' } { $TTLMinutes = 24*60*7 }
            { $TimeToLive -eq '14 Days' } { $TTLMinutes = 24*60*14 }
            { $TimeToLive -eq '30 Days' } { $TTLMinutes = 24*60*30 }
            { $TimeToLive -eq '90 Days' } { $TTLMinutes = 24*60*90 }
		}
		$GroupName = $GroupNameTextBox.Text
		$TTLSeconds = $TTLMinutes * 60
		$EndTime = (Get-Date).AddMinutes($TTLMinutes)
		$RemoveDate = $EndTime.ToShortDateString()
		$RemoveTime = $EndTime.ToShortTimeString()
		$OU = Get-ADDomain -Server $Domain | Select UsersContainer
		$LDAP = "LDAP://" + $OU.UsersContainer
		$destinationOuObject = [ADSI]($LDAP)
		$TempGroup = $destinationOuObject.Create("group","CN=$GroupName")
		$TempGroup.PutEx(2,"objectClass",@("dynamicObject","Group"))
		$TempGroup.Put("entryTTL",$TTLSeconds)
		$TempGroup.Put("sAMAccountName", $GroupName)
		$TempGroup.Put("displayName", $GroupName)
		$TempGroup.Put("description",("Temp Group, will be deleted automatically on $RemoveDate at $RemoveTime"))
		$TempGroup.SetInfo()

		ForEach ($Group in $Groups) {
			Add-ADGroupMember -Server $Domain -Identity $Group -Members $GroupName
		}
		
		ForEach ($User in $Users) {
			Add-ADGroupMember -Server $Domain -Identity $GroupName -Members $User
		}
		
		Add-Type -AssemblyName PresentationCore,PresentationFramework
			$ButtonType = [System.Windows.MessageBoxButton]::OK
			$MessageboxTitle = "Group Created"
			$MessageboxBody = "Group Name '$GroupName' created. Users '$Users' added as memebers."
			$MessageIcon = [System.Windows.MessageBoxImage]::Information
			[System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
			ClearForm
	})

    ### Launch Form
    $Form.ShowDialog()

}

MakeForm
