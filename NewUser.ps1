<#
    .SYNOPSIS
    Queries your AD Forest for all domains available then guides you through creating a new user in any one of them.
    Search for "### Facilities ###" to define your location addresses to use with the $Facilities array.

    .DESCRIPTION
    Usage: Create a shortcut pointing to (powershell.exe -file "C:\Scripts\NewUser\NewUser.ps1")
#>
#################################################
<#
    TO-DO:
    Mailbox Creation Selection - I use Office365 with AD Sync so at some point I want to figure out how to do a forced sync and assign a license with a checkbox
    Fix reload to clear all variable data and start fresh - for now it will close between each user
    Make Username and Display Name fields editable without crashing user validation and other bad things

    I make a few assumptions here so if your domain is different then go through the script and adjust the queries accordingly:
        1) Root OUs "Builtin" and "Users" are not used for any normal user functions (which Best Practices says not to do anyways!)
        2) Users will be created in your default user OU. Google how to change this if you haven't already.
        3) Admin accounts not in the OUs from #1 end with "*(Admin)" in the display name and are excluded from search.
        4) Service accounts not in the OUs from #1 start with "SERVICE*" in the display name and are excluded from search.
        5) Remote users will be assigned a blank address that your upload/update process from HR will fix with their home or mailing addresses.
        6) A password of 15 characters meeting Microsoft Complexity requirements will be randomly set for the new user. Search "$Password" and change the length value if you want it longer or shorter.
        7) Default domain groups in "Builtin" and "Users" are excluded by existing quries but given their elevated privledges I don't see that as a problem - they should be manually assigned if wanted.
        8) There is no possible way I was going to find every variation of phone number patterns. The ones for North American based numbers are correct but you should edit the inserts if you need another pattern.
        9) The country box is populated by querying Windows for all the unique countries it knows about and uses the 2-digit ISO code for each to assign to the user and encode the phone number format.
        10) I tried to make defaults useful anywhere but in the end they are specific to my needs. Change them if you want to, but it will be at your own risk...

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 7, 2019
    Modified: October 25, 2019
#>
#################################################

### Pre-Defining Select Variables ###############
[array]$Facilities = "Portland","Remote"
$Company = "MyCompany"
$Path = "C:\Scripts\UserStuff" # Path to image files (probably the same as the script location)
$IconImage = "symbol.ico" # Filename for the application icon - blank for none ("")
$BackgroundImage = "background.jpg" # Filename for the background image - blank for none ("")
$fromAddr = "MyCompany IT <noreply@MyDomain.TLD>" # Enter the FROM address for the e-mail alert - Must be inside quotes.
$toAddr = "itadmins@MyDomain.TLD" # Enter the default address for the e-mail alert if no manager selected - Must be inside quotes.
$smtpServer = "smtp.MyDomain.TLD" # Enter the FQDN or IP of a SMTP relay - Must be inside quotes.
$NameFormat = "FirstInit" # Username format - FirstInit, DottedFull, FullName
$AccountExpire = "Never" # Default account expiration interval - 30, 60, 90, Never

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
Function CancelForm {
    $Form.Close()
    $Form.Dispose()
}


### Create Base Form ############################
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

    $firstNameTextBoxLabel = New-Object System.Windows.Forms.Label
    $firstNameTextBoxLabel.Location = New-Object System.Drawing.Size(2,10)
    $firstNameTextBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $firstNameTextBoxLabel.TextAlign = "MiddleLeft"
    $firstNameTextBoxLabel.Text = "First Name:"
    $firstNameTextBoxLabel.BackColor = "Transparent"

    $firstNameTextBox = New-Object System.Windows.Forms.TextBox
    $firstNameTextBox.Location = New-Object System.Drawing.Size(80,10)
    $firstNameTextBox.Size = New-Object System.Drawing.Size(222,25)
    $firstNameTextBox.Font = $ObjFont

    $lastNameTextBoxLabel = New-Object System.Windows.Forms.Label
    $lastNameTextBoxLabel.Location = New-Object System.Drawing.Size(2,40)
    $lastNameTextBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $lastNameTextBoxLabel.TextAlign = "MiddleLeft"
    $lastNameTextBoxLabel.Text = "Last Name:"
    $lastNameTextBoxLabel.BackColor = "Transparent"

    $lastNameTextBox = New-Object System.Windows.Forms.TextBox
    $lastNameTextBox.Location = New-Object System.Drawing.Size(80,40)
    $lastNameTextBox.Size = New-Object System.Drawing.Size(222,25)
    $lastNameTextBox.Font = $ObjFont

    $FirstInit = $False
    $DottedFull = $False
    $FullName = $False
    If ($NameFormat -eq "FirstInit") {$FirstInit = $True}
    If ($NameFormat -eq "DottedFull") {$DottedFull = $True}
    If ($NameFormat -eq "FullName") {$FullName = $True}

    $NameFormatGroupBox = New-Object System.Windows.Forms.GroupBox
    $NameFormatGroupBox.Location = New-Object System.Drawing.Size(5,120)
    $NameFormatGroupBox.AutoSize = $True
    $NameFormatGroupBox.Font = $ObjFontBold
    $NameFormatGroupBox.Text = "Choose Username Format"

    $NamePatternRadio1 = New-Object System.Windows.Forms.RadioButton
    $NamePatternRadio1.Location = New-Object System.Drawing.Size(20,25)
    $NamePatternRadio1.Size = New-Object System.Drawing.Size(350,25)
    $NamePatternRadio1.Checked = $FirstInit
    $NamePatternRadio1.Font = $ObjFont
    $NamePatternRadio1.Text = "First Inital Last Name = lskywalker"
 
    $NamePatternRadio2 = New-Object System.Windows.Forms.RadioButton
    $NamePatternRadio2.Location = New-Object System.Drawing.Size(20,55)
    $NamePatternRadio2.Size = New-Object System.Drawing.Size(350,25)
    $NamePatternRadio2.Checked = $DottedFull
    $NamePatternRadio2.Font = $ObjFont
    $NamePatternRadio2.Text = "Dotted Full Name = luke.skywalker"
 
    $NamePatternRadio3 = New-Object System.Windows.Forms.RadioButton
    $NamePatternRadio3.Location = New-Object System.Drawing.Size(20,85)
    $NamePatternRadio3.Size = New-Object System.Drawing.Size(350,25)
    $NamePatternRadio3.Checked = $FullName
    $NamePatternRadio3.Font = $ObjFont
    $NamePatternRadio3.Text = "Full Name (Continuous) = lukeskywalker"

    $adDomainsDropDownBoxLabel = New-Object System.Windows.Forms.Label
    $adDomainsDropDownBoxLabel.Location = New-Object System.Drawing.Size(2,70)
    $adDomainsDropDownBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $adDomainsDropDownBoxLabel.TextAlign = "MiddleLeft"
    $adDomainsDropDownBoxLabel.Text = "Domain:"
    $adDomainsDropDownBoxLabel.BackColor = "Transparent"

    [array]$adDomains = (Get-ADForest).domains
    $adDomainsDropDownBox = New-Object System.Windows.Forms.ComboBox
    $adDomainsDropDownBox.Location = New-Object System.Drawing.Size(80,70)
    $adDomainsDropDownBox.Size = New-Object System.Drawing.Size(222,25)
    $adDomainsDropDownBox.Font = $ObjFont
    $adDomainsDropDownBox.DropDownStyle = "DropDownList"
    ForEach ($adDomain in $adDomains) { $adDomainsDropDownBox.Items.Add($adDomain) }
    $adDomainsDropDownBox.SelectedIndex = 0

    $adDomainsSelectButton = New-Object System.Windows.Forms.Button 
    $adDomainsSelectButton.Location = New-Object System.Drawing.Size(310,8)
    $adDomainsSelectButton.Size = New-Object System.Drawing.Size(110,90)
    $adDomainsSelectButton.Text = "Check Username and Domain"

    $DisplayNameTextLabel = New-Object System.Windows.Forms.Label
    $DisplayNameTextLabel.Location = New-Object System.Drawing.Size(440,10)
    $DisplayNameTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $DisplayNameTextLabel.TextAlign = "MiddleLeft"
    $DisplayNameTextLabel.Text = "Display Name:"
    $DisplayNameTextLabel.BackColor = "Transparent"

    $DisplayNameTextBox = New-Object System.Windows.Forms.TextBox
    $DisplayNameTextBox.Location = New-Object System.Drawing.Size(540,10)
    $DisplayNameTextBox.Size = New-Object System.Drawing.Size(300,25)
    $DisplayNameTextBox.Font = $ObjFontBold
    $DisplayNameTextBox.Enabled = $False

    $UsernameTextLabel = New-Object System.Windows.Forms.Label
    $UsernameTextLabel.Location = New-Object System.Drawing.Size(440,40)
    $UsernameTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $UsernameTextLabel.TextAlign = "MiddleLeft"
    $UsernameTextLabel.Text = "Username:"
    $UsernameTextLabel.BackColor = "Transparent"

    $UsernameTextBox = New-Object System.Windows.Forms.TextBox
    $UsernameTextBox.Location = New-Object System.Drawing.Size(540,40)
    $UsernameTextBox.Size = New-Object System.Drawing.Size(300,25)
    $UsernameTextBox.Font = $ObjFontBold
    $UsernameTextBox.Enabled = $False

    $ManagerDropDownLabel = New-Object System.Windows.Forms.Label
    $ManagerDropDownLabel.Location = New-Object System.Drawing.Size(440,70)
    $ManagerDropDownLabel.Size = New-Object System.Drawing.Size(100,25)
    $ManagerDropDownLabel.TextAlign = "MiddleLeft"
    $ManagerDropDownLabel.Text = "Manager:"
    $ManagerDropDownLabel.BackColor = "Transparent"

    $ManagerDropDownBox = New-Object System.Windows.Forms.ComboBox
    $ManagerDropDownBox.Location = New-Object System.Drawing.Size(540,70)
    $ManagerDropDownBox.Size = New-Object System.Drawing.Size(300,25)
    $ManagerDropDownBox.DropDownStyle = "DropDownList"
    $ManagerDropDownBox.Font = $ObjFont
    $ManagerDropDownBox.Enabled = $False

    $JobTitleTextLabel = New-Object System.Windows.Forms.Label
    $JobTitleTextLabel.Location = New-Object System.Drawing.Size(440,100)
    $JobTitleTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $JobTitleTextLabel.TextAlign = "MiddleLeft"
    $JobTitleTextLabel.Text = "Job Title:"
    $JobTitleTextLabel.BackColor = "Transparent"

    $JobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $JobTitleTextBox.Location = New-Object System.Drawing.Size(540,100)
    $JobTitleTextBox.Size = New-Object System.Drawing.Size(300,25)
    $JobTitleTextBox.Font = $ObjFont
    $JobTitleTextBox.Enabled = $False

    $EmailTextLabel = New-Object System.Windows.Forms.Label
    $EmailTextLabel.Location = New-Object System.Drawing.Size(440,130)
    $EmailTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $EmailTextLabel.TextAlign = "MiddleLeft"
    $EmailTextLabel.Text = "Email Address:"
    $EmailTextLabel.BackColor = "Transparent"

    $EmailTextBox = New-Object System.Windows.Forms.TextBox
    $EmailTextBox.Location = New-Object System.Drawing.Size(540,130)
    $EmailTextBox.Size = New-Object System.Drawing.Size(300,25)
    $EmailTextBox.Font = $ObjFont
    $EmailTextBox.Enabled = $False

    $DepartmentTextLabel = New-Object System.Windows.Forms.Label
    $DepartmentTextLabel.Location = New-Object System.Drawing.Size(440,160)
    $DepartmentTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $DepartmentTextLabel.TextAlign = "MiddleLeft"
    $DepartmentTextLabel.Text = "Department:"
    $DepartmentTextLabel.BackColor = "Transparent"

    $DepartmentTextBox = New-Object System.Windows.Forms.TextBox
    $DepartmentTextBox.Location = New-Object System.Drawing.Size(540,160)
    $DepartmentTextBox.Size = New-Object System.Drawing.Size(300,25)
    $DepartmentTextBox.Font = $ObjFont
    $DepartmentTextBox.Enabled = $False

    $CompanyTextLabel = New-Object System.Windows.Forms.Label
    $CompanyTextLabel.Location = New-Object System.Drawing.Size(440,190)
    $CompanyTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $CompanyTextLabel.TextAlign = "MiddleLeft"
    $CompanyTextLabel.Text = "Company:"
    $CompanyTextLabel.BackColor = "Transparent"

    $CompanyTextBox = New-Object System.Windows.Forms.TextBox
    $CompanyTextBox.Location = New-Object System.Drawing.Size(540,190)
    $CompanyTextBox.Size = New-Object System.Drawing.Size(300,25)
    $CompanyTextBox.Font = $ObjFont
    $CompanyTextBox.Enabled = $False

    $CountryDropDownLabel = New-Object System.Windows.Forms.Label
    $CountryDropDownLabel.Location = New-Object System.Drawing.Size(440,220)
    $CountryDropDownLabel.Size = New-Object System.Drawing.Size(100,25)
    $CountryDropDownLabel.TextAlign = "MiddleLeft"
    $CountryDropDownLabel.Text = "Country:"
    $CountryDropDownLabel.BackColor = "Transparent"

    $CountryDropDownBox = New-Object System.Windows.Forms.ComboBox
    $CountryDropDownBox.Location = New-Object System.Drawing.Size(540,220)
    $CountryDropDownBox.Size = New-Object System.Drawing.Size(300,25)
    $CountryDropDownBox.DropDownStyle = "DropDownList"
    $CountryDropDownBox.Font = $ObjFont
    $CountryDropDownBox.Enabled = $False

    $OfficePhoneTextLabel = New-Object System.Windows.Forms.Label
    $OfficePhoneTextLabel.Location = New-Object System.Drawing.Size(440,250)
    $OfficePhoneTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $OfficePhoneTextLabel.TextAlign = "MiddleLeft"
    $OfficePhoneTextLabel.Text = "Office Phone:"
    $OfficePhoneTextLabel.BackColor = "Transparent"

    $OfficePhoneTextBox = New-Object System.Windows.Forms.TextBox
    $OfficePhoneTextBox.Location = New-Object System.Drawing.Size(540,250)
    $OfficePhoneTextBox.Size = New-Object System.Drawing.Size(300,25)
    $OfficePhoneTextBox.Font = $ObjFont
    $OfficePhoneTextBox.Add_TextChanged({
        $this.Text = $this.Text -replace '\D'
        $this.Select($this.Text.Length, 0);
    })
    $OfficePhoneTextBox.Enabled = $False

    $CellPhoneTextLabel = New-Object System.Windows.Forms.Label
    $CellPhoneTextLabel.Location = New-Object System.Drawing.Size(440,280)
    $CellPhoneTextLabel.Size = New-Object System.Drawing.Size(100,25)
    $CellPhoneTextLabel.TextAlign = "MiddleLeft"
    $CellPhoneTextLabel.Text = "Cell Phone:"
    $CellPhoneTextLabel.BackColor = "Transparent"

    $CellPhoneTextBox = New-Object System.Windows.Forms.TextBox
    $CellPhoneTextBox.Location = New-Object System.Drawing.Size(540,280)
    $CellPhoneTextBox.Size = New-Object System.Drawing.Size(300,25)
    $CellPhoneTextBox.Font = $ObjFont
    $CellPhoneTextBox.Add_TextChanged({
        $this.Text = $this.Text -replace '\D'
        $this.Select($this.Text.Length, 0);
    })
    $CellPhoneTextBox.Enabled = $False

    $FacilityDropDownBoxLabel = New-Object System.Windows.Forms.Label
    $FacilityDropDownBoxLabel.Location = New-Object System.Drawing.Size(440,310)
    $FacilityDropDownBoxLabel.Size = New-Object System.Drawing.Size(100,25)
    $FacilityDropDownBoxLabel.TextAlign = "MiddleLeft"
    $FacilityDropDownBoxLabel.Text = "Facility:"
    $FacilityDropDownBoxLabel.BackColor = "Transparent"

    $FacilityDropDownBox = New-Object System.Windows.Forms.ComboBox
    $FacilityDropDownBox.Location = New-Object System.Drawing.Size(540,310)
    $FacilityDropDownBox.Size = New-Object System.Drawing.Size(300,25)
    $FacilityDropDownBox.DropDownStyle = "DropDownList"
    $FacilityDropDownBox.Font = $ObjFont
    $FacilityDropDownBox.Enabled = $False

    $Thirty = $False
    $Sixty = $False
    $Ninety = $False
    $Never = $False
    If ($AccountExpire -eq "30") {$Thirty = $True}
    If ($AccountExpire -eq "60") {$Sixty = $True}
    If ($AccountExpire -eq "90") {$Ninety = $True}
    If ($AccountExpire -eq "Never") {$Never = $True}

    $ExpirationGroupBox = New-Object System.Windows.Forms.GroupBox
    $ExpirationGroupBox.Location = New-Object System.Drawing.Size(440,340)
    $ExpirationGroupBox.Size = New-Object System.Drawing.Size(400,60)
    $ExpirationGroupBox.Font = $ObjFontBold
    $ExpirationGroupBox.Enabled = $False
    $ExpirationGroupBox.Text = "Account Expiration"

    $ExpirationRadio1 = New-Object System.Windows.Forms.RadioButton
    $ExpirationRadio1.Location = New-Object System.Drawing.Size(15,25)
    $ExpirationRadio1.Size = New-Object System.Drawing.Size(90,25)
    $ExpirationRadio1.Checked = $Thirty
    $ExpirationRadio1.Font = $ObjFont
    $ExpirationRadio1.Text = "30 Days"

    $ExpirationRadio2 = New-Object System.Windows.Forms.RadioButton
    $ExpirationRadio2.Location = New-Object System.Drawing.Size(105,25)
    $ExpirationRadio2.Size = New-Object System.Drawing.Size(90,25)
    $ExpirationRadio2.Checked = $Sixty
    $ExpirationRadio2.Font = $ObjFont
    $ExpirationRadio2.Text = "60 Days"

    $ExpirationRadio3 = New-Object System.Windows.Forms.RadioButton
    $ExpirationRadio3.Location = New-Object System.Drawing.Size(195,25)
    $ExpirationRadio3.Size = New-Object System.Drawing.Size(90,25)
    $ExpirationRadio3.Checked = $Ninety
    $ExpirationRadio3.Font = $ObjFont
    $ExpirationRadio3.Text = "90 Days"

    $ExpirationRadio4 = New-Object System.Windows.Forms.RadioButton
    $ExpirationRadio4.Location = New-Object System.Drawing.Size(285,25)
    $ExpirationRadio4.Size = New-Object System.Drawing.Size(90,25)
    $ExpirationRadio4.Checked = $Never
    $ExpirationRadio4.Font = $ObjFont
    $ExpirationRadio4.Text = "Never"

    $CancelButton1 = New-Object System.Windows.Forms.Button 
    $CancelButton1.Location = New-Object System.Drawing.Size(660,470)
    $CancelButton1.Size = New-Object System.Drawing.Size(180,50)
    $CancelButton1.Text = "Nevermind"
    $CancelButton1.Add_Click({ CancelForm })

    ### Add controls to the form
    $Form.Controls.Add($firstNameTextBoxLabel)
    $Form.Controls.Add($firstNameTextBox)
    $Form.Controls.Add($lastNameTextBoxLabel)
    $Form.Controls.Add($lastNameTextBox)
    $Form.Controls.Add($adDomainsDropDownBoxLabel)
    $Form.Controls.Add($adDomainsDropDownBox)
    $Form.Controls.Add($adDomainsSelectButton)
    $Form.Controls.Add($NameFormatGroupBox)
    $Form.Controls.Add($DisplayNameTextLabel)
    $Form.Controls.Add($DisplayNameTextBox)
    $Form.Controls.Add($UsernameTextLabel)
    $Form.Controls.Add($UsernameTextBox)
    $Form.Controls.Add($ManagerDropDownLabel)
    $Form.Controls.Add($ManagerDropDownBox)
    $Form.Controls.Add($JobTitleTextLabel)
    $Form.Controls.Add($JobTitleTextBox)
    $Form.Controls.Add($EmailTextLabel)
    $Form.Controls.Add($EmailTextBox)
    $Form.Controls.Add($DepartmentTextLabel)
    $Form.Controls.Add($DepartmentTextBox)
    $Form.Controls.Add($CompanyTextLabel)
    $Form.Controls.Add($CompanyTextBox)
    $Form.Controls.Add($CountryDropDownLabel)
    $Form.Controls.Add($CountryDropDownBox)
    $Form.Controls.Add($OfficePhoneTextLabel)
    $Form.Controls.Add($OfficePhoneTextBox)
    $Form.Controls.Add($CellPhoneTextLabel)
    $Form.Controls.Add($CellPhoneTextBox)
    $Form.Controls.Add($FacilityDropDownBoxLabel)
    $Form.Controls.Add($FacilityDropDownBox)
    $Form.Controls.Add($CancelButton1)
    $Form.Controls.Add($ExpirationGroupBox)
    $NameFormatGroupBox.Controls.AddRange(@($NamePatternRadio1,$NamePatternRadio2,$NamePatternRadio3))
    $ExpirationGroupBox.Controls.AddRange(@($ExpirationRadio1,$ExpirationRadio2,$ExpirationRadio3,$ExpirationRadio4))

    $adDomainsSelectButton.Add_Click({
        $firstNameTextBox.Enabled = $False
        $lastNameTextBox.Enabled = $False
        $adDomainsDropDownBox.Enabled = $False
        $Script:FirstName = $firstNameTextBox.Text
        $Script:LastName = $lastNameTextBox.Text
        $Script:Domain = $adDomainsDropDownBox.SelectedItem.ToString()
		$CompanyTextBox.Text = $Company
        $first = $FirstName.ToLower()
        $last = $LastName.ToLower()
        $firstW = $first -replace "\W"
        $lastW = $last -replace "\W"
        If ($first -eq "" -OR $first -eq $Null) { $DisplayName = $last }
        Else { $DisplayName = $first + " " + $last }
        $NameInfo = (Get-Culture).TextInfo
        $Script:SelectedDisplayName = $NameInfo.ToTitleCase($DisplayName)
        If ($NamePatternRadio1.Checked) { $Script:ProposedUsername = $firstW[0] + $lastW }
        If ($NamePatternRadio2.Checked) { $Script:ProposedUsername = $firstW + "." + $lastW }
        If ($NamePatternRadio3.Checked) { $Script:ProposedUsername = $firstW + $lastW }
        $Script:SelectedFirstName = $NameInfo.ToTitleCase($First)
        $Script:SelectedLastName = $NameInfo.ToTitleCase($Last)
        $Form.Controls.Remove($adDomainsSelectButton)
        $Form.Controls.Remove($NameFormatGroupBox)
        UserNameValidation
        RemainingInformation
    })

    ### Launch Form
    $Form.ShowDialog()
}


### Parse Responses #############################
Function UserNameValidation {
    If ($LastName -eq " " -OR $LastName -eq "" -OR $LastName -eq $Null) {
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Last Name Error"
        $MessageboxBody = "Please pick a valid Last Name. First name is optional, but I have to have a last name."
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
        MakeNewForm
    }
    $CheckUserName = Get-ADUser -Server $Domain -LDAPFilter "(SamAccountName=$ProposedUsername)"
    If ($CheckUserName -eq $Null) { Write-Output "Username Not In Use" }
    Else {
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "User Name Error"
        $MessageboxBody = "User Name '$ProposedUsername' already exists in the $Domain domain. Please try again or contact your administrator for help."
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
        MakeNewForm
    }
}


### Get More Info ###############################
Function RemainingInformation {
    $Form.Controls.Remove($adGroupListBoxPrecursor)
    $ManagerDropDownBox.Enabled = $True
    $JobTitleTextBox.Enabled = $True
    $EmailTextBox.Enabled = $True
    $DepartmentTextBox.Enabled = $True
    $CompanyTextBox.Enabled = $True
    $CountryDropDownBox.Enabled = $True
    $OfficePhoneTextBox.Enabled = $True
    $CellPhoneTextBox.Enabled = $True
    $FacilityDropDownBox.Enabled = $True
    $ExpirationGroupBox.Enabled = $True

    $adGroupListBox = New-Object System.Windows.Forms.ListBox 
    $adGroupListBox.Location = New-Object System.Drawing.Size(2,100) 
    $adGroupListBox.Size = New-Object System.Drawing.Size(420,420) 
    $adGroupListBox.Font = $ObjFont
    $adGroupListBox.Sorted = $True
    $Form.Controls.Add($adGroupListBox)

    $RemainingInfoButton = New-Object System.Windows.Forms.Button
    $RemainingInfoButton.Size = New-Object System.Drawing.Size(180,50)
    $RemainingInfoButton.Location = New-Object System.Drawing.Size(440,470)
    $RemainingInfoButton.Text = "Finish Him!"
    $Form.Controls.Add($RemainingInfoButton)

    ### Generate the user's email address if a domain UPN exists
    $DomainUPN = Get-ADForest | Select-Object UPNSuffixes -ExpandProperty UPNSuffixes
    If ($DomainUPN -eq "" -OR $DomainUPN -eq $Null) { $EmailText = "" }
    Else { $EmailText = $ProposedUsername + "@" + $DomainUPN }
    $EmailTextBox.Text = $EmailText

    $UsernameTextBox.Text = $ProposedUsername
    $DisplayNameTextBox.Text = $SelectedDisplayName

    ### Populate the group selection box
    $DN = Get-ADDomain -Server $Domain | Select-Object DistinguishedName
    $BuiltinOU = "*CN=Builtin,"
    $UsersOU = "*CN=Users,"
    $BuiltinOU = $BuiltinOU + $DN.DistinguishedName
    $UsersOU = $UsersOU + $DN.DistinguishedName
    $BuiltinOU = $BuiltinOU.ToString()
    $UsersOU = $UsersOU.ToString()
    [array]$adGroupItems = Get-ADGroup -Server $Domain -Filter * | Where-Object {($_.DistinguishedName -NotLike $BuiltinOU)} | Where-Object {($_.DistinguishedName -NotLike $UsersOU)} | Select-Object Name | Sort-Object Name
    $adGroupListBox.SelectionMode = "MultiExtended" # MultiExtended, MultiSimple, None, One
    ForEach ($adGroupItem in $adGroupItems) {
        $adListItem = $adGroupItem.Name
        $adGroupListBox.Items.Add($adListItem)
    }
    $adGroupListBox.TopIndex = 0
    $Script:adSelectedGroups = $adGroupListBox.SelectedItems

    ### Populate the manager selection box
    $ManagerList = [System.Collections.Generic.List[object]](Get-ADUser -Server $Domain -Filter {(Enabled -eq $True) -AND (Name -NotLike "SERVICE*") -AND (Name -NotLike "*(Admin)") -AND (Name -NotLike "*$")} -Properties * | Where-Object {($_.DistinguishedName -NotLike $UsersOU) -OR ($_.DistinguishedName -NotLike $UsersOU)} | Select-Object Name,SamAccountName | Sort-Object Name)
    If ($ManagerList -ne "" -OR $ManagerList -ne $Null) {
        $ManagerDropDownBox.ValueMember = "SamAccountName"
        $ManagerDropDownBox.DisplayMember = "Name"
        $ManagerDropDownBox.DataSource = $ManagerList
        $ManagerDropDownBox.SelectedItem = $Null
    }

    ### Populate the country selection box
    $CountryList = [System.Collections.Generic.List[object]]([CultureInfo]::GetCultures([System.Globalization.CultureTypes]::SpecificCultures) | ForEach-Object { (New-Object System.Globalization.RegionInfo $_.Name) } | Select-Object -Unique EnglishName,TwoLetterISORegionName | Sort-Object EnglishName)
    $CountryDropDownBox.ValueMember = "TwoLetterISORegionName"
    $CountryDropDownBox.DisplayMember = "EnglishName"
    $CountryDropDownBox.DataSource = $CountryList
    $CountryDropDownBox.SelectedValue = "US"

    ### Populate the facility selection box
    ForEach ($Facility in $Facilities) {
        $FacilityDropDownBox.Items.Add($Facility)
    }
    $FacilityDropDownBox.SelectedIndex = 0

    $RemainingInfoButton.Add_Click( {
        $Form.Controls.Remove($RemainingInfoButton)
        $NameInfo = (Get-Culture).TextInfo
		If ($ExpirationRadio1.Checked) { $Script:Expiration = 30 }
		If ($ExpirationRadio2.Checked) { $Script:Expiration = 60 }
		If ($ExpirationRadio3.Checked) { $Script:Expiration = 90 }
		If ($ExpirationRadio4.Checked) { $Script:Expiration = 0 }
        $Script:ProposedUsername = $UsernameTextBox.Text
        $Script:SelectedUsername = $UsernameTextBox.Text
        $Script:SelectedCompany = $CompanyTextBox.Text
        $Script:SelectedJobTitle = $NameInfo.ToTitleCase($JobTitleTextBox.Text)
        $Script:SelectedDepartment = $NameInfo.ToTitleCase($DepartmentTextBox.Text)
        $Script:SelectedManager = ($ManagerDropDownBox.SelectedItem).SamAccountName
        $Script:SelectedEmail = $EmailTextBox.Text
        $Script:SelectedCountry = ($CountryDropDownBox.SelectedItem).TwoLetterISORegionName
        $Script:SelectedOfficePhone = $OfficePhoneTextBox.Text
        $Script:SelectedCellPhone = $CellPhoneTextBox.Text
        $Script:SelectedFacility = $FacilityDropDownBox.SelectedItem
        UserNameValidation
        TimeToMakeTheUser
    } )
}


### Random Password Generator ###################
Function Get-RandomCharacters($length) { 
    $characters = "1234567890abcdefghiklmnoprstuvwxyzABCDEFGHKLMNOPRSTUVWXYZ"
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    Return [String]$characters[$random]
}


### Make The User ###############################
Function TimeToMakeTheUser {
    $DomainUPN = Get-ADForest | Select-Object UPNSuffixes -ExpandProperty UPNSuffixes
    If ($DomainUPN -eq "" -OR $DomainUPN -eq $Null) { $Script:UPN = $ProposedUsername + "@" + $Domain }
    Else { $Script:UPN = $ProposedUsername + "@" + $DomainUPN }

    ### Apply country transforms to the office phone number
    If ($SelectedOfficePhone -ne "") {
        Switch ($SelectedOfficePhone) {
            { $SelectedCountry -eq "US" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"")
            }
            { $SelectedCountry -eq "CA" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+1-")
            }
            { $SelectedCountry -eq "GB" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+44-") }
			{ $SelectedCountry -eq "AU" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+61-") }
			{ $SelectedCountry -eq "AF" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+93-") }
			{ $SelectedCountry -eq "AL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+355-") }
			{ $SelectedCountry -eq "DZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+213-") }
			{ $SelectedCountry -eq "AR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+54-") }
			{ $SelectedCountry -eq "AM" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+374-") }
			{ $SelectedCountry -eq "AT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+43-") }
			{ $SelectedCountry -eq "AZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+994-") }
			{ $SelectedCountry -eq "BH" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+973-") }
			{ $SelectedCountry -eq "BD" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+880-") }
			{ $SelectedCountry -eq "BY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+375-") }
			{ $SelectedCountry -eq "BE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+32-") }
			{ $SelectedCountry -eq "BZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+501-") }
			{ $SelectedCountry -eq "BO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+591-") }
			{ $SelectedCountry -eq "BA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+387-") }
			{ $SelectedCountry -eq "BR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+55-") }
			{ $SelectedCountry -eq "BN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+673-") }
			{ $SelectedCountry -eq "BG" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+359-") }
			{ $SelectedCountry -eq "KH" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+855-") }
			{ $SelectedCountry -eq "CL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+56-") }
			{ $SelectedCountry -eq "CO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+57-") }
			{ $SelectedCountry -eq "CR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+506-") }
			{ $SelectedCountry -eq "HR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+385-") }
			{ $SelectedCountry -eq "CZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+420-") }
			{ $SelectedCountry -eq "DK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+45-") }
			{ $SelectedCountry -eq "DO" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "EC" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+593-") }
			{ $SelectedCountry -eq "EG" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+20-") }
			{ $SelectedCountry -eq "SV" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+503-") }
			{ $SelectedCountry -eq "EE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+372-") }
			{ $SelectedCountry -eq "ET" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+251-") }
			{ $SelectedCountry -eq "FO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+298-") }
			{ $SelectedCountry -eq "FI" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+358-") }
			{ $SelectedCountry -eq "FR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+33-") }
			{ $SelectedCountry -eq "GE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+995-") }
			{ $SelectedCountry -eq "DE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+49-") }
			{ $SelectedCountry -eq "GR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+30-") }
			{ $SelectedCountry -eq "GL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+299-") }
			{ $SelectedCountry -eq "GT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+502-") }
			{ $SelectedCountry -eq "HN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+504-") }
			{ $SelectedCountry -eq "HK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+852-") }
			{ $SelectedCountry -eq "HU" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+36-") }
			{ $SelectedCountry -eq "IS" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+354-") }
			{ $SelectedCountry -eq "IN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+91-") }
			{ $SelectedCountry -eq "ID" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+62-") }
			{ $SelectedCountry -eq "IR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+98-") }
			{ $SelectedCountry -eq "IQ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+964-") }
			{ $SelectedCountry -eq "IE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+353-") }
			{ $SelectedCountry -eq "PK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+92-") }
			{ $SelectedCountry -eq "IL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+972-") }
			{ $SelectedCountry -eq "IT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+39-") }
			{ $SelectedCountry -eq "JM" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "JP" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+81-") }
			{ $SelectedCountry -eq "JO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+962-") }
			{ $SelectedCountry -eq "KZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+7-") }
			{ $SelectedCountry -eq "KE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+254-") }
			{ $SelectedCountry -eq "KR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+82-") }
			{ $SelectedCountry -eq "KW" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+965-") }
			{ $SelectedCountry -eq "KG" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+996-") }
			{ $SelectedCountry -eq "LA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+856-") }
			{ $SelectedCountry -eq "LV" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+371-") }
			{ $SelectedCountry -eq "LB" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+961-") }
			{ $SelectedCountry -eq "LY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+218-") }
			{ $SelectedCountry -eq "LI" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+423-") }
			{ $SelectedCountry -eq "LT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+370-") }
			{ $SelectedCountry -eq "LU" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+352-") }
			{ $SelectedCountry -eq "MO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+853-") }
			{ $SelectedCountry -eq "MK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+389-") }
			{ $SelectedCountry -eq "MY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+60-") }
			{ $SelectedCountry -eq "MV" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+960-") }
			{ $SelectedCountry -eq "MT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+356-") }
			{ $SelectedCountry -eq "MX" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+52-") }
			{ $SelectedCountry -eq "MN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+976-") }
			{ $SelectedCountry -eq "ME" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+382-") }
			{ $SelectedCountry -eq "MA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+212-") }
			{ $SelectedCountry -eq "NP" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+977-") }
			{ $SelectedCountry -eq "NL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+31-") }
			{ $SelectedCountry -eq "NZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+64-") }
			{ $SelectedCountry -eq "NI" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+505-") }
			{ $SelectedCountry -eq "NG" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+234-") }
			{ $SelectedCountry -eq "NO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+47-") }
			{ $SelectedCountry -eq "OM" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+968-") }
			{ $SelectedCountry -eq "PA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+507-") }
			{ $SelectedCountry -eq "PY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+595-") }
			{ $SelectedCountry -eq "CN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+86-") }
			{ $SelectedCountry -eq "PE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+51-") }
			{ $SelectedCountry -eq "PH" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+63-") }
			{ $SelectedCountry -eq "PL" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+48-") }
			{ $SelectedCountry -eq "PT" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+351-") }
			{ $SelectedCountry -eq "MC" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+377-") }
			{ $SelectedCountry -eq "PR" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "QA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+974-") }
			{ $SelectedCountry -eq "RO" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+40-") }
			{ $SelectedCountry -eq "RU" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+7-") }
			{ $SelectedCountry -eq "RW" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+250-") }
			{ $SelectedCountry -eq "SA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+966-") }
			{ $SelectedCountry -eq "SN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+221-") }
			{ $SelectedCountry -eq "RS" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+381-") }
			{ $SelectedCountry -eq "CS" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+382-") }
			{ $SelectedCountry -eq "SG" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+65-") }
			{ $SelectedCountry -eq "SK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+421-") }
			{ $SelectedCountry -eq "SI" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+386-") }
			{ $SelectedCountry -eq "ZA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+27-") }
			{ $SelectedCountry -eq "ES" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+34-") }
			{ $SelectedCountry -eq "LK" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+94-") }
			{ $SelectedCountry -eq "SE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+46-") }
			{ $SelectedCountry -eq "CH" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+41-") }
			{ $SelectedCountry -eq "SY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+963-") }
			{ $SelectedCountry -eq "TW" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+886-") }
			{ $SelectedCountry -eq "TJ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+992-") }
			{ $SelectedCountry -eq "TH" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+66-") }
			{ $SelectedCountry -eq "TT" } {
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(6,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(3,"-")
                $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "TN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+216-") }
			{ $SelectedCountry -eq "TR" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+90-") }
			{ $SelectedCountry -eq "TM" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+993-") }
			{ $SelectedCountry -eq "AE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+971-") }
			{ $SelectedCountry -eq "UA" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+380-") }
			{ $SelectedCountry -eq "UY" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+598-") }
			{ $SelectedCountry -eq "UZ" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+998-") }
			{ $SelectedCountry -eq "VN" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+84-") }
			{ $SelectedCountry -eq "YE" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+967-") }
			{ $SelectedCountry -eq "ZW" } { $SelectedOfficePhone = $SelectedOfficePhone.Insert(0,"+263-") }
			default { $SelectedOfficePhone = $SelectedOfficePhone }
        }
    }

    ### Apply country transforms to the cell phone number
    If ($SelectedCellPhone -ne "") {
        Switch ($SelectedCellPhone) {
            { $SelectedCountry -eq "US" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"")
            }
            { $SelectedCountry -eq "CA" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+1-")
            }
            { $SelectedCountry -eq "GB" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+44-") }
            { $SelectedCountry -eq "AU" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+61-") }
			{ $SelectedCountry -eq "AF" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+93-") }
			{ $SelectedCountry -eq "AL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+355-") }
			{ $SelectedCountry -eq "DZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+213-") }
			{ $SelectedCountry -eq "AR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+54-") }
			{ $SelectedCountry -eq "AM" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+374-") }
			{ $SelectedCountry -eq "AT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+43-") }
			{ $SelectedCountry -eq "AZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+994-") }
			{ $SelectedCountry -eq "BH" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+973-") }
			{ $SelectedCountry -eq "BD" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+880-") }
			{ $SelectedCountry -eq "BY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+375-") }
			{ $SelectedCountry -eq "BE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+32-") }
			{ $SelectedCountry -eq "BZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+501-") }
			{ $SelectedCountry -eq "BO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+591-") }
			{ $SelectedCountry -eq "BA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+387-") }
			{ $SelectedCountry -eq "BR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+55-") }
			{ $SelectedCountry -eq "BN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+673-") }
			{ $SelectedCountry -eq "BG" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+359-") }
			{ $SelectedCountry -eq "KH" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+855-") }
			{ $SelectedCountry -eq "CL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+56-") }
			{ $SelectedCountry -eq "CO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+57-") }
			{ $SelectedCountry -eq "CR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+506-") }
			{ $SelectedCountry -eq "HR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+385-") }
			{ $SelectedCountry -eq "CZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+420-") }
			{ $SelectedCountry -eq "DK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+45-") }
			{ $SelectedCountry -eq "DO" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "EC" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+593-") }
			{ $SelectedCountry -eq "EG" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+20-") }
			{ $SelectedCountry -eq "SV" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+503-") }
			{ $SelectedCountry -eq "EE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+372-") }
			{ $SelectedCountry -eq "ET" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+251-") }
			{ $SelectedCountry -eq "FO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+298-") }
			{ $SelectedCountry -eq "FI" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+358-") }
			{ $SelectedCountry -eq "FR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+33-") }
			{ $SelectedCountry -eq "GE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+995-") }
			{ $SelectedCountry -eq "DE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+49-") }
			{ $SelectedCountry -eq "GR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+30-") }
			{ $SelectedCountry -eq "GL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+299-") }
			{ $SelectedCountry -eq "GT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+502-") }
			{ $SelectedCountry -eq "HN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+504-") }
			{ $SelectedCountry -eq "HK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+852-") }
			{ $SelectedCountry -eq "HU" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+36-") }
			{ $SelectedCountry -eq "IS" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+354-") }
			{ $SelectedCountry -eq "IN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+91-") }
			{ $SelectedCountry -eq "ID" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+62-") }
			{ $SelectedCountry -eq "IR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+98-") }
			{ $SelectedCountry -eq "IQ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+964-") }
			{ $SelectedCountry -eq "IE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+353-") }
			{ $SelectedCountry -eq "PK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+92-") }
			{ $SelectedCountry -eq "IL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+972-") }
			{ $SelectedCountry -eq "IT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+39-") }
			{ $SelectedCountry -eq "JM" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "JP" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+81-") }
			{ $SelectedCountry -eq "JO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+962-") }
			{ $SelectedCountry -eq "KZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+7-") }
			{ $SelectedCountry -eq "KE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+254-") }
			{ $SelectedCountry -eq "KR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+82-") }
			{ $SelectedCountry -eq "KW" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+965-") }
			{ $SelectedCountry -eq "KG" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+996-") }
			{ $SelectedCountry -eq "LA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+856-") }
			{ $SelectedCountry -eq "LV" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+371-") }
			{ $SelectedCountry -eq "LB" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+961-") }
			{ $SelectedCountry -eq "LY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+218-") }
			{ $SelectedCountry -eq "LI" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+423-") }
			{ $SelectedCountry -eq "LT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+370-") }
			{ $SelectedCountry -eq "LU" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+352-") }
			{ $SelectedCountry -eq "MO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+853-") }
			{ $SelectedCountry -eq "MK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+389-") }
			{ $SelectedCountry -eq "MY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+60-") }
			{ $SelectedCountry -eq "MV" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+960-") }
			{ $SelectedCountry -eq "MT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+356-") }
			{ $SelectedCountry -eq "MX" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+52-") }
			{ $SelectedCountry -eq "MN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+976-") }
			{ $SelectedCountry -eq "ME" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+382-") }
			{ $SelectedCountry -eq "MA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+212-") }
			{ $SelectedCountry -eq "NP" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+977-") }
			{ $SelectedCountry -eq "NL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+31-") }
			{ $SelectedCountry -eq "NZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+64-") }
			{ $SelectedCountry -eq "NI" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+505-") }
			{ $SelectedCountry -eq "NG" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+234-") }
			{ $SelectedCountry -eq "NO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+47-") }
			{ $SelectedCountry -eq "OM" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+968-") }
			{ $SelectedCountry -eq "PA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+507-") }
			{ $SelectedCountry -eq "PY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+595-") }
			{ $SelectedCountry -eq "CN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+86-") }
			{ $SelectedCountry -eq "PE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+51-") }
			{ $SelectedCountry -eq "PH" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+63-") }
			{ $SelectedCountry -eq "PL" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+48-") }
			{ $SelectedCountry -eq "PT" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+351-") }
			{ $SelectedCountry -eq "MC" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+377-") }
			{ $SelectedCountry -eq "PR" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "QA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+974-") }
			{ $SelectedCountry -eq "RO" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+40-") }
			{ $SelectedCountry -eq "RU" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+7-") }
			{ $SelectedCountry -eq "RW" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+250-") }
			{ $SelectedCountry -eq "SA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+966-") }
			{ $SelectedCountry -eq "SN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+221-") }
			{ $SelectedCountry -eq "RS" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+381-") }
			{ $SelectedCountry -eq "CS" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+382-") }
			{ $SelectedCountry -eq "SG" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+65-") }
			{ $SelectedCountry -eq "SK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+421-") }
			{ $SelectedCountry -eq "SI" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+386-") }
			{ $SelectedCountry -eq "ZA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+27-") }
			{ $SelectedCountry -eq "ES" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+34-") }
			{ $SelectedCountry -eq "LK" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+94-") }
			{ $SelectedCountry -eq "SE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+46-") }
			{ $SelectedCountry -eq "CH" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+41-") }
			{ $SelectedCountry -eq "SY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+963-") }
			{ $SelectedCountry -eq "TW" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+886-") }
			{ $SelectedCountry -eq "TJ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+992-") }
			{ $SelectedCountry -eq "TH" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+66-") }
			{ $SelectedCountry -eq "TT" } {
                $SelectedCellPhone = $SelectedCellPhone.Insert(6,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(3,"-")
                $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+1-")
            }
			{ $SelectedCountry -eq "TN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+216-") }
			{ $SelectedCountry -eq "TR" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+90-") }
			{ $SelectedCountry -eq "TM" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+993-") }
			{ $SelectedCountry -eq "AE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+971-") }
			{ $SelectedCountry -eq "UA" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+380-") }
			{ $SelectedCountry -eq "UY" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+598-") }
			{ $SelectedCountry -eq "UZ" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+998-") }
			{ $SelectedCountry -eq "VN" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+84-") }
			{ $SelectedCountry -eq "YE" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+967-") }
			{ $SelectedCountry -eq "ZW" } { $SelectedCellPhone = $SelectedCellPhone.Insert(0,"+263-") }
            default { $SelectedCellPhone = $SelectedCellPhone }
        }
    }

### Facilities ##################################
    If ($SelectedFacility -eq "Portland") {
        $StreetAddress = "123 NW Fake Street"
        $PostalCode = "97204"
        $State = "OR"
        $City = "Portland"
    }
    If ($SelectedFacility -eq "Seattle") {
        $StreetAddress = ""
        $PostalCode = "98115"
        $State = "WA"
        $City = "Seattle"
    }
    If ($SelectedFacility -eq "Sydney") {
        $StreetAddress = ""
        $PostalCode = "2067"
        $State = "NSW"
        $City = "Sydney"
    }
    If ($SelectedFacility -eq "Newcastle") {
        $StreetAddress = ""
        $PostalCode = "NE1 3PJ"
        $State = ""
        $City = "Newcastle upon Tyne"
    }
    If ($SelectedFacility -eq "Remote") {
        $StreetAddress = ""
        $PostalCode = ""
        $State = ""
        $City = ""
    }
#################################################
    ### Create the user
    $Password = (Get-RandomCharacters -length 15)
    $ExpirationString = $Expiration.ToString()
    If ($SelectedManager -eq $Null) { $SelectedManager = "" }
    New-ADUser -Server $Domain -SamAccountName $SelectedUserName -Name $SelectedDisplayName -Enabled $True -UserPrincipalName $UPN -DisplayName $SelectedDisplayName -GivenName $SelectedFirstName -Surname $SelectedLastName -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) -ChangePasswordAtLogon $True | Wait-Job
    If ($City -ne "" -OR $City -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -City $City}
    If ($SelectedFacility -ne "" -OR $SelectedFacility -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Office $SelectedFacility}
    If ($SelectedOfficePhone -ne "" -OR $SelectedOfficePhone -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -OfficePhone $SelectedOfficePhone}
    If ($SelectedCellPhone -ne "" -OR $SelectedCellPhone -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -MobilePhone $SelectedCellPhone}
    If ($StreetAddress -ne "" -OR $StreetAddress -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -StreetAddress $StreetAddress}
    If ($SelectedJobTitle -ne "" -OR $SelectedJobTitle -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Title $SelectedJobTitle}
    If ($State -ne "" -OR $State -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -State $State}
    If ($SelectedCountry -ne "" -OR $SelectedCountry -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Country $SelectedCountry}
    If ($SelectedDepartment -ne "" -OR $SelectedDepartment -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Department $SelectedDepartment}
    If ($PostalCode -ne "" -OR $PostalCode -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -PostalCode $PostalCode}
    If ($SelectedManager -ne "" -OR $SelectedManager -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Manager $SelectedManager}
    If ($SelectedCompany -ne "" -OR $SelectedCompany -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -Company $SelectedCompany}
    If ($SelectedEmail -ne "" -OR $SelectedEmail -ne $Null) {Set-ADUser -Server $Domain -Identity $SelectedUserName -EmailAddress $SelectedEmail}
	If ($Expiration -ne 0 -OR $Expiration -ne $Null) { Set-ADAccountExpiration -Server $Domain -Identity $SelectedUserName -TimeSpan $ExpirationString }
    If ($adSelectedGroups -ne "" -OR $adSelectedGroups -ne $Null) { ForEach ($Group in $adSelectedGroups) {Add-ADGroupMember -Server $Domain -Identity $Group -Members $SelectedUserName} }

    $TestEnabled = (Get-ADUser -Identity $SelectedUserName).Enabled
    If ($TestEnabled -eq $False) { Set-ADUser -Server $Domain -Identity $SelectedUserName -Enabled $True }

    ### Manager Email 
    $ExpiresOn = (Get-Date).AddDays($Expiration)
    If ($Expiration -eq 0) { $ExpirationDate = "Never" }
    If ($Expiration -ne 0) { $ExpirationDate = Get-Date $ExpiresOn -Format "d MMMM, yyyy" }
    If ($SelectedManager -ne "" -OR $SelectedManager -ne $Null) { $toAddr = (Get-ADUser -Server $Domain -Identity $SelectedManager -Properties * | Select-Object EmailAddress | ForEach-Object {$_.EmailAddress}) }
    $Subject = "New User Created - $SelectedUserName"
    $Body = @("A new user in the $Domain domain has been created. During creation this user was assigned to you as their new manager. Below you will find the details of this user's account. Please note that the password must be changed at the first login and a ticket must be raised if there are sign-in issues.<br />")
    $Body += "<ul>
        <li>Full Name = $SelectedDisplayName</li>
        <li>Username = $SelectedUserName</li>
        <li>Password = $Password</li>
        <li>Job Title = $SelectedJobTitle</li>
        <li>Department = $SelectedDepartment</li>
        <li>Location = $SelectedFacility</li>
        <li>Office Phone = $SelectedOfficePhone</li>
        <li>Email = $SelectedEmail</li>
        <li>Account Expires = $ExpirationDate</li>
        </ul><br />
        If a mailbox was requested, creation of the new mailbox can take up to 5 hours to provision, please wait until tomorrow before sending any emails to the new user's address. If you have any questions, please submit a ticket and someone will assist you as soon as possible.
        "

    ### Check if operation was successful
    Try {
        Get-ADUser -Server $Domain -Identity $SelectedUserName | Out-Null
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "User Created"
        $MessageboxBody = "The user $SelectedUserName ($SelectedDisplayName) has been successfully created in the $Domain domain. Please make a note of the password and provide it to the user: $Password. Please check Active Directory to ensure everything is correct."
        $MessageIcon = [System.Windows.MessageBoxImage]::Information
        [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
        Send-MailMessage -To $toAddr -From $fromAddr -Subject $Subject -Body "$Body" -SmtpServer $smtpServer -BodyAsHtml
        CancelForm
    }
    Catch {
        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::OK
        $MessageboxTitle = "Something Went Wrong"
        $MessageboxBody = "I tried to find the new user in Active Directory but something went wrong. Please check to see if the user has been created. If this is not the first time you have seen this, please contact your Adminisrator for assistance."
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
        CancelForm
    }
}


MakeForm
