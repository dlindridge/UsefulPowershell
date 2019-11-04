<#
    .SYNOPSIS
    Get FSMO Role assignments and other stuff for the selected domain.

    .DESCRIPTION
    Usage: Create a shortcut pointing to (powershell.exe -file "C:\Scripts\UserStuff\FSMOQuery.ps1")
#>
#################################################
<#
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 23, 2019
    Modified: October 23, 2019
#>
#################################################

$Path = "C:\Scripts\UserStuff" # Path to image files (probably the same as the script location)
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

 HideConsole


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
    
    [array]$adDomains = (Get-ADForest).Domains
    $adDomainsDropDownBox = New-Object System.Windows.Forms.ComboBox
    $adDomainsDropDownBox.Location = New-Object System.Drawing.Size(80,10)
    $adDomainsDropDownBox.Size = New-Object System.Drawing.Size(225,25)
    $adDomainsDropDownBox.Font = $ObjFont
    $adDomainsDropDownBox.TabIndex = 1
    $adDomainsDropDownBox.DropDownStyle = "DropDownList"
    ForEach ($adDomain in $adDomains) { $adDomainsDropDownBox.Items.Add($adDomain) }
    $adDomainsDropDownBox.SelectedIndex = 0
    $Form.Controls.Add($adDomainsDropDownBox)

    $SelectDomainButton = New-Object System.Windows.Forms.Button 
    $SelectDomainButton.Location = New-Object System.Drawing.Size(310,10)
    $SelectDomainButton.Size = New-Object System.Drawing.Size(111,25)
    $SelectDomainButton.Text = "Select Domain"
    $SelectDomainButton.TabIndex = 0
    $Form.Controls.Add($SelectDomainButton)

    $StartOverButton = New-Object System.Windows.Forms.Button 
    $StartOverButton.Location = New-Object System.Drawing.Size(310,10)
    $StartOverButton.Size = New-Object System.Drawing.Size(111,25)
    $StartOverButton.Text = "Start Over"
    $StartOverButton.TabIndex = 0

    $CancelButton1 = New-Object System.Windows.Forms.Button 
    $CancelButton1.Location = New-Object System.Drawing.Size(520,410)
    $CancelButton1.Size = New-Object System.Drawing.Size(180,50)
    $CancelButton1.Text = "Nevermind"
    $CancelButton1.TabIndex = 2
    $CancelButton1.Add_Click({ CancelForm })
    $Form.Controls.Add($CancelButton1)

    $ForestModeTextBoxLabel = New-Object System.Windows.Forms.Label
    $ForestModeTextBoxLabel.Location = New-Object System.Drawing.Size(2,40)
    $ForestModeTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $ForestModeTextBoxLabel.TextAlign = "MiddleLeft"
    $ForestModeTextBoxLabel.Text = "Forest Mode:"
    $ForestModeTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($ForestModeTextBoxLabel)

    $ForestModeTextBox = New-Object System.Windows.Forms.TextBox
    $ForestModeTextBox.Location = New-Object System.Drawing.Size(115,40)
    $ForestModeTextBox.Size = New-Object System.Drawing.Size(225,25)
    $ForestModeTextBox.Font = $ObjFont
    $ForestModeTextBox.TabStop = $False
    $Form.Controls.Add($ForestModeTextBox)

    $DomainModeTextBoxLabel = New-Object System.Windows.Forms.Label
    $DomainModeTextBoxLabel.Location = New-Object System.Drawing.Size(2,70)
    $DomainModeTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $DomainModeTextBoxLabel.TextAlign = "MiddleLeft"
    $DomainModeTextBoxLabel.Text = "Domain Mode:"
    $DomainModeTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($DomainModeTextBoxLabel)

    $DomainModeTextBox = New-Object System.Windows.Forms.TextBox
    $DomainModeTextBox.Location = New-Object System.Drawing.Size(115,70)
    $DomainModeTextBox.Size = New-Object System.Drawing.Size(225,25)
    $DomainModeTextBox.Font = $ObjFont
    $DomainModeTextBox.TabStop = $False
    $Form.Controls.Add($DomainModeTextBox)

    $UPNSuffTextBoxLabel = New-Object System.Windows.Forms.Label
    $UPNSuffTextBoxLabel.Location = New-Object System.Drawing.Size(2,100)
    $UPNSuffTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $UPNSuffTextBoxLabel.TextAlign = "MiddleLeft"
    $UPNSuffTextBoxLabel.Text = "UPN Suffixes:"
    $UPNSuffTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($UPNSuffTextBoxLabel)

    $UPNSuffTextBox = New-Object System.Windows.Forms.TextBox
    $UPNSuffTextBox.Location = New-Object System.Drawing.Size(115,100)
    $UPNSuffTextBox.Size = New-Object System.Drawing.Size(225,25)
    $UPNSuffTextBox.Font = $ObjFont
    $UPNSuffTextBox.TabStop = $False
    $Form.Controls.Add($UPNSuffTextBox)

    $PDCeTextBoxLabel = New-Object System.Windows.Forms.Label
    $PDCeTextBoxLabel.Location = New-Object System.Drawing.Size(2,130)
    $PDCeTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $PDCeTextBoxLabel.TextAlign = "MiddleLeft"
    $PDCeTextBoxLabel.Text = "PDC Emulator:"
    $PDCeTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($PDCeTextBoxLabel)

    $PDCeTextBox = New-Object System.Windows.Forms.TextBox
    $PDCeTextBox.Location = New-Object System.Drawing.Size(115,130)
    $PDCeTextBox.Size = New-Object System.Drawing.Size(225,25)
    $PDCeTextBox.Font = $ObjFont
    $PDCeTextBox.TabStop = $False
    $Form.Controls.Add($PDCeTextBox)

    $RIDMastTextBoxLabel = New-Object System.Windows.Forms.Label
    $RIDMastTextBoxLabel.Location = New-Object System.Drawing.Size(2,160)
    $RIDMastTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $RIDMastTextBoxLabel.TextAlign = "MiddleLeft"
    $RIDMastTextBoxLabel.Text = "RID Master:"
    $RIDMastTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($RIDMastTextBoxLabel)

    $RIDMastTextBox = New-Object System.Windows.Forms.TextBox
    $RIDMastTextBox.Location = New-Object System.Drawing.Size(115,160)
    $RIDMastTextBox.Size = New-Object System.Drawing.Size(225,25)
    $RIDMastTextBox.Font = $ObjFont
    $RIDMastTextBox.TabStop = $False
    $Form.Controls.Add($RIDMastTextBox)

    $InfMastTextBoxLabel = New-Object System.Windows.Forms.Label
    $InfMastTextBoxLabel.Location = New-Object System.Drawing.Size(2,190)
    $InfMastTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $InfMastTextBoxLabel.TextAlign = "MiddleLeft"
    $InfMastTextBoxLabel.Text = "Inf. Master:"
    $InfMastTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($InfMastTextBoxLabel)

    $InfMastTextBox = New-Object System.Windows.Forms.TextBox
    $InfMastTextBox.Location = New-Object System.Drawing.Size(115,190)
    $InfMastTextBox.Size = New-Object System.Drawing.Size(225,25)
    $InfMastTextBox.Font = $ObjFont
    $InfMastTextBox.TabStop = $False
    $Form.Controls.Add($InfMastTextBox)

    $SchemaMastTextBoxLabel = New-Object System.Windows.Forms.Label
    $SchemaMastTextBoxLabel.Location = New-Object System.Drawing.Size(2,220)
    $SchemaMastTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $SchemaMastTextBoxLabel.TextAlign = "MiddleLeft"
    $SchemaMastTextBoxLabel.Text = "Schema Master:"
    $SchemaMastTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($SchemaMastTextBoxLabel)

    $SchemaMastTextBox = New-Object System.Windows.Forms.TextBox
    $SchemaMastTextBox.Location = New-Object System.Drawing.Size(115,220)
    $SchemaMastTextBox.Size = New-Object System.Drawing.Size(225,25)
    $SchemaMastTextBox.Font = $ObjFont
    $SchemaMastTextBox.TabStop = $False
    $Form.Controls.Add($SchemaMastTextBox)

    $NamingMastTextBoxLabel = New-Object System.Windows.Forms.Label
    $NamingMastTextBoxLabel.Location = New-Object System.Drawing.Size(2,250)
    $NamingMastTextBoxLabel.Size = New-Object System.Drawing.Size(110,25)
    $NamingMastTextBoxLabel.TextAlign = "MiddleLeft"
    $NamingMastTextBoxLabel.Text = "Naming Master:"
    $NamingMastTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($NamingMastTextBoxLabel)

    $NamingMastTextBox = New-Object System.Windows.Forms.TextBox
    $NamingMastTextBox.Location = New-Object System.Drawing.Size(115,250)
    $NamingMastTextBox.Size = New-Object System.Drawing.Size(225,25)
    $NamingMastTextBox.Font = $ObjFont
    $NamingMastTextBox.TabStop = $False
    $Form.Controls.Add($NamingMastTextBox)

    $GlobalCatTextBoxLabel = New-Object System.Windows.Forms.Label
    $GlobalCatTextBoxLabel.Location = New-Object System.Drawing.Size(350,40)
    $GlobalCatTextBoxLabel.Size = New-Object System.Drawing.Size(150,25)
    $GlobalCatTextBoxLabel.TextAlign = "MiddleLeft"
    $GlobalCatTextBoxLabel.Text = "Global Catalog Servers:"
    $GlobalCatTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($GlobalCatTextBoxLabel)

    $GlobalCatListBox = New-Object System.Windows.Forms.ListBox 
    $GlobalCatListBox.Location = New-Object System.Drawing.Size(350,70) 
    $GlobalCatListBox.Size = New-Object System.Drawing.Size(350,116) 
    $GlobalCatListBox.Font = $ObjFont
    $GlobalCatListBox.TabStop = $False
    $GlobalCatListBox.Sorted = $True
    $Form.Controls.Add($GlobalCatListBox)

    $ForestSitesTextBoxLabel = New-Object System.Windows.Forms.Label
    $ForestSitesTextBoxLabel.Location = New-Object System.Drawing.Size(350,190)
    $ForestSitesTextBoxLabel.Size = New-Object System.Drawing.Size(150,25)
    $ForestSitesTextBoxLabel.TextAlign = "MiddleLeft"
    $ForestSitesTextBoxLabel.Text = "Forest Sites:"
    $ForestSitesTextBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($ForestSitesTextBoxLabel)

    $ForestSitesListBox = New-Object System.Windows.Forms.ListBox 
    $ForestSitesListBox.Location = New-Object System.Drawing.Size(350,220) 
    $ForestSitesListBox.Size = New-Object System.Drawing.Size(350,145) 
    $ForestSitesListBox.Font = $ObjFont
    $ForestSitesListBox.TabStop = $False
    $ForestSitesListBox.Sorted = $True
    $Form.Controls.Add($ForestSitesListBox)

    $SelectDomainButton.Add_Click({
        $Script:Domain = $adDomainsDropDownBox.SelectedItem
        $Form.Controls.Add($StartOverButton)
        $Form.Controls.Remove($SelectDomainButton)
        $adForest = Get-ADForest $Domain
        $adDomain = Get-ADDomain $Domain
        $ForestModeTextBox.Text = $adForest.ForestMode
        $DomainModeTextBox.Text = $adDomain.DomainMode
        $UPNSuffTextBox.Text = $adForest.UPNSuffixes
        $PDCeTextBox.Text = $adDomain.PDCEmulator
        $RIDMastTextBox.Text = $adDomain.RIDMaster
        $InfMastTextBox.Text = $adDomain.InfrastructureMaster
        $SchemaMastTextBox.Text = $adForest.SchemaMaster
        $NamingMastTextBox.Text = $adForest.DomainNamingMaster
        ForEach ($Server in $adForest.GlobalCatalogs) {
            $GlobalCatListBox.Items.Add($Server)
        }
        ForEach ($Site in $adForest.Sites) {
            $ForestSitesListBox.Items.Add($Site)
        }
    })

    $StartOverButton.Add_Click({
        $Form.Controls.Add($SelectDomainButton)
        $Form.Controls.Remove($StartOverButton)
        $GlobalCatListBoxItems = $GlobalCatListBox.Items
        while ($GlobalCatListBoxItems -ne $Null) {
            $GlobalCatListBox.Items.Remove($GlobalCatListBoxItems[0])
            $GlobalCatListBoxItems = $GlobalCatListBox.Items
        }
        $ForestSitesListBoxItems = $ForestSitesListBox.Items
        while ($ForestSitesListBoxItems -ne $Null) {
            $ForestSitesListBox.Items.Remove($ForestSitesListBoxItems[0])
            $ForestSitesListBoxItems = $ForestSitesListBox.Items
        }
        $ForestModeTextBox.Clear()
        $DomainModeTextBox.Clear()
        $UPNSuffTextBox.Clear()
        $PDCeTextBox.Clear()
        $RIDMastTextBox.Clear()
        $InfMastTextBox.Clear()
        $SchemaMastTextBox.Clear()
        $NamingMastTextBox.Clear()
    })


    ### Launch Form
    $Form.ShowDialog()
}



MakeForm
