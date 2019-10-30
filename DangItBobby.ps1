<#
    .SYNOPSIS
    This is a GUI version of the original script by bluesole (https://bluesoul.me). I've done a little cleanup, but most of my
    effort was put into getting it into a Form. It looks for a user logged into an available computer on your network and
    allows you to disable active NICs. Useful if you are getting hit by a Cryptolocker attack.

    .DESCRIPTION
    Usage: Create a shortcut pointing to (powershell.exe -file "C:\Scripts\UserStuff\DangItBobby.ps1")
    https://mcpmag.com/articles/2014/02/18/progress-bar-to-a-graphical-status-box.aspx
#>
#################################################
<#
    TO-DO: Put an optional message on the target machine about what is happening

    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: October 25, 2019
    Modified: October 29, 2019
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

ShowConsole


### Cancel Form #################################
Function CancelForm {
    $Form.Close()
    $Form.Dispose()
}


### Clear Form #################################
Function ClearForm {
    $UserNameTextBox.Text = "SamAccountName"
    $ProgressBar1.Value = 0
    $ComputerListBox.Items.Clear()
    $ComputerListBox.Enabled = $False
    Clear-Variable -Name $Query
    Clear-Variable -Name $UserName
    Clear-Variable -Name $UserCheck
    Clear-Variable -Name $CompCount
    Clear-Variable -Name $Computers
    Clear-Variable -Name $Hits
    $Form.Enabled = $True
}


### Build Form ##################################
Function MakeForm {
    ### Set form parameters
    Add-Type -AssemblyName System.Windows.Forms
    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Find Logged In User and Disable Active NICs"
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
    $PBi = 0
    If ($IconImage -ne "" -OR $IconImage -ne $Null) { $Form.Icon = New-Object system.drawing.icon ($Icon) }
    If ($BackgroundImage -ne "" -OR $BackgroundImage -ne $Null) { $Form.BackgroundImage = [system.drawing.image]::FromFile($Background) }
    $ObjFont = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Regular)
    $ObjFontBold = New-Object System.Drawing.Font("Microsoft Sans Serif",10,[System.Drawing.FontStyle]::Bold)
    $Credential = Get-Credential -Message "Please provide credentials of a user with admin permission on the target machine.`n Example: DOMAIN\admin.user"


    ### Set and add form objects
    $adDomainsDropDownBoxLabel = New-Object System.Windows.Forms.Label
    $adDomainsDropDownBoxLabel.Left = 2
    $adDomainsDropDownBoxLabel.Top = 10
    $adDomainsDropDownBoxLabel.Size = New-Object System.Drawing.Size(75,25)
    $adDomainsDropDownBoxLabel.TextAlign = "MiddleLeft"
    $adDomainsDropDownBoxLabel.Text = "Domain:"
    $adDomainsDropDownBoxLabel.BackColor = "Transparent"
    $Form.Controls.Add($adDomainsDropDownBoxLabel)

    [array]$adDomains = (Get-ADForest).Domains
    $adDomainsDropDownBox = New-Object System.Windows.Forms.ComboBox
    $adDomainsDropDownBox.Left = 80
    $adDomainsDropDownBox.Top = 10
    $adDomainsDropDownBox.Location = New-Object System.Drawing.Size(80,10)
    $adDomainsDropDownBox.Size = New-Object System.Drawing.Size(222,25)
    $adDomainsDropDownBox.Font = $ObjFont
    $adDomainsDropDownBox.DropDownStyle = "DropDownList"
    ForEach ($adDomain in $adDomains) { $adDomainsDropDownBox.Items.Add($adDomain) }
    $adDomainsDropDownBox.SelectedIndex = 0
    $Form.Controls.Add($adDomainsDropDownBox)

    $UsernameTextLabel = New-Object System.Windows.Forms.Label
    $UsernameTextLabel.Left = 2
    $UsernameTextLabel.Top = 40
    $UsernameTextLabel.Size = New-Object System.Drawing.Size(75,25)
    $UsernameTextLabel.TextAlign = "MiddleLeft"
    $UsernameTextLabel.Text = "Username:"
    $UsernameTextLabel.BackColor = "Transparent"
    $Form.Controls.Add($UsernameTextLabel)

    $UserNameTextBox = New-Object System.Windows.Forms.TextBox
    $UserNameTextBox.Left = 80
    $UserNameTextBox.Top = 40
    $UserNameTextBox.Size = New-Object System.Drawing.Size(222,25)
    $UserNameTextBox.Font = $ObjFont
    $UserNameTextBox.Text = "SamAccountName"
    $Form.Controls.Add($UserNameTextBox)

    $CompPatternTextLabel = New-Object System.Windows.Forms.Label
    $CompPatternTextLabel.Left = 2
    $CompPatternTextLabel.Top = 70
    $CompPatternTextLabel.Size = New-Object System.Drawing.Size(300,25)
    $CompPatternTextLabel.TextAlign = "MiddleLeft"
    $CompPatternTextLabel.Text = "Computer Name Pattern:"
    $CompPatternTextLabel.BackColor = "Transparent"
    $Form.Controls.Add($CompPatternTextLabel)

    $CompPatternTextBox = New-Object System.Windows.Forms.TextBox
    $CompPatternTextBox.Left = 30
    $CompPatternTextBox.Top = 100
    $CompPatternTextBox.Size = New-Object System.Drawing.Size(273,25)
    $CompPatternTextBox.Font = $ObjFont
    $CompPatternTextBox.Text = "*"
    $Form.Controls.Add($CompPatternTextBox)

    $ComputerListBox = New-Object System.Windows.Forms.ListBox 
    $ComputerListBox.Left = 2
    $ComputerListBox.Top = 130
    $ComputerListBox.Size = New-Object System.Drawing.Size(300,360) 
    $ComputerListBox.Font = $ObjFont
    $ComputerListBox.Sorted = $True
    $ComputerListBox.Enabled = $False
    $ComputerListBox.SelectionMode = "One" # MultiExtended, MultiSimple, None, One
    $Form.Controls.Add($ComputerListBox)

    $ProgressBar1 = New-Object System.Windows.Forms.ProgressBar
    $ProgressBar1.Left = 310
    $ProgressBar1.Top = 10
    $ProgressBar1.Size = New-Object System.Drawing.Size(200,25)
    $ProgressBar1.Text = "Checking Computers"
    $ProgressBar1.Value = 0
    $ProgressBar1.Step = 1
    $ProgressBar1.Style = "Continuous"
    $PBCompCount = Get-ADComputer -Filter {Enabled -eq 'true'}
    $ProgressBar1.Maximum = $PBCompCount.Count

    $PBTextLabel = New-Object System.Windows.Forms.Label
    $PBTextLabel.Left = 310
    $PBTextLabel.Top = 35
    $PBTextLabel.Size = New-Object System.Drawing.Size(200,25)
    $PBTextLabel.TextAlign = "MiddleCenter"
    $PBTextLabel.Text = "Working..."
    $PBTextLabel.BackColor = "Transparent"
    $PBTextLabel.Font = $ObjFontBold

    $UserSelectButton = New-Object System.Windows.Forms.Button 
    $UserSelectButton.Left = 340
    $UserSelectButton.Top = 150
    $UserSelectButton.Size = New-Object System.Drawing.Size(111,75)
    $UserSelectButton.Text = "Find User"
    $Form.Controls.Add($UserSelectButton)

    $ComputerSelectButton = New-Object System.Windows.Forms.Button 
    $ComputerSelectButton.Left = 340
    $ComputerSelectButton.Top = 230
    $ComputerSelectButton.Size = New-Object System.Drawing.Size(111,75)
    $ComputerSelectButton.Text = "Select Computer"
    $ComputerSelectButton.Enabled = $False
    $Form.Controls.Add($ComputerSelectButton)

    $CancelButton = New-Object System.Windows.Forms.Button 
    $CancelButton.Left = 340
    $CancelButton.Top = 440
    $CancelButton.Size = New-Object System.Drawing.Size(180,50)
    $CancelButton.Text = "Nevermind"
    $CancelButton.Add_Click({ CancelForm })
    $Form.Controls.Add($CancelButton)

    $UserSelectButton.Add_Click({
        $ProgressBar1.Value = 0
        $Prefix = $CompPatternTextBox.Text
        $Form.Controls.Add($ProgressBar1)
        If ($Computers -ne $Null) { Clear-Variable -Name $Computers }
        $ComputerListBox.Items.Clear()
        $Username = $UserNameTextBox.Text
        $Domain = $adDomainsDropDownBox.SelectedItem
        $UserCheck = Get-ADUser -Server $Domain -Identity $Username -ErrorAction SilentlyContinue
        If ($UserCheck -eq $Null -AND $Username -ne "" -AND $Username -ne "SamAccountName"){
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageboxTitle = "User Not Found"
            $MessageboxBody = "Username $Username cannot be found in $Domain.`nPlease verify username and domain are correct."
            $MessageIcon = [System.Windows.MessageBoxImage]::Error
            [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
            ClearForm
        }
        ElseIf ($Username -eq "" -OR $Username -eq $Null -OR $Username -eq "SamAccountName") {
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageboxTitle = "Username Blank"
            $MessageboxBody = "Please pick a valid username to search for."
            $MessageIcon = [System.Windows.MessageBoxImage]::Error
            [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
            ClearForm
        }
        Else { FindComputers }
    
    })

    $ComputerSelectButton.Add_Click({
        $SelectedComputers = $ComputerListBox.SelectedItems
        DisableComputers
    })


    Function FindComputers {
        $Form.Enabled = $False
        $Form.Controls.Add($PBTextLabel)
        #Start search
        If ($Prefix -eq "") { $Prefix = "*" }
        $Computers = Get-ADComputer -Filter {Enabled -eq 'true' -and SamAccountName -like $Prefix}
        $CompCount = $Computers.Count
        $ProgressBar1.Maximum = $Computers.Count

        #Create mutable array for catching computers that match.
        $Global:HitsX = @()
        $Global:Hits = {$HitsX}.Invoke()

        #Start main foreach loop, search processes on all computers
        ForEach ($Comp in $Computers) {
            $Computer = $Comp.Name
            $progressbar1.Increment(1) 
            $Form.Refresh()
            $Reply = $null
            $Reply = Test-Connection $Computer -count 1 -quiet
            If ($Reply -eq 'True') {
                If ($Computer -eq $env:COMPUTERNAME) {
                    #Get explorer.exe processes without credentials parameter if the query is executed on the localhost
                    $proc = Get-WmiObject win32_process -ErrorAction SilentlyContinue -Computer $Computer -Filter "Name = 'explorer.exe'"
                }
                Else {
                    #Get explorer.exe processes with credentials for remote hosts
                    $proc = Get-WmiObject win32_process -ErrorAction SilentlyContinue -Credential $Credential -Computer $Computer -Filter "Name = 'explorer.exe'"
                }			
                    #If $proc is empty return msg else search collection of processes for username
                If ([string]::IsNullOrEmpty($proc)) {
                    $progress++
                }
                Else {	
                    $progress++			
                    ForEach ($p in $proc) {				
                        $temp = ($p.GetOwner()).User
                        If ($temp -eq $Username) {
                            $Global:Hits.Add($Computer)
                        }
                    }
                }	
            }
        }
        PopulateComputers
    }


    Function PopulateComputers {
        $Hits | Sort-Object -Property @{Expression={$_.Trim()}} -Unique
        If ($Hits -eq $Null -OR $Hits -lt 1) {
            Add-Type -AssemblyName PresentationCore,PresentationFramework
            $ButtonType = [System.Windows.MessageBoxButton]::OK
            $MessageboxTitle = "User Not Logged In"
            $MessageboxBody = "Unable to locate $Username logged in to any available computer in $Domain."
            $MessageIcon = [System.Windows.MessageBoxImage]::Error
            [System.Windows.MessageBox]::Show($MessageboxBody,$MessageboxTitle,$ButtonType,$MessageIcon)
            ClearForm
        }
        ElseIf ($Hits -ge 1) {
            ForEach ($Hit in $Hits) {
                $ComputerListBox.Items.Add($Hit)
            }
        }
        $Form.Controls.Remove($PBTextLabel)
        $Form.Enabled = $True
        $ComputerListBox.Enabled = $True
        $ComputerSelectButton.Enabled = $True
    }


    Function DisableComputers {
        $Form.Controls.Remove($ProgressBar1)
        $Form.Enabled = $False
        $Query = "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'True'"
        $NICs = Get-WMIObject -ComputerName $SelectedComputers -Credential $Credential -Query $Query | Select-Object Index
        ForEach ($NIC in $NICs) {
            $WMI = Get-WMIObject -Class Win32_NetworkAdapter -filter "Index LIKE $($NIC.Index)" -ComputerName $SelectedComputers -Credential $Credential
            $WMI.Disable()
        }
    ClearForm
    }


    ### Launch Form
    $Form.ShowDialog()
}

### Start Form As Function ######################
MakeForm
