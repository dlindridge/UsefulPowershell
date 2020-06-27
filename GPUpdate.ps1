$PSTools = "C:\Scripts\PSTools\psexec.exe" # Full path to psexec on the source system

#################################################
Import-Module ActiveDirectory

# PrimaryGroupId "516" is Domain Controllers
$Computers = Get-ADComputer -Filter {(Enabled -eq $True) -AND (PrimaryGroupId -ne "516") -AND (operatingSystem -Like "*Windows*")} -Properties * | Select-Object Name | Sort-Object Name

ForEach ($Computer in $Computers) {
	$ComputerName = $Computer.Name
	Write-Host -ForegroundColor Yellow "Is $ComputerName online? " -NoNewLine
	$Online = Test-Connection $Computer.Name -Count 1 -Quiet
	If ($Online -eq 'True') {
		Write-Host -ForegroundColor Green "Yes - Starting GPUpdate"
		$CompName = "\\" + $Computer.Name
		$CMD = "CMD"
		[Array]$cmdParams = "/c", $PSTools, $CompName, "/d", "cmd", "/c", "gpupdate /force"
		& $CMD $cmdParams
	}
	Else {
		Write-Host -ForegroundColor Red "No - Skipped"
	}
	Write-Host
}

Start-Sleep -Seconds 5
