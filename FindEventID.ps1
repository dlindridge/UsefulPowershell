<#
	.SYNOPSIS
	Looks through the event log to find a specific Event ID number and outputs results to file choice (CSV or TXT).

	.PARAMETER Path
	Directory to place output file(s) to.

	.PARAMETER EventID
	Event ID number to search for.
	
	.PARAMETER CSV
	Switch - Designates output should go to a CSV file. This is good for determining when and who triggered an event.
	
	.PARAMETER TXT
	Switch - Designates output should go to a TXT file. This is good for determining WHY an event was triggered.

	.DESCRIPTION
	Usage: .\FindEventID.ps1 -Path D:\Some\Directory -EventID 1074 (-CSV) (-TXT)
#>
#################################################
<#
Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
https://github.com/dlindridge/UsefulPowershell
Created: Feb 03, 2020
Modified: Feb 03, 2020
#>
#################################################

Param (
    [Parameter(Mandatory=$True)]
    $Path,
    [Parameter(Mandatory=$True)]
    $EventID,
	[switch]$CSV,
	[switch]$TXT
)

# Get local computer name
$CompName = $ENV:computername

# Format output files
$Date = (Get-Date).ToString("yyyyMMMdd-HHmmss")
$CSVOut = "$CompName-$Date-EID$EventID.csv"
$TXTOut = "$CompName-$Date-EID$EventID.txt"
$CSVOutFile = Join-Path -Path $Path -ChildPath $CSVOut
$TXTOutFile = Join-Path -Path $Path -ChildPath $TXTOut

# Generate CSV output file
If ($CSV -eq $True) {
	Get-EventLog -LogName System | Where-Object {$_.EventID -eq $EventID} | Export-CSV $CSVOutFile -NoTypeInformation
	}

# Generate Format List output as TXT file
If ($TXT -eq $True) {
	Get-EventLog -LogName System | Where-Object {$_.EventID -eq $EventID} | fl | Out-File $TXTOutFile
	}
