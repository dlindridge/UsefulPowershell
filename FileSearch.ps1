<#
	.SYNOPSIS
	Recursively searches directory to find a file or folder either by specific name or by pattern.
	
	.PARAMETER Root
	Root directory to search.

	.PARAMETER Name
	File Name or pattern to search for.
	
	.PARAMETER Type
	Search for 'Files' or 'Folders' (optional). Unspecified will search for both.
	
	.PARAMETER Output
	Specifies folder and filename to place output CSV. (Optional)
	Without this parameter output will go to screen.

	.DESCRIPTION
	Usage: .\FileSearch.ps1 -Root D:\Some\Directory -Name *text* (-Type Files) (-Outfile C:\Users\joe.somebody\Desktop\output.csv)
#>
#################################################
<#
Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
https://github.com/dlindridge/UsefulPowershell
Created: February 6, 2020
Modified: February 6, 2020
#>
#################################################

Param (
    [Parameter(Mandatory=$True)]
    $Root,
    [Parameter(Mandatory=$True)]
	$Name,
	[ValidateSet("Files","Folders")]
	[String]$Type,
	$Output
)


If ($Output -eq $Null) {
	If ($Type -eq "Files") {
		Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { !$_.PSIsContainer }
		}
	ElseIf ($Type -eq "Folders") {
		Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }
		}
	Else {
		Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue
	}
}

If ($Output -ne $Null) {
	If ($Type -eq "Files") {
	Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { !$_.PSIsContainer } | Export-CSV $Output -NoTypeInformation
	}
	ElseIf ($Type -eq "Folders") {
	Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer } | Export-CSV $Output -NoTypeInformation
	}
	Else {
		Get-ChildItem -Path $Root -Include $Name -Recurse -Force -ErrorAction SilentlyContinue | Export-CSV $Output -NoTypeInformation
	}
}

