<#
    .SYNOPSIS
    Recursively searches defined directory to delete files and empty folders older than a specified number of days.
    Optionally leaves a warning file behind about the temporary nature of the directory.

    .PARAMETER Warning
    Will the script leave a warning file about the temporary nature of the directory. Default is YES.

    .PARAMETER Days
    Number of days that a file can be left in this location. Default is 15.

    .PARAMETER Path
    Directory to search and remove files and folders from.

    .DESCRIPTION
    Usage: .\DeleteOlderThan.ps1 -Path D:\Some\Directory (-Warning NO) (-Days 20)
#>
#################################################
<#
Delete files from a directory older than a specified number of days.
Replaces FORFILES DOS command (depricated).

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: May 17, 2015
Modified: October 24, 2019
#>
#################################################

Param (
    $Warning = "YES",
    $Days = 15,
    [Parameter(Mandatory=$True)]
    $Path
)

$limitDays = $Days
$Warning = $Warning.ToUpper()
$limit = (Get-Date).AddDays(-$limitDays)

# Delete files older than the $limit.
Get-ChildItem -Path $path -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | Remove-Item -Force

# Delete any empty directories left behind after deleting the old files.
Get-ChildItem -Path $path -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Force -Recurse

If ($Warning -eq "YES") {
	$WarningFile = Join-Path -Path $path -ChildPath "!--ATTENTION - PLEASE READ--!.txt"
	Get-ChildItem -Path $WarningFile | Remove-Item -Force
	Add-Content $WarningFile "This directory is for temporary storage of files for transfer or use elsewhere." | Wait-Job
	Add-Content $WarningFile "" | Wait-Job
	Add-Content $WarningFile "Files left in this directory will be automatically deleted after $limitDays days." | Wait-Job
	Add-Content $WarningFile "" | Wait-Job
	Add-Content $WarningFile "This directory is *NOT* backed up and is considered temporary storage!" | Wait-Job
	Get-ChildItem -Path $path -File "!--ATTENTION - PLEASE READ--!.txt" | Set-ItemProperty -Name IsReadOnly -Value $True
}
