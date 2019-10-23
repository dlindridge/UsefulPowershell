#################################################
<#
Delete files from a directory older than a specified number of days.
Replaces FORFILES DOS command (depricated).

Author: Derek Lindridge
https://www.linkedin.com/in/dereklindridge/
Created: May 17, 2015
Modified: October 23, 2019
#>
#################################################

$limitDays = 15
$Warning = "YES" # Leave a read-only file behind that this directory is cleaned regularly?

$limit = (Get-Date).AddDays(-$limitDays)
$path = "D:\SomeDirectory"
$Warning = $Warning.ToUpper()

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