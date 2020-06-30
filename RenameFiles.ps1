    <#
        .SYNOPSIS
        Find files in a root directory (Recursivly) and rename or copy them to the same relative directory

        .PARAMETER Root
        Root directory to start the search in.

        .PARAMETER Source
        Source File Name.

        .PARAMETER Destination
        Destination File Name.

        .PARAMETER Action
        "Copy" or "Rename" accepted parameters

        .DESCRIPTION
        .\RenameFiles.ps1 -Root D:\Sort -Source folder.jpg -Destination poster.jpg -Action Copy
    #>
#################################################
<#
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    https://github.com/dlindridge/UsefulPowershell
    Created: September 28, 2019
    Modified: October 22, 2019
#>
#################################################

    Param (
        [Parameter(Mandatory=$True)]
        $Root,
        [Parameter(Mandatory=$True)]
        $Source,
        [Parameter(Mandatory=$True)]
	$Destination,
        [Parameter(Mandatory=$False)]
	$Action
    )

### Parameter Validation ########################
if ($Action -ne "Copy" -AND $Action -ne "Rename") {
	Write-Host -ForegroundColor Red "You must include the -Action parameter with either Copy or Rename"
	Break
}

### Rename Action ###############################
if ($Action -eq "Rename") {
	# Get files in or below the root directory to process
	$files = Get-ChildItem -Path $Root -Filter $Source -Recurse
	# Process each file
	ForEach ($file In $files) {
		# Prepare the destination file name
		$destFile = Get-ChildItem -Path $file.Directory -Filter $Destination
		# Check if the destination file exists in the final destination
		if (-NOT $destFile.Name -eq $Destination) {
			Write-Host -ForegroundColor Green 'Moving' $file.FullName 'to' $Destination
			# Prepare the final destination file name
			$destFinal = Join-Path $file.Directory -ChildPath $Destination
			# Rename the file
			Move-Item -Path $file.FullName -Destination $destFinal
		}
		Else {
#			Write-Host -ForegroundColor Red 'Something is wrong in' $file.Directory
		}
	}
}

### Copy Action #################################
if ($Action -eq "Copy") {
	# Get files in or below the root directory to process
	$files = Get-ChildItem -Path $Root -Filter $Source -Recurse
	# Process each file
	ForEach ($file In $files) {
		# Prepare the destination file name
		$destFile = Get-ChildItem -Path $file.Directory -Filter $Destination
		# Check if the destination file exists in the final destination
		if (-NOT $destFile.Name -eq $Destination) {
			Write-Host -ForegroundColor Green 'Copying' $file.FullName 'to' $Destination
			# Prepare the final destination file name
			$destFinal = Join-Path $file.Directory -ChildPath $Destination
			# Copy the file
			Copy-Item -Path $file.FullName -Destination $destFinal
		}
		Else {
#			Write-Host -ForegroundColor Red 'Something is wrong in' $file.Directory
		}
	}
}


