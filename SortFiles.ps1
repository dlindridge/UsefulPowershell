    <#
    .SYNOPSIS
    Sort all files in a root directory to new folders in the same directory based on the original file name.

    .PARAMETER Root
    Root directory to start the search in.

    .USAGE
    .\SortFiles.ps1 -Root D:\Sort
    #>
#################################################
<#
    Author: Derek Lindridge
    https://www.linkedin.com/in/dereklindridge/
    Created: September 28, 2019
    Modified: October 22, 2019
#>
#################################################

Param (
    [Parameter(Mandatory=$True)]
    $Root
)


# Get files in the root directory
$files = Get-ChildItem -Path $Root | Where-Object { -NOT $_.PsIsContainer }

# Process each file
ForEach ($file In $files) {
    # Name of the individual file
    $fileName = $file.BaseName
    # Prepare the name of the new directory
    $folder = Join-Path -Path $Root -ChildPath $fileName
    # What to do if the folder exists
    if (Test-Path $folder) {
        $destFolder = Join-Path -Path $folder -ChildPath $file.Name
        Move-Item $file.FullName $destFolder
    }
    # What to do if the folder does not exist
    if (-NOT (Test-Path $folder)) {
        $makeFolder = New-Item -Path $Root -Type Directory -Name $fileName.Substring(0) -ErrorAction SilentlyContinue
        Move-Item $file.FullName $makeFolder.FullName
    }
}
