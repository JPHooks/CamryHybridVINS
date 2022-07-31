# This script allows you to edit 1 temp and 1 perm PS1 file with any changes you wish to make
# Drop those in the script root directory and run it and it will generate 14 total scripts, 1 for each model.
# Output goes into the Output directory

# Set the script directory whether it's ran from ISE or not
if ($PSScriptRoot){$scriptdir = $PSScriptRoot}
else{$scriptdir = Split-Path -Path $psISE.CurrentFile.FullPath}

# Source PS1 Files
$SourceTemp = "$Scriptdir\SourceScripts\TempVins.ps1"
$SourcePerm = "$Scriptdir\SourceScripts\PermVins.ps1"

# Output Directory
$OutputDir = "$Scriptdir\Output"
if (!(Test-Path "$OutputDir")){
    New-Item -ItemType Directory "$Scriptdir\Output" -Force | Out-Null
    }

# Pull in the source files
$TempData = gc $SourceTemp
$PermData = gc $SourcePerm

# Get Possible Prefixes from script
$StartLog = $false
$Prefixes = foreach ($line in $TempData){
    if ($line -like "# Possible Prefixes*"){
        $StartLog = $True
        }
    if ($StartLog){
        $line
        }
    if ($line -like ""){
        $StartLog = $False
        }
    }

# Remove first and last string from Prefix array
$Prefixes = $Prefixes[1..($Prefixes.Count-2)]

# Go through each Prefix, edit the source file, and output to the correct filename
foreach ($Prefix in $Prefixes){

    # Empty Arrays to store the new PowerShell code
    $NewTemp = $NewPerm = @()

    # Parse names out of the Prefix for file naming and editing
    $FileName = ($Prefix -split ' ')[0] -Replace '\$'
    $FileVinCode = (($Prefix -split ' "')[1])[3]

    # Change the $VinPrefix variable to desired prefix
    foreach ($TempLine in $TempData){
        if ($TempLine -like '$VinPrefix*'){
            $NewTemp += '$VinPrefix = $' + $FileName
            }
        else{
            $NewTemp += $TempLine
            }
        }
    foreach ($PermLine in $PermData){
        if ($PermLine -like '$VinPrefix*'){
            $NewPerm += '$VinPrefix = $' + $FileName
            }
        else{
            $NewPerm += $PermLine
            }
        }

    # Export the new code to the new filename
    $NewTemp | Out-File "$OutputDir\TempVins - $FileName $FileVinCode.ps1"
    $NewPerm | Out-File "$OutputDir\PermVins - $FileName $FileVinCode.ps1"

    }