# This script compares the two source scripts to show you what's different

# Set the script directory whether it's ran from ISE or not
if ($PSScriptRoot){$scriptdir = $PSScriptRoot}
else{$scriptdir = Split-Path -Path $psISE.CurrentFile.FullPath}

# Source PS1 Files
$SourceTemp = "$Scriptdir\SourceScripts\PermVins.ps1"
$SourcePerm = "$Scriptdir\SourceScripts\TempVins.ps1"

# Pull in the source files
$TempData = gc $SourceTemp
$PermData = gc $SourcePerm

# Compare the two
Compare-Object $TempData $PermData