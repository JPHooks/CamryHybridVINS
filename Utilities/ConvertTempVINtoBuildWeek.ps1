# Clear Screen (mostly for testing)
cls

# Get the VIN
Write-Host "Paste in the VIN you wish to look up build date on." -ForegroundColor Cyan
[String]$VIN = Read-Host

# Validate temp vin
if ($VIN[13] -match "\d"){
    Write-Host "Warning: VIN does not appear to be a temp vin.  There is not a letter in the serial number.`n" -ForegroundColor Yellow
    }

# Get week number from VIN
[String]$WeekStr = $VIN[11] + $VIN[12]
[Int]$WeekNum = $WeekStr

# Set the start date
[datetime]$startd = "2022-01-01"

# Calculate number of days to add from start date
$Days = $WeekNum * 7
$BuildWeek = ($startd.AddDays($Days)).ToString('yyyy-MM-dd')

# Write build week to screen
Write-Host "VIN shows build date should be the week of $BuildWeek" -ForegroundColor Green