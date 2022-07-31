# In order to run multiple threads, you can clone this script for each prefix and then set which one the script will run as
# This lets you run 7 powershell.exe processes (1 for each hybrid trim) at once or only run it for the trim you care about

# Possible Prefixes (The hyphen gets replaced by checksum function)
$LE1 = "4T1C31AK-NU"
$LE2 = "4T1H31AK-NU"
$SE1 = "4T1G31AK-NU"
$SE2 = "4T1S31AK-NU"
$XSE = "4T1K31AK-NU"
$XLE = "4T1F31AK-NU"
$SENightShade = "4T1T31AK-NU"

### Define Variables

# Set the Trim you care about here or if you are copying this script for each trim which one will this run as
$VinPrefix = $XSE

# Which serial prefixes do you want? Temp Vins start with the week they estimate it to be built.  As of writing this script it's week 24-26.
# This is likely going to change in the future but should be current as of July 2022.
$VinPreSerial = @("24", "25", "26", "27")

# Possible letters that appear after the first 2 numbers of the 6 digit serial
$VinMidSerial = @("A", "B", "C", "D", "E", "F")

# Set this to True or False.  False will give you a text file with the VINs.  True will hit Toyota's API to search for them and give a CSV.
$HitToyotaAPI = $True

# Define Options you care about, if any exist the flag will be Yes
$CarOptions = @("Driver Assist")

#############################################################################
### Don't edit anything below this line unless you know what you're doing ###
#############################################################################

# Function to provide checksum character.  This is a unique number or X based on last 6 digits of VIN.
# This was directly taken from https://github.com/rgoldfinger/rav4_scraping/blob/main/util/getCheckCode.ts
# I used https://www.typescriptlang.org/play to understand TypeScript enough to convert it to PowerShell
Function GetChecksumChar {
    [cmdletbinding()]
    Param ([string]$Vin)

    $values = @(1, 2, 3, 4, 5, 6, 7, 8, 0, 1, 2, 3, 4, 5, 0, 7, 0, 9, 2, 3, 4, 5, 6, 7, 8, 9)
    $weights = @(8, 7, 6, 5, 4, 3, 2, 10, 0, 9, 8, 7, 6, 5, 4, 3, 2)

    $Sum = 0
    for ($i=0; $i -lt 17; $i++){
        $c = $Vin[$i]
        if ($c -match "[A-Z]"){
            $Value = $Values[[byte][char]$c - [byte][char]'A']
            }
        elseif ($c -match "\d"){
            $Value = $c
            }
        [Int]$Val = $Value.ToString()
        $Sum = $Sum + $weights[$i] * $Val
        }

    $Sum = $Sum % 11

    if ($Sum -eq 10){
        Return "X"
        }

    else{
        Return $Sum.ToString()
        }
    }

# Set script start time so can output how long it took to run
$ScriptStart = Get-Date

# Set the script directory whether it's ran from ISE or not
if ($PSScriptRoot){$scriptdir = $PSScriptRoot}
else{$scriptdir = Split-Path -Path $psISE.CurrentFile.FullPath}

# If Logs, Output, and Database folders don't exist, create them
if (!(Test-Path "$Scriptdir\Output")){
    New-Item -ItemType Directory "$Scriptdir\Output" -Force | Out-Null
    }

if (!(Test-Path "$Scriptdir\Logs")){
    New-Item -ItemType Directory "$Scriptdir\Logs" -Force | Out-Null
    }

if (!(Test-Path "$Scriptdir\Database")){
    New-Item -ItemType Directory "$Scriptdir\Database" -Force | Out-Null
    }

# Set human readable model from VinPrefix for logging
Switch ($VinPrefix[3]){
    C { $VinModel = "LE" }
    H { $VinModel = "LE" }
    G { $VinModel = "SE" }
    S { $VinModel = "SE" }
    K { $VinModel = "XSE" }
    F { $VinModel = "XLE" }
    T { $VinModel = "SE NightShade" }
    }

# Date to String for Use in CSV Export
$DateToday = (Get-Date).ToString("MM/dd/yyyy")

# If you want to hit the Toyota API it will do web queries
if ($HitToyotaAPI){
    
    # Create Hash Table for dealer category
    $DealerCategory = @{
        G = "Ground"
        F = "Freight"
        A = "Allocated"
        }

    # Create Hash Table for dealer code (Dealers.csv must be in script root directory!)
    $DealerInfo = Import-CSV "$Scriptdir\Dealers.csv"
    $DealerInfo = $DealerInfo | Sort ID -Unique
    $DealerCD = @{}
    foreach ($Dealer in $DealerInfo){
        $PaddedDealerID = $(([string]$Dealer.ID).PadLeft(5,'0'))
        $DealerCD.Add($PaddedDealerID, @($Dealer.Name, $Dealer.Address, $Dealer.City, $Dealer.State, $Dealer.Zip, $Dealer.Link))
        }

    # Build out session data for the web queries
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $session.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36"            
    $sessionheaders = @{"origin"="https://guest.dealer.toyota.com";"referer"="https://guest.dealer.toyota.com/"}

    # Check for database file and import it into a hash table if it exists (Used for First Seen column)
    if (Test-Path "$Scriptdir\Database\$($VinModel)_$($VinPrefix[3])_Temp_VinDB.csv"){
        $VinDBData = Import-Csv "$Scriptdir\Database\$($VinModel)_$($VinPrefix[3])_Temp_VinDB.csv"
        $VinDBData = $VinDBData | Sort Vin -Unique
        $VinDB = @{}
        foreach ($dbitem in $VinDBData){
            $VinDB.Add($dbitem.Vin, $dbitem.FirstSeen)
            }
        }

    # Empty Array for Errors and new DB
    $APIErrors = $NewVinDB = @()

    $VINS = foreach ($PreSerial in $VinPreSerial){
        foreach ($MidSerial in $VinMidSerial){
            for ($i=0;$i-le 999;$i++){
                $vin = $VinPrefix + $PreSerial + $MidSerial + $(([string]$i).PadLeft(3,'0'))
                $cleanvin = $vin -replace '-', (GetChecksumChar $vin)

                try{$Webdata = Invoke-WebRequest -Uri "https://api.rti.toyota.com/marketplace-inventory/vehicles/$($cleanvin)?isVspec=true" -WebSession $session -Headers $sessionheaders}
                catch{
                    # Log any errors that are not 404 in case Toyota starts throttling or something
                    if ($_.Exception.Message -notmatch '404'){
                        $obj = new-object psobject
                        $obj | add-member noteproperty Date (Get-Date).ToString()
                        $obj | add-member noteproperty Model $VinModel
                        $obj | add-member noteproperty Prefix $VinPrefix
                        $obj | add-member noteproperty VIN $cleanvin
                        $obj | add-member noteproperty APIError $_.Exception.Message
                        $APIErrors += $obj
                        }
                    # Clear the Webdata Variable if it exists so the row doesn't get added to the main export
                    if ($Webdata){
                        Clear-Variable webdata
                        }
                    }

                if ($Webdata.StatusCode -eq 200){
                    $WebJSON = $Webdata.Content | ConvertFrom-Json

                    # Convert DealerCategory to Human Readable Text
                    if ($DealerCategory."$($WebJSON.dealerCategory)"){
                        $DCategory = $DealerCategory."$($WebJSON.dealerCategory)"
                        }
                    else{
                        $DCategory = $WebJSON.dealerCategory
                        }

                    # Convert DealerCD to Human Readable Text
                    if ($DealerCD."$($WebJSON.dealerCd)"){
                        $DealerName = ($DealerCD."$($WebJSON.dealerCd)")[0]
                        $DealerCity = ($DealerCD."$($WebJSON.dealerCd)")[2]
                        $DealerState = ($DealerCD."$($WebJSON.dealerCd)")[3]
                        $DealerZip = ($DealerCD."$($WebJSON.dealerCd)")[4]
                        }
                    else{
                        # If Dealer Code isn't found in the CSV, put the link instead so it can be manually checked and clear out other Dealer variables
                        $DealerName = "https://www.toyota.com/dealers/dealer/$($WebJSON.dealerCd)"
                        if ($dealercity){Clear-Variable dealercity, dealerstate, dealerzip}
                        }

                    # Convert TempVin serial to build week
                    if ($WebJSON.isTempVin){
                        [String]$WeekStr = $WebJSON.Vin[11] + $WebJSON.Vin[12]
                        [Int]$WeekNum = $WeekStr
                        [datetime]$WeekStart = "2022-01-01"
                        [Int]$Days = $WeekNum * 7
                        [String]$BuildWeek = ($WeekStart.AddDays($Days)).ToString('yyyy-MM-dd')
                        }
                    else{
                        if ($BuildWeek){Clear-Variable BuildWeek}
                        }

                    # Provide Car link and dealer link data
                    $CarLink = "https://guest.dealer.toyota.com/v-spec/$($WebJSON.vin)/detail"
                    $DealerLink = "https://www.toyota.com/dealers/dealer/$($WebJSON.dealerCd)"

                    # Reference the Vin database to determine first seen date
                    if ($VinDB."$($WebJSON.Vin)"){
                        $DateFirst = $VinDB."$($WebJSON.Vin)"
                        
                        $obj = new-object psobject
                        $obj | add-member noteproperty VIN $WebJSON.Vin
                        $obj | add-member noteproperty FirstSeen $DateFirst
                        $NewVinDB += $obj
                        }
                    
                    else{
                        $DateFirst = $DateToday
                        $obj = new-object psobject
                        $obj | add-member noteproperty VIN $WebJSON.Vin
                        $obj | add-member noteproperty FirstSeen $DateFirst
                        $NewVinDB += $obj
                        }

                    # Determine if there are any options you care about
                    $OptionsFlag = $false
                    foreach ($optionitem in $WebJSON.options.marketingName){
                        foreach ($CarOption in $CarOptions){
                            if ($optionitem -match $CarOption){
                                $OptionsFlag = $true
                                }                        
                            }
                        }

                    $obj = new-object psobject
                    $obj | add-member noteproperty FirstSeen $DateFirst
                    $obj | add-member noteproperty LastUpdated $DateToday
                    $obj | add-member noteproperty VIN $WebJSON.Vin
                    $obj | add-member noteproperty isTempVin $WebJSON.isTempVin
                    $obj | add-member noteproperty TempBuildWeek $BuildWeek
                    $obj | add-member noteproperty isPresold $WebJSON.isPresold
                    $obj | add-member noteproperty holdStatus $WebJSON.holdStatus
                    $obj | add-member noteproperty dealerCategory $DCategory
                    $obj | add-member noteproperty dealerCd $WebJSON.dealerCd
                    $obj | add-member noteproperty DealerName $DealerName
                    $obj | add-member noteproperty DealerCity $DealerCity
                    $obj | add-member noteproperty DealerState $DealerState
                    $obj | add-member noteproperty DealerZip $DealerZip
                    $obj | add-member noteproperty distributorCd $WebJSON.distributorCd
                    $obj | add-member noteproperty year $WebJSON.year
                    $obj | add-member noteproperty bodyStyleDesc $WebJSON.bodyStyleDesc
                    $obj | add-member noteproperty extColor ($WebJSON.extColor.marketingName -replace ' \[extra_cost_color\]', '')
                    $obj | add-member noteproperty intColor $WebJSON.intcolor.marketingName
                    $obj | add-member noteproperty PriceInvoice $WebJSON.price.totalDealerInvoice
                    $obj | add-member noteproperty PriceTotalMSRP $WebJSON.price.totalMsrp
                    $obj | add-member noteproperty PriceAdvertised $WebJSON.price.advertizedPrice
                    $obj | add-member noteproperty HasMyOptions $OptionsFlag
                    $obj | add-member noteproperty Options ($WebJSON.options.marketingName -join ', ')
                    $obj | add-member noteproperty CarLink $CarLink 
                    $obj | add-member noteproperty DealerLink $DealerLink
                    $obj
                    }

                }
            }
        }
    $VINS | Export-CSV -NoTypeInformation "$Scriptdir\Output\$($VinModel)_$($VinPrefix[3])_Temp.csv"
	$NewVinDB | Export-CSV -NoTypeInformation "$Scriptdir\Database\$($VinModel)_$($VinPrefix[3])_Temp_VinDB.csv"
    }

# Otherwise it will just dump the VINs to a text file
else{
    $VINS = foreach ($PreSerial in $VinPreSerial){
        foreach ($MidSerial in $VinMidSerial){
            for ($i=0;$i-le 999;$i++){
                $vin = $VinPrefix + $PreSerial + $MidSerial + $(([string]$i).PadLeft(3,'0'))
                $vin -replace '-', (GetChecksumChar $vin)
                }
            }
        }

    Set-Content -Path "$Scriptdir\Output\$($VinModel)_$($VinPrefix[3])_Temp.txt" -Encoding Unicode -Value $VINS
    }

# Export any API errors to CSV appending existing file
if ($APIErrors){
    $APIErrors | Export-Csv -NoTypeInformation -Append "$Scriptdir\Logs\ScriptErrors.csv"
    }

# Get Script End and export script runtime length appending existing file
$ScriptEnd = Get-Date
$ScriptRuntime = New-TimeSpan –Start $ScriptStart –End $ScriptEnd
$ScriptRuntimeMinutes = [math]::Round($ScriptRuntime.TotalMinutes,2)
"$VinModel $VinPrefix Temp script took $ScriptRuntimeMinutes Minutes to Run" | Out-File "$Scriptdir\Logs\ScriptRuntime.txt" -Append