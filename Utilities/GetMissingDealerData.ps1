# First, run the Temp/Perm scripts.  Then sort by the dealerCd column and copy the dealer code from any dealers missing details
# Paste the missing dealerCd's into a new Excel column then use Data > Remove duplicates so the list is unique and add them to the GetDealersData.txt file
# Running this script will get their details from the website and produce a CSV.  You can update the Dealers.csv with this new data to make it more complete

# Set the script directory whether it's ran from ISE or not
if ($PSScriptRoot){$scriptdir = $PSScriptRoot}
else{$scriptdir = Split-Path -Path $psISE.CurrentFile.FullPath}

# Import list of Dealers to get details on
$MissingDealers = gc "$Scriptdir\GetMissingDealerData.txt"

# Go through each item in the list
$DealerInfo = foreach ($Dealer in $MissingDealers){
    # Pull the dealer data from Toyota API
    try{$WebData = Invoke-WebRequest -Uri "https://www.toyota.com/service/tcom/dealerRefresh/dealerCode/$Dealer"}
    catch{if ($WebData) {clear-variable webdata}}
    
    # If page exists convert to JSON
    if ($WebData.StatusCode -eq 200){
        $WebJSON = $Webdata.Content | ConvertFrom-Json

        # Get the values wanted from the JSON
        $Name = $webjson.showDealerLocatorDataArea.dealerLocator.dealerLocatorDetail.dealerParty.specifiedOrganization.companyName.value
        $Address = $webjson.showDealerLocatorDataArea.dealerLocator.dealerLocatorDetail.dealerParty.specifiedOrganization.postalAddress.lineOne.value
        $City = $webjson.showDealerLocatorDataArea.dealerLocator.dealerLocatorDetail.dealerParty.specifiedOrganization.postalAddress.cityName.value
        $State = $webjson.showDealerLocatorDataArea.dealerLocator.dealerLocatorDetail.dealerParty.specifiedOrganization.postalAddress.stateOrProvinceCountrySubDivisionID.value
        $Zip = $webjson.showDealerLocatorDataArea.dealerLocator.dealerLocatorDetail.dealerParty.specifiedOrganization.postalAddress.postcode.value
        }

    # ID not found in Toyota API, clear variables so the report shows these blank
    else{
        Clear-Variable Name, Address, City, State, Zip
        }

    # Add the dealer data to the DealerInfo Array
    $obj = new-object psobject
    $obj | add-member noteproperty ID $Dealer
    $obj | add-member noteproperty Name $Name
    $obj | add-member noteproperty Address $Address
    $obj | add-member noteproperty City $City
    $obj | add-member noteproperty State $State
    $obj | add-member noteproperty Zip $Zip
    $obj | add-member noteproperty Link "https://www.toyota.com/dealers/dealer/$Dealer"
    $obj
    }

$DealerInfo | Export-Csv -NoTypeInformation "$Scriptdir\DealerResults.csv"