# This is the script I use when there are missing dealers.  I copy the original Dealers.CSV and the new DealerResults.CSV 
# into the same directory then dedupe and merge them into a new one.  Then I rename that new one and replace the old 
# in the main directory.

# Set directory for CSVs
$CSVDir = "D:\Misc\Temp"

# Import the original CSV and new CSV
$Dealer = ipcsv "$CSVDir\Dealers.csv"
$New = ipcsv "$CSVDir\DealerResults.csv"

# Combine the two 
$Combine = @()
$Combine += $Dealer
$Combine += $New

# Compare difference between the two
$Combine.Count
($Combine | Sort ID -Unique).count

# Show duplicates
# ($Combine | Group ID) | ? {$_.count -gt 1}

# Clean up duplicates
$CombineUnique = $Combine | Sort ID -Unique

# Export new Dealers.csv
$CombineUnique | epcsv "$CSVDir\NewDealers.csv"