# Set the script directory whether it's ran from ISE or not
if ($PSScriptRoot){$scriptdir = $PSScriptRoot}
else{$scriptdir = Split-Path -Path $psISE.CurrentFile.FullPath}

# Get all CSVs in script rood directory
$Original = @(gci $scriptdir "*.csv")

# Warn if no CSVs were found
if ($Original.count -eq 0){
    Write-Host "There were no CSVs found... Press Enter to Exit the Script."-ForegroundColor Red
    Read-Host
    Exit
    }

# Import all CSVs to variable
$CSVData = @()
foreach ($CSV in $Original){
    $CSVData += Import-Csv $CSV.FullName
    }

# Dump out the merged CSVs based on model (bodyStyleDesc)
$CSVData | ? {$_.bodyStyleDesc -like 'LE HYBRID SEDAN'} | Export-Csv -NoTypeInformation "$Scriptdir\LEHybridSedan.csv"
$CSVData | ? {$_.bodyStyleDesc -like 'SE HYBRID NIGHTSHADE'} | Export-Csv -NoTypeInformation "$Scriptdir\SEHybridNightshade.csv"
$CSVData | ? {$_.bodyStyleDesc -like 'SE HYBRID SEDAN'} | Export-Csv -NoTypeInformation "$Scriptdir\SEHybridSedan.csv"
$CSVData | ? {$_.bodyStyleDesc -like 'XLE HYBRID SEDAN'} | Export-Csv -NoTypeInformation "$Scriptdir\XLEHybridSedan.csv"
$CSVData | ? {$_.bodyStyleDesc -like 'XSE HYBRID SEDAN'} | Export-Csv -NoTypeInformation "$Scriptdir\XSEHybridSedan.csv"

# Get path to all merged CSVs
$CSVs = @(gci $scriptdir "*hybrid*.csv")

# Set output file name
$outputfilename = "CamryHybridAllocationSpreadsheet.xlsx"

# Load the Excel Application but leave it invisible to the end user
$excelapp = new-object -comobject Excel.Application
$excelapp.DisplayAlerts = $False
$excelapp.Visible = $False

# Set the total number of workbooks (1 for each CSV)
$excelapp.sheetsInNewWorkbook = $csvs.Count
$xlsx = $excelapp.Workbooks.Add()
$sheet=1

# Go through each CSV and add it to the Excel
foreach ($csv in $csvs){
    # Create a sheet and give it a name
    $worksheet = $xlsx.Worksheets.Item($sheet)
    $worksheet.Name = ($csv.Name -replace '.csv')

    # Connect to the CSV and add the query, import the data, then delete the query
    $TxtConnector = ("TEXT;" + $csv.FullName)
    $Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
    $query = $worksheet.QueryTables.item($Connector.name)
    $query.TextFileOtherDelimiter = $Excelapp.Application.International(5)
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1
    
    # The sleeps help prevent the error of removing the query
    Start-Sleep -Seconds 1
    $query.Refresh() | Out-Null
    Start-Sleep -Seconds 1
    $query.Delete() | Out-Null
    Start-Sleep -Seconds 1
    
    # Increment the sheet number
    $sheet++
    }

# Excel is built, save into the output location then close out of Excel
$output = $scriptdir + "\" + $outputfilename
$xlsx.SaveAs($output)
$xlsx.Close()
$excelapp.Quit()

# Remove All CSVs
$Original | Remove-Item -Force
$CSVs | Remove-Item -Force