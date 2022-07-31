
==========================================
=           General Information          =
==========================================

There are two scripts that need to be ran for each Trim code.  You modify the VinPrefix variable to determine which trim code will be used.
	TempVins.ps1	This will try generating temp vins which start with week build date and a letter (26A for example)
	PermVins.ps1	This will try generating non-temp vins which start with 3 numbers (041 for example)

There are 7 trim codes in total so 14 scripts you can run at one time for "multi-threading".  This is how I have the script set up if you copied all data from my OneDrive.  It will hit the Toyota API URL 14 VINs at a time if all of the scripts are running concurrently.  I tested this and didn't see any errors generated that weren't 404 so I imagine this does not trigger any kind of DDOS protection from Toyota but your results may vary.  Pay attention to the Logs folder and if you see a ScriptErrors.csv in there, open it up and figure out what happened.  If it wasn't a 404 error when you hit the API it should have the exception message here.  

	LE	Has 2 trim codes (C and H)
	SE	Has 2 trim codes (G and S)
	SE NS	1 trim code (T) NS stands for Nightshade
	XSE	1 trim code (K)
	XLE	1 trim code (F)

Logs will go into the Logs folder and the CSV Output from the scripts will go in the Output folder. 
	Logs are appended so you might want to manually delete them after some time though I doubt they would grow very large. 
	Output files are overwritten 

**NEW** Database Folder
	This will store VINs that were discovered and when they were first seen to keep up with how old the vehicle might be

**NEW** Utilities!! Read about them at the bottom of the script

!! HOW YOU CAN HELP !!
My VIN logic isn't that great so this will not get every possible VIN in Toyota's allocation system.  If you find VINs missing, please send them my way and I will do my best to decipher them and update the script to generate and scan Toyota's API for them.  

!!! IMPORTANT !!!
You must have the Dealers.csv file in the script root directory for it to be imported and used in the outputted CSV.  You can update this CSV using the utility described below.  It's very handy to sort by state or zip code so it's important to update this CSV regularly.


=================================
= Executing all scripts at once =
=================================

_Run All PowerShell Scripts.bat

I included a handy bat file that will launch all PowerShell scripts at one time. Simply edit this bat file to contain the names of your PowerShell scripts and then double click it. 

As long as the PowerShell scripts are in the same folder as the BAT file it will execute all of them at once with the flag that lets you bypass requiring them to be signed.  

A word of warning, make sure you read the scripts before executing them in case someone went in and made changes that will damage your computer.

		

===============================================
= Notes about the fields in the results table =
===============================================

Below is what I found searching around the internet for other Allocation Sheets.  Forgive me if anything is incorrect.

As far as I can tell, you want to filter the spreadsheet like the below.  This should show you vehicles that are not held/resold and are either at the dealership or on its way to them.  You can then call these dealers on the VINs and validate if they're available and if you can put a deposit on one to purchase it.
	holdStatus		Available
	isPreSold		False		
	dealerCategory	Ground,Freight

holdStatus:
If this is marked "DealerHold", this means the Dealer has marked that they want to receive the vehicle (they aren't trying to trade it, etc, etc). If you have placed a deposit or have reserved this vehicle, seeing "DealerHold" should be somewhat comforting. (for those who don't trust verbals)

isPreSold
If you have placed a deposit on a vehicle or have reserved it, I would not worry about this flag. When I placed my deposit this was never 'true' for me, and I have not seen it marked 'true' for any other VI I've looked up.​

dealerCategory:
A = allocated F = freight G = ground
A: allocated
F: Your truck is on a boat, or sitting in a port, or possibly on a train.
G: Your truck is on a truck being delivered to the dealer, or is at the dealer. Not the same as transported (train, boat, tooth fairy, etc).​

isTempVin
First 2 characters in the 6 character serial of a temp VIN is the week number of the year the vehicle will be built.  For 2022, 25-June 25, 26-July 2, 27-July 9, 28-July 16 and so forth.

Temps VINs are assigned once the production sequence of the vehicle has actually been assigned. You can enter it into the Toyota owner's site and it will show the basic configuration. That number disappears from the system once the vehicle hits the production line and an actual VIN has been assigned. The dealer sees the vehicle on the Dealer Daily system within about 24 hours of production being completed and has access to the VIN that's been assigned at that point. If they can provide an actual VIN, the vehicle has already been built and is somewhere in the transport sequence.


==========================================
= Script Description in Utilities Folder =
==========================================

GetMissingDealerData.ps1
========================
My original dealers.csv was copied from https://github.com/rgoldfinger/rav4_scraping/blob/main/util/dealers.ts

Someone generated numbers from 00001 - 99999 and hit the Toyota API at https://www.toyota.com/service/tcom/dealerRefresh/dealerCode/#####

I'm guessing that was done a long time ago as there are a lot of missing dealers from the dealers.ts source file.  

If you run my scripts, you can copy the DealerCD field (dealer code) and run it through this script.  It will connect to the Toyota API and give you a CSV that contains any found dealer codes.  Then you can update dealers.csv to include the new information so that the next time you run the script it will have all of the dealer data and not be missing anything.


ConvertTempVINtoBuildWeek.ps1
=============================
This lets you input a VIN and get the build week according to logic I found about the temp VIN.

It's not very useful since I added this functionality to the main scripts but I didn't want to delete my initial script for it in case it was useful in the future.


CombineCSVsIntoOneExcel.ps1
===========================
If you drop all of the generated CSVs into the scriptroot directory of this script and run this, it will combine them all into an XLSX file.  

You must have Excel installed on the system where this runs as it launches an Excel process to do the work in.

If you get an error, chances are it did not remove the CSV queries from the XLSX file.  You can manually do this by opening the generated Excel file and go to Dava > Queries & Connections and right click on the queries and delete them.  This isn't really necessary but it bugs me when Excel says "Security Warning External Data Connections have been disabled!!" in the top bar and removing the queries fixes that becuase it will no longer be looking for an external CSV to update from.

***ADDED*** I streamlined this to suck in all of the CSVs and then output based off the bodyStyleDesc column.  It also deletes the old CSVs to keep things clean.


*** NEW *** 

MissingDealerData_Combine.ps1
=============================
This lets you take the new Dealers CSV and the original Dealers CSV and merge/dedup them easily


CombineCSVsIntoOneExcel_OneSheet.ps1
====================================
Sucks in all of the CSVs into a single Worksheet rather than creating a worksheet for each trim level.  

Personally I don't want them split out b/c I'm looking for the Driver Assist so I want to see both XLE and XSE but you do you.


Copy_Scripts_And_Rename.ps1
===========================
So this one is really handy.  Instead of having to modify 14 different scripts, you modify two.  

I put the source scripts in the Utilities\SourceScripts folder.  I make my edits to these two scripts to improve them. 

Then I run the Copy Scripts and Rename script and it will suck in the two source scripts and generate 14 scripts, 1 for each trim.  Very useful!


Compare_Source_Scripts.ps1
==========================
This script will suck in the two source scripts mentioned above and tell you what's different about them.  

I'm making the same changes to both scripts and might make a mistake with a copy paste so this helps prevent that. 