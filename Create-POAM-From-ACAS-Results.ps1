## Author: JLD

## PURPOSE: Create POAM from ACAS results

## REQUIREMENTS: 
### - Acas results file to import into this script
### - A POAM templated excel file for use after script execution. You will be copy/pasting the output from this script into your POAM excel file.

## ACAS Results must be converted to a MS-DOS (.csv) file in order for Powershell to import them properly and for this script to work. To make this change on a .xls(x) file, select 'Save As' and then change the format type.

## HOW TO USE:

## Step 1: Convert ACAS results to a MS-DOS (.csv) file in order for Powershell to import them properly. To make this change on a .xls(x) file, select 'Save As' and then change the format type. <<< If you do not do this properly, Powershell will not import properly.
## Step 2: The above step will get rid of any additional tabs in the results file. If you need the other tabs, copy/paste them into separate MS-DOS .csv files and run this script against each of them one-by-one.
## Step 3: Execute script. Answer prompt with filename. Press Enter when ready to continue.
## Step 4: Results will be output to a file in your C:\Temp\ directory.
## Step 5: Open up this Results output file from this script in in excel
## Step 6: Copy ALL the data (besides the column headers) to your clipboard.
## Step 7: Open up the blank POAM template that you have
## Step 8: Paste all of your copied clipboard data to this POAM spreadsheet
## Step 9: Re-format as needed (Wrap Text, Add any missing border lines) in your blank POAM template.
## Step 10: All set.




## MAKE CHANGES HERE

$ACASResults = Import-csv -Path "C:\Temp\ACAS-scan-results.csv" # Change this to the filepath of your ACAS results file. MAKE SURE this is a .csv (MS-DOS) type of file. Refer to Step 1 above

## END CHANGES SECTION





## BEGIN SCRIPT ######################################################################################################################################################################################################################




$Filename = Read-Host -Prompt "What would you like the name of the output file to be? Please do NOT include file extension in your answer"

if ($Filename -like "*.csv*" -or $Filename -like "*.txt*"){

write-host -ForegroundColor Red "Please do not include the file extension in this prompt. Re-run the script again and provide ONLY a filename."
break

}


## Output directory path. Must end with "\"
$OutFilePath = "C:\Temp\"

## Formatting the output file
$FinalDestination = $OutFilePath + $Filename + ".csv"

## Declaring the object that will be built
$POAMObj = @()

## Estimated completion date. Change the "15" below to modify estimated completion date.
[string]$ECDDate = (get-date).AddDays(15).ToSTring('MM-dd-yyyy')


## Iterate through the imported csv ACAS results
foreach ($Diffoption in $ACASResults){

if ($Diffoption.Severity -match "Critical"){

$RawSevere = "I"
$ImpactSevere = "Very High"

}

if ($Diffoption.Severity -match "High"){

$RawSevere = "I"
$ImpactSevere = "High"

}

if ($Diffoption.Severity -match "Medium"){

$RawSevere = "II"
$ImpactSevere = "Moderate"

}

if ($Diffoption.Severity -match "Low"){

$RawSevere = "III"
$ImpactSevere = "Low"

}

## Building the object that will be output to csv. Make any changes necessary here
$POAMObj += New-Object PSObject -Property @{

    CVD = $Diffoption.Synopsis
    SCN = ""
    Office = "66 ABG/SCOO"
    Security = ""
    Resources = ""
    Scheduled = $ECDDate
    Milestone = "Researching/Testing before completion. Estimated completion date: $ECDDate"
    MilestoneTwo = ""
    Source = "ACAS Scan Results"
    Status = "On-going"
    Comment = "Researching impact from remediating vulnerability"
    RawSeverity = $RawSevere
    Mitigation = ""
    Severity = $ImpactSevere
    Relevance = $ImpactSevere
    Likelihood = $ImpactSevere
    Impact = $ImpactSevere
    ImpactDesc = $Diffoption.Description
    ResidRisk = $ImpactSevere

} ## end of Object Properties

} ## end of Foreach loop


## Output the POAM object to a csv file 
$PoamObj | Select-Object -Property CVD, SCN, Office, Security, Resources, Scheduled, Milestone, MilestoneTwo, Source, status, Comment, RawSeverity, Mitigation, Severity, Relevance, Likelihood, Impact, ImpactDesc, ResidRisk | sort-object -Property CVD, SCN, Office, Security, Resources, Scheduled, Milestone, MilestoneTwo, Source, status, Comment, RawSeverity, Mitigation, Severity, Relevance, Likelihood, Impact, ImpactDesc, ResidRisk | Export-csv -Path $FinalDestination -Append -NoTypeInformation


if (Test-Path $FinalDestination){
write-host -ForegroundColor Green "New output file has been created and stored here: $FinalDestination"
}
else {
write-host -ForegroundColor Red "Unable to create output file here: $FinalDestination      Please double-check this is a valid directory/filepath and re-run script."
}