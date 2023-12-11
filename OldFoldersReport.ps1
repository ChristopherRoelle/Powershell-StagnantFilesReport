<#Written By: Christopher Roelle#>
<#Written On: 12/11/2020#>
<#Purpose Of: Finds folders that havent been modified in 'x' amount of days and outputs in a color-coded HTML file (can be exported in parts if excedes a certain amount)
                Red Table Rows are outside of threshold and should be addressed.
<#Last Updated: 03/22/2021#>
<#
===============|CHANGELOG|================
12/15/21 - First functional version of script. Outputs each file with date last modified
           metadata, as well as relative and absolute path.
12/18/21 - Added color-coded HTML Report Functionality for easier reading.
           Red denotes file date modified metadata older than $dayThreshold.
01/05/21 - Added better support for outputting parts based on number of lines
           outputted to the HTML report. Adjust $outputsUntilNextPart to break
           file into smaller parts. Chrome seems to support 15,000 lines before
           output gets cut off or crashes.
03/22/21 - Added auto-path generation for report if path doesnt exist.
           If $outputFilePath is pointing to a path that doesnt exist, then
           we force the generation of folders to create a valid directory.
           Added dynamic directory naming, using the date the report is ran,
           this will auto clean-up the ran reports and prevent overwriting similar
           named reports if the old report is left in the directory.
           Added a $filepathAppend variable. Use this if running reports with different
           parameters, default is "". This is appended after folder timestamp.
#>

$debug = 0 <#Outputs the echo statements#>
$runDate = Get-Date

$exportInParts = 1 <#Useful if HTML document is too large to fully load in browser#>
$ftpRootDirectory = "" <#The path the script will check inside recursively#>

$outputFilePath = "" + "\\" <#@[User::JobProcessingDirectory]#><#This is just the path.#>
$outputFilename = "OldFolderReport-" + $runDate.month +"-"+ $runDate.day +"-"+ $runDate.year <#This is just the name, exclude extension#>
$dayThreshold = 1095 <#How many days should this report exclude prior to deeming a file old.#>
$outputsUntilNextPart = 15000 <#How many outputs to the file before starting another#>
<#End Configuration#>

$outputFileExt = "html"
$partCount = 1

<#FUNCTIONS#>
function printHTMLHeader
{
    "`<!DOCTYPE html`>" | Out-file -filePath $output -append
    "`<html`>" | Out-file -filePath $output -append
    "`<head`>" | Out-file -filePath $output -append

    if($exportInParts -eq $true)
    {
        "`<title`>$dayThreshold-Day Stagnant FTP Report Part $partCount`</title`>" | Out-file -filePath $output -append
    }
    else
    {
        "`<title`>$dayThreshold-Day Stagnant FTP Report`</title`>" | Out-file -filePath $output -append
    }
    
    if($exportInParts -eq 1){ $exportInPartsParameterText = "True - After $outputsUntilNextPart files logged"} else { $exportInPartsParameterText = "False"}
    if($debug -eq 1){ $debugParameterText = "True"} else { $debugParameterText = "False"}

    "`<style`>" | Out-file -filePath $output -append
    "#t01 { width: 95%; margin: 0 auto; }" | Out-file -filePath $output -append
    "#t01 tr:nth-child(even) { background-color: #eee; }" | Out-file -filePath $output -append
    "#t01 tr:nth-child(odd) { background-color: #fff; }" | Out-file -filePath $output -append
    "#t01 th { color: white; background-color: black; }" | Out-file -filePath $output -append
    ".heading { font-weight: bold; }" | Out-file -filePath $output -append
    "#t01 tr.old { background-color: #ff6666; }" | Out-file -filePath $output -append
    "`</style`>" | Out-file -filePath $output -append

    "`</head`>" | Out-file -filePath $output -append
    "`<body`>" | Out-file -filePath $output -append

    "`<h5`>Report Date: $runDate`</h5`>" | Out-file -filePath $output -append
    "`<p`>The following FTP files are beyond their $dayThreshold-day threshold and need action taken.`</p`>" | Out-file -filePath $output -append
    "`<p`>`<small style='color=gray;'`>Report ran with the following parameters: Threshold: $dayThreshold days, Output in parts: $exportInPartsParameterText, Debug Mode: $debugParameterText`</small`>`</p`>" | Out-file -filePath $output -append
    "`<table style='width=100%;' id='t01'`>" | Out-file -filePath $output -append

    "`<tr`>`<td class='heading'`>File #`</td`>`<td class='heading'`>Path`</td`>`<td class='heading'`>Name`</td`>`<td class='heading'`>Last Modified`</td`>`</tr`>" | Out-file -filePath $output -append
}

function printHTMLFooter
{
    "`</table`>" | Out-file -filePath $output -append
    "`</body`>" | Out-file -filePath $output -append
    "`</html`>" | Out-file -filePath $output -append
}

function printHTMLFooterWithTotal
{
    "`</table`>" | Out-file -filePath $output -append
    $endDate = Get-Date
    $elapsedTime = NEW-TIMESPAN -Start $runDate -End $endDate
    "`<h5`>Scanned $count folder(s).`</h5`>" | Out-file -filePath $output -append
    "`<h5`>There were $OldCount folder(s) that are older than $dayThreshold day(s).`</h5`>" | Out-file -filePath $output -append
    "`<h5`>The report took " + $elapsedTime.hours + ":" + $elapsedTime.minutes + ":" + $elapsedTime.seconds + " to run. (HH:MM:SS)`</h5`>" | Out-file -filePath $output -append
    "`</body`>" | Out-file -filePath $output -append
    "`</html`>" | Out-file -filePath $output -append
    if($debug -eq 1){ Write-output ("The report took " + $elapsedTime.hours + ":" + $elapsedTime.minutes + ":" + $elapsedTime.seconds + " to run. (HH:MM:SS)") }
}

function fileNameBuilder
{
    if($exportInParts -eq $true)
    {
        $filename = $outputFilePath + $outputFilename + "-Part" + $partCount + "." + $outputFileExt
        return $filename
    }
    else
    {
        $filename = $outputFilePath + $outputFilename + "." + $outputFileExt
        return $filename
    }
}

<#END FUNCTIONS#>

<#BEGIN PROCESS#>
If(!(test-path $outputFilePath)){New-Item -ItemType Directory -Force -Path $outputFilePath}<# Generates path if doesnt exist #>

$output = fileNameBuilder
printHTMLHeader

$FolderNamesToExclude = @("*\#OLD\*") <# !!! Probably broken, notlike doesnt seem to iterate through arrays, NEEDS FIX !!! #>
$count = 1
$OldCount = 0


Get-ChildItem -Path $ftpRootDirectory -Recurse | Where {$_.FullName -notlike $FolderNamesToExclude} | where-object { $_.PSIsContainer -eq $true } | 
Foreach-Object{
    if($exportInParts -eq 1 -and $count -gt 1 -and ($count % $outputsUntilNextPart) -eq 0)
    { 

        "`<tr`>`<td`>END OF PART`</td`>`</tr`>" | Out-file -filePath $output -append
        printHTMLFooter
        $partCount = $partCount + 1
        $output = fileNameBuilder
        printHTMLHeader
        
    }
    if($_.LastWriteTime -lt (Get-Date).AddDays(-$dayThreshold)) <#If last Write Date is < Day + Threshold (Days back)#>
    {
        if($debug -eq 1){ Write-Output ($_.FullName + " This file is old. Added to Log - Part $partCount") }

        "`<tr class='old'`>`<td`>" + $count + "</td`>`<td`>" + $_.FullName + "</td`>`<td`>" + $_.Name + "</td`>`<td`>" + $_.LastWriteTime + "`</td`>`</tr`>" | Out-file -filePath $output -append
        $count = $count + 1
        $OldCount += 1
        }
    else
    {
        if($debug -eq 1){ Write-Output ($_.FullName + " This file is NOT old. Ignored.") }
        "`<tr class='notOld'`>`<td`>" + $count + "</td`>`<td`>" + $_.FullName + "</td`>`<td`>" + $_.Name + "</td`>`<td`>" + $_.LastWriteTime + "`</td`>`</tr`>" | Out-file -filePath $output -append
        $count = $count + 1
    }
}

printHTMLFooterWithTotal

Write-Output("Runtime Complete.")
