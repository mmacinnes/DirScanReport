#****************************************************************************************
# Name: dirScan.ps1
# Description: This script scans a directory and its sub directories to determine if the
#              directory is ready to be archived.
#              The script will output the results to the console and to a report file.
#              The script will prompt the user for the directory to scan and the cut off 
#              date for archiving.
#              
#              The script output will be in the current directory as: 
#                       dirScanReport_yyyy-MM-dd_HH-mm-ss.txt
#
# usefule reference 
#  https://igorpuhalo.wordpress.com/2019/08/29/overcoming-long-path-problem-in-powershell/
#  https://learn.microsoft.com/en-us/windows/win32/fileio/naming-a-file
#
#****************************************************************************************
class dateInfo {
    [int]$fileCount
    [int]$subDirCount
    [datetime]$minCreationDate
    [datetime]$maxCreationDate
    [datetime]$minModifiedDate
    [datetime]$maxModifiedDate
}

function Get-DirectoryInfo {
    param (
        $directory, $dateInfo 
    )
  
    # Get all files in the directory
    try {
        $files = Get-ChildItem -LiteralPath ('\\?\'+$directory) -File -ErrorAction Stop
    }
    catch {
        <#Do this if a terminating exception happens#>

        Write-Host "Error scanning files: $directory : $_"
        Write-Host " "
        #Write-Output "Error scanning files: $directory : $_"
        return
    }
   
    # Count the number of files
    $dateInfo.fileCount = $dateInfo.fileCount + $files.Count

    if ($files.Count -eq 0) {}
    else {
        # Get the minimum and maximum creation dates
        $minCreationDate = $files | Measure-Object -Property CreationTime -Minimum | Select-Object -ExpandProperty Minimum
        $maxCreationDate = $files | Measure-Object -Property CreationTime -Maximum | Select-Object -ExpandProperty Maximum

        $minModifiedDate = $files | Measure-Object -Property LastWriteTime -Minimum | Select-Object -ExpandProperty Minimum
        $maxModifiedDate = $files | Measure-Object -Property LastWriteTime -Maximum | Select-Object -ExpandProperty Maximum
        
        if ($maxCreationDate -gt (Get-Date)) {
            $maxCreationDate = (Get-Date)
        }
        if ($maxModifiedDate -gt (Get-Date)) {
            $maxModifiedDate = (Get-Date)
        }

        if ($dateInfo.minCreationDate -gt $minCreationDate) {
            $dateInfo.minCreationDate = $minCreationDate
        }
        if ($dateInfo.maxCreationDate -lt $maxCreationDate) {
            $dateInfo.maxCreationDate = $maxCreationDate
        }
        if ($dateInfo.minModifiedDate -gt $minModifiedDate) {
            $dateInfo.minModifiedDate = $minModifiedDate
        }
        if ($dateInfo.maxModifiedDate -lt $maxModifiedDate) {
            $dateInfo.maxModifiedDate = $maxModifiedDate
        }
    }
    Get-SubDirectoryInfo -directory $directory -dateInfo $dateInfo
}

function Get-SubDirectoryInfo {
    param (
        $directory, $dateInfo
    )

    try {
        $Directories = Get-ChildItem -LiteralPath $directory -Directory -ErrorAction Stop
    }
    catch {
        <#Do this if a terminating exception happens#>
        Write-Host "Error scanning sub directory: $directory : $_"
        Write-Host " "
        #Write-Output "Error scanning sub directory: $directory : $_"
        return
    }

    # Count the number of sub directories
    $dateInfo.subDirCount = $dateInfo.subDirCount + $Directories.Count

    foreach ($dir in $Directories) {
        Get-DirectoryInfo -directory $dir.FullName -dateInfo $dateInfo
    }
}

function isReadyToArchive {
    param (
        $archiveDate, $dateInfo
    )

    $readyToArchive = $false
    if ($dateInfo.fileCount -eq 0) {
        $readyToArchive = $true
    }
    elseif ($dateInfo.maxModifiedDate -lt $cutOffDate ) {
        $readyToArchive = $true
    }
    elseif ($dateInfo.maxCreationDate -lt $cutOffDate ) {
        $readyToArchive = $true
    }
    return $readyToArchive
}

function writeInfo {
    param (
        $message, $saveToFile, $filePath
    )
    Write-Host $message

    if ($saveToFile) {
        Try {
            #Add Content to the report File
            Add-content -Path  $filePath -Value $message
        }
        Catch {
            Write-host -f Red "Error:" $_.Exception.Message
        }
    }
}

function outputInfo {
    param (
        $dir, $cutOffDate, $dateInfo, $filePath
    )

    $readyToArchive = (isReadyToArchive -archiveDate $cutOffDate -dateInfo $dateInfo)

    $message = ""
    $saveToFile = $false
    if ($readyToArchive) {
        $saveToFile = $true
    }

    # Output the results to the console
    $message = "Directory:" + $dir.FullName
    
    writeInfo $message $saveToFile $filePath

    if ($readyToArchive) {
        $message = "Ready to archive"
    } 
    else {
        $message = "Not ready to archive"
    }
    writeInfo $message $saveToFile $filePath
    
    $message = "Number of sub directories:" + $dateInfo.subDirCount.ToString()
    writeInfo $message $saveToFile $filePath

    $message = "Number of files:" + $dateInfo.fileCount
    writeInfo $message $saveToFile $filePath

    if ($dateInfo.fileCount -eq 0) {
        $message = "No files found in " + $dir.FullName
        writeInfo $message $saveToFile $filePath
    }
    if ($dateInfo.maxCreationDate -lt $cutOffDate ) {
        $message = "No files created after " + $dateInfo.maxCreationDate.Date.ToString("yyyy/MM/dd")
        writeInfo $message $saveToFile $filePath
    }
    if ($dateInfo.maxModifiedDate -lt $cutOffDate ) {
        $message = "No files modified after " + $dateInfo.maxModifiedDate.Date.ToString("yyyy/MM/dd")
        writeInfo $message $saveToFile $filePath
    }
    else {
        $message = "creation Min:" + $dateInfo.minCreationDate.Date.ToString("yyyy/MM/dd") + "  Max:" + $dateInfo.maxCreationDate.Date.ToString("yyyy/MM/dd")
        writeInfo $message $saveToFile $filePath
        $message = "modified Min:" + $dateInfo.minModifiedDate.Date.ToString("yyyy/MM/dd") + "  Max:" + $dateInfo.maxModifiedDate.Date.ToString("yyyy/MM/dd") 
        writeInfo $message $saveToFile $filePath
    }
    $message = " "
    writeInfo $message $saveToFile $filePath
}

#****************************************************************************************
# Main
#****************************************************************************************

# Set report file path
$currDate = (Get-Date).tostring("yyyy-MM-dd_HH-mm-ss")

# Set report file path
#$filePath = "C:\\temp\\dirScanReport_" + $currDate + ".txt"
$filePath = (Get-Location).Path +"\\dirScanReport_"+ $currDate + ".txt"

# Set folder to scan
$directory = "C:\\BuildingKnowledge"
$dateInfo = [dateInfo]::new() 

$dateInfo.fileCount = 0
$dateInfo.subDirCount = 0
$dateInfo.minCreationDate = [datetime]::MaxValue 
$dateInfo.maxCreationDate = [datetime]::MinValue 
$dateInfo.minModifiedDate = [datetime]::MaxValue 
$dateInfo.maxModifiedDate = [datetime]::MinValue

$userDirectory = Read-Host "Enter the directory to scan:"
if ($userDirectory -ne "") {
    $directory = $userDirectory
}
else {
    Write-Host "Using default directory: $directory"
    $directory = Get-Location
}
#Set the cut off date for archiving
$cutOffDate = Read-Host "Enter the cut off date (yyyy-mm-dd):"
if ($cutOffDate -ne "") {
    $cutOffDate = Get-Date -Date $cutOffDate
}
else {
    $currYear = ((Get-Date).AddYears(-4)).Year
    $cutOffDate = (Get-Date -Year $currYear -Month 01 -Day 01)
    Write-Host "Using default date: " $cutOffDate.ToString("yyyy-MM-dd")
}

$message = " "
writeInfo $message $true $filePath
$message = "Directory to scan: " + $directory
writeInfo $message $true $filePath
$message = "Cut off date: " + $cutOffDate.ToString("yyyy-MM-dd")
writeInfo $message $true $filePath
$message = " "
writeInfo $message $true $filePath

try {
    $Directories = Get-ChildItem -Path $directory -Directory -ErrorAction Stop
}
catch {
    <#Do this if a terminating exception happens#>
    $message = "Error scanning directories: $directory : $_"
    writeInfo $message $true $filePath
    $message = " "
    writeInfo $message $true $filePath
    return
}

$dirToArchiveCnt = 0

foreach ($dir in $Directories) {

    $dateInfo.subDirCount = 0
    $dateInfo.fileCount = 0
    $dateInfo.minCreationDate = [datetime]::MaxValue 
    $dateInfo.maxCreationDate = [datetime]::MinValue 
    $dateInfo.minModifiedDate = [datetime]::MaxValue 
    $dateInfo.maxModifiedDate = [datetime]::MinValue
    
    Get-DirectoryInfo -directory $dir.FullName -dateInfo $dateInfo
    
    if ((isReadyToArchive -archiveDate $cutOffDate -dateInfo $dateInfo)) {
        $dirToArchiveCnt++
    }

    outputInfo -dir $dir -cutOffDate $cutOffDate -dateInfo $dateInfo -filePath $filePath 
}
$message = "Number of directories to Archive: " + $dirToArchiveCnt.ToString()
writeInfo $message $true $filePath
$message = " "
writeInfo $message $true $filePath

return 


