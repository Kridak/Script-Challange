# Set source directory
$sourceDir = "~/Desktop/Shell"

# Set target directory
$targetDir = "~/Desktop/SS"

# Set archive directory
$archiveDir = "~/Desktop/Arc"

# Set log directory
$logDir = "~/Desktop/Shell"

# Set file masks
$fileMask1 = "*.xlsx"
$fileMask2 = "*.xlsm"

# Move files from source directory to target directory
Get-ChildItem -Path $sourceDir -Filter $fileMask1 -Recurse | ForEach-Object {
    if (Test-Path -Path "$targetDir\$($_.Name)") {
        Write-Error "ERROR: Multiple files with the same file mask found!"
        exit 1
    }
    else {
        Move-Item -Path $_.FullName -Destination $targetDir
    }
}

Get-ChildItem -Path $sourceDir -Filter $fileMask2 -Recurse | ForEach-Object {
    if (Test-Path -Path "$targetDir\$($_.Name)") {
        Write-Error "ERROR: Multiple files with the same file mask found!"
        exit 1
    }
    else {
        Move-Item -Path $_.FullName -Destination $targetDir
    }
}

# Check if Excel files are empty
Get-ChildItem -Path $targetDir -Filter $fileMask1 -Recurse | ForEach-Object {
    if (-not (Get-Content $_.FullName)) {
        Write-Error "ERROR: Excel file $($_.Name) is empty!"
        exit 1
    }
}

Get-ChildItem -Path $targetDir -Filter $fileMask2 -Recurse | ForEach-Object {
    if (-not (Get-Content $_.FullName)) {
        Write-Error "ERROR: Excel file $($_.Name) is empty!"
        exit 1
    }
}

# Convert Excel files to UTF-8 encoded CSV files
Get-ChildItem -Path $targetDir -Filter $fileMask1 -Recurse | ForEach-Object {
    $csvFile = "$($_.DirectoryName)\$($_.BaseName).csv"
    Get-Content $_.FullName | Out-File -Encoding UTF8 -FilePath $csvFile
}

Get-ChildItem -Path $targetDir -Filter $fileMask2 -Recurse | ForEach-Object {
    $csvFile = "$($_.DirectoryName)\$($_.BaseName).csv"
    Get-Content $_.FullName | Out-File -Encoding UTF8 -FilePath $csvFile
}

# Archive original Excel files
Get-ChildItem -Path $targetDir -Filter $fileMask1 -Recurse | ForEach-Object {
    $currDate = Get-Date -Format "yyyyMMdd"
    $currTime = Get-Date -Format "HHmmss"
    $archiveFile = "$archiveDir\$($_.Name)_$currDate_$currTime.zip"
    Compress-Archive -Path $_.FullName -DestinationPath $archiveFile
}

Get-ChildItem -Path $targetDir -Filter $fileMask2 -Recurse | ForEach-Object {
    $currDate = Get-Date -Format "yyyyMMdd"
    $currTime = Get-Date -Format "HHmmss"
    $archiveFile = "$archiveDir\$($_.Name)_$currDate_$currTime.zip"
    Compress-Archive -Path $_.FullName -DestinationPath $archiveFile -Force
}

# Delete original Excel files
Get-ChildItem -Path $targetDir -Filter $fileMask1 -Recurse | ForEach-Object {
    Remove-Item -Path $_.FullName
}

Get-ChildItem -Path $targetDir -Filter $fileMask2 -Recurse | ForEach-Object {
    Remove-Item -Path $_.FullName
}

# Log operations
$currDate = Get-Date -Format "yyyyMMdd"
$currTime = Get-Date -Format "HHmmss"
$logFile = "$logDir\log_$currDate_$currTime.txt"
$logMessage = "[$currDate $currTime] Excel files moved from $sourceDir to $targetDir, converted to UTF-8 encoded CSV files, archived in $archiveDir, and deleted from $targetDir"
Add-Content -Path $logFile -Value $logMessage

Write-Host "Done!"