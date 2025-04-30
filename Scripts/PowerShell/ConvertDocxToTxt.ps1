# PowerShell Script: Convert .docx files to .txt format in specified directory and subdirectories
# Author: AI Assistant
# Date: Auto-generated

param (
    [string]$SourceDir = ".",
    [string[]]$SourceFiles,
    [switch]$DeleteOriginal = $false,
    [switch]$Backup = $true,
    [switch]$Silent = $false
)

# Set script execution path to current script directory
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Set-Location -Path $scriptPath

# Initialize counters
$totalFiles = 0
$convertedFiles = 0
$failedFiles = 0
$skippedFiles = 0

# If backup option is enabled, create a backup directory
$backupDir = $null
if ($Backup) {
    $backupDirName = "DocxBackup_" + (Get-Date -Format "yyyyMMdd_HHmmss")
    $backupDir = Join-Path -Path $scriptPath -ChildPath $backupDirName
    New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
    Write-Host "Backup directory created: $backupDir" -ForegroundColor Cyan
}

# Create a Word application COM object
try {
    $word = New-Object -ComObject Word.Application
    # Set Word to be invisible
    $word.Visible = $false
    # Disable alerts
    $word.DisplayAlerts = 0
}
catch {
    Write-Error "Cannot create Word application, please ensure Microsoft Word is installed: $_"
    exit 1
}

# Get files to process
$docxFiles = @()
if ($SourceFiles) {
    foreach ($file in $SourceFiles) {
        # Handle both relative and absolute paths
        $filePath = $file
        if (-not [System.IO.Path]::IsPathRooted($file)) {
            $filePath = Join-Path -Path $scriptPath -ChildPath $file
        }
        
        if (Test-Path -Path $filePath -PathType Leaf) {
            if ($filePath -like "*.docx") {
                $docxFiles += Get-Item -Path $filePath
            }
            else {
                Write-Warning "Skipped non-DOCX file: $file"
            }
        }
        else {
            Write-Warning "File not found: $file"
        }
    }
}
else {
    # Set source directory path (handle both relative and absolute paths)
    $docxDir = $SourceDir
    if (-not [System.IO.Path]::IsPathRooted($SourceDir)) {
        $docxDir = Join-Path -Path $scriptPath -ChildPath $SourceDir
    }

    # Check if source directory exists
    if (-not (Test-Path -Path $docxDir -PathType Container)) {
        Write-Error "Directory $SourceDir does not exist: $docxDir"
        exit 1
    }

    # Recursively get all .docx files
    $docxFiles = Get-ChildItem -Path $docxDir -Filter "*.docx" -Recurse -File
}

# Show the number of files found
$totalFiles = $docxFiles.Count
if ($SourceFiles) {
    Write-Host "Found $totalFiles .docx files from specified source files" -ForegroundColor Yellow
}
else {
    Write-Host "Found $totalFiles .docx files in $SourceDir directory and its subdirectories" -ForegroundColor Yellow
}

# Create a log file to record conversion results
$logFile = Join-Path -Path $scriptPath -ChildPath "DocxToTxtConversion.log"
"Conversion started at $(Get-Date)" | Out-File -FilePath $logFile
"Parameters: SourceDir=$SourceDir, SourceFiles=$($SourceFiles -join ','), DeleteOriginal=$DeleteOriginal, Backup=$Backup, Silent=$Silent" | Out-File -FilePath $logFile -Append

# Iterate through each .docx file for conversion
foreach ($file in $docxFiles) {
    $docxPath = $file.FullName
    $txtPath = [System.IO.Path]::ChangeExtension($docxPath, ".txt")
    
    # Check if the target file already exists
    if (Test-Path -Path $txtPath -PathType Leaf) {
        $skippedFiles++
        $message = "Skipped file: $($file.Name) - Target file already exists"
        if (-not $Silent) { Write-Host $message -ForegroundColor Yellow }
        $message | Out-File -FilePath $logFile -Append
        continue
    }
    
    if (-not $Silent) { Write-Host "Converting file: $($file.Name) ..." -ForegroundColor White }
    
    # If backup option is enabled, backup the original file
    if ($Backup) {
        # For absolute paths in source, create a relative path for backup
        $relPath = $file.FullName
        if ([System.IO.Path]::IsPathRooted($SourceDir)) {
            # Try to create a simpler structure in the backup
            $relPath = $file.Name
        }
        else {
            # Use relative path within the source directory
            $relPath = $file.FullName.Substring($docxDir.Length)
        }
        
        $backupFilePath = Join-Path -Path $backupDir -ChildPath $relPath
        $backupFileDir = Split-Path -Parent -Path $backupFilePath
        
        # Ensure the target directory exists
        if (-not (Test-Path -Path $backupFileDir -PathType Container)) {
            New-Item -Path $backupFileDir -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $docxPath -Destination $backupFilePath -Force
        if (-not $Silent) { Write-Host "  Backed up to: $backupFilePath" -ForegroundColor DarkGray }
    }
    
    try {
        # Open the file
        $doc = $word.Documents.Open($docxPath)
        
        # Save as .txt format (wdFormatText = 2)
        $saveFormat = 2
        $doc.SaveAs([ref]$txtPath, [ref]$saveFormat)
        
        # Close the file
        $doc.Close()
        
        $convertedFiles++
        $message = "Success: $docxPath -> $txtPath"
        $message | Out-File -FilePath $logFile -Append
        if (-not $Silent) { Write-Host "  Successfully converted: $($file.Name)" -ForegroundColor Green }
        
        # If original files should be deleted
        if ($DeleteOriginal) {
            Remove-Item -Path $docxPath -Force
            if (-not $Silent) { Write-Host "  Original file deleted" -ForegroundColor DarkGray }
        }
    }
    catch {
        $failedFiles++
        $message = "Failed: $docxPath - Error: $_"
        $message | Out-File -FilePath $logFile -Append
        if (-not $Silent) { Write-Host "  Conversion failed: $($file.Name) - Error: $_" -ForegroundColor Red }
    }
}

# Close the Word application
$word.Quit()

# Release COM objects
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

# Display the conversion result summary
$summary = @"
Conversion completed! 
Total files: $totalFiles
Successfully converted: $convertedFiles
Skipped: $skippedFiles
Failed: $failedFiles
See log file for details: $logFile
"@

Write-Host $summary -ForegroundColor Cyan
$summary | Out-File -FilePath $logFile -Append

if ($Backup) {
    Write-Host "Backup directory: $backupDir" -ForegroundColor Cyan
    "Backup directory: $backupDir" | Out-File -FilePath $logFile -Append
}

Write-Host "Press any key to continue..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") 