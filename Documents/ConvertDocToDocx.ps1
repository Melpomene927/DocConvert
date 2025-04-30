# PowerShell Script: Convert .doc files to .docx format in DOC directory and subdirectories
# Author: AI Assistant
# Date: Auto-generated

param (
    [string]$SourceDir = ".",
    [switch]$DeleteOriginal = $false,
    [switch]$Backup = $true,
    [switch]$Silent = $false
)

# Set script execution path to current script directory
$scriptPath = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
Set-Location -Path $scriptPath

# Set source directory path (handle both relative and absolute paths)
$docDir = $SourceDir
if (-not [System.IO.Path]::IsPathRooted($SourceDir)) {
    $docDir = Join-Path -Path $scriptPath -ChildPath $SourceDir
}

# Check if source directory exists
if (-not (Test-Path -Path $docDir -PathType Container)) {
    Write-Error "Directory $SourceDir does not exist: $docDir"
    exit 1
}

# Initialize counters
$totalFiles = 0
$convertedFiles = 0
$failedFiles = 0
$skippedFiles = 0

# If backup option is enabled, create a backup directory
$backupDir = $null
if ($Backup) {
    $backupDirName = "DocBackup_" + (Get-Date -Format "yyyyMMdd_HHmmss")
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

# Recursively get all .doc files
$docFiles = Get-ChildItem -Path $docDir -Filter "*.doc" -Recurse -File

# Show the number of files found
$totalFiles = $docFiles.Count
Write-Host "Found $totalFiles .doc files in $SourceDir directory and its subdirectories" -ForegroundColor Yellow

# Create a log file to record conversion results
$logFile = Join-Path -Path $scriptPath -ChildPath "DocConversion.log"
"Conversion started at $(Get-Date)" | Out-File -FilePath $logFile
"Parameters: SourceDir=$SourceDir, DeleteOriginal=$DeleteOriginal, Backup=$Backup, Silent=$Silent" | Out-File -FilePath $logFile -Append

# Iterate through each .doc file for conversion
foreach ($file in $docFiles) {
    $docPath = $file.FullName
    $docxPath = [System.IO.Path]::ChangeExtension($docPath, ".docx")
    
    # Check if the target file already exists
    if (Test-Path -Path $docxPath -PathType Leaf) {
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
            $relPath = $file.FullName.Substring($docDir.Length)
        }
        
        $backupFilePath = Join-Path -Path $backupDir -ChildPath $relPath
        $backupFileDir = Split-Path -Parent -Path $backupFilePath
        
        # Ensure the target directory exists
        if (-not (Test-Path -Path $backupFileDir -PathType Container)) {
            New-Item -Path $backupFileDir -ItemType Directory -Force | Out-Null
        }
        
        Copy-Item -Path $docPath -Destination $backupFilePath -Force
        if (-not $Silent) { Write-Host "  Backed up to: $backupFilePath" -ForegroundColor DarkGray }
    }
    
    try {
        # Open the file
        $doc = $word.Documents.Open($docPath)
        
        # Save as .docx format (16 = wdFormatDocumentDefault = .docx format)
        $saveFormat = 16
        $doc.SaveAs([ref]$docxPath, [ref]$saveFormat)
        
        # Close the file
        $doc.Close()
        
        $convertedFiles++
        $message = "Success: $docPath -> $docxPath"
        $message | Out-File -FilePath $logFile -Append
        if (-not $Silent) { Write-Host "  Successfully converted: $($file.Name)" -ForegroundColor Green }
        
        # If original files should be deleted
        if ($DeleteOriginal) {
            Remove-Item -Path $docPath -Force
            if (-not $Silent) { Write-Host "  Original file deleted" -ForegroundColor DarkGray }
        }
    }
    catch {
        $failedFiles++
        $message = "Failed: $docPath - Error: $_"
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