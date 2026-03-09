# Test Install Script for Siderise Security Badge Printer
# This simulates the Intune installation process for testing

[CmdletBinding()]
param(
    [string]$SourcePath = "c:\Users\JEJ\OneDrive - SIDERISE\Documents\Projects\Card Printer\out\intune-publish\Source",
    [string]$TestInstallPath = "$env:TEMP\TestInstall_SecurityBadgePrinter"
)

$ErrorActionPreference = 'Stop'

Write-Host "=== Testing Security Badge Printer Installation ===" -ForegroundColor Cyan
Write-Host ""

# Clean up any previous test
if (Test-Path $TestInstallPath) {
    Write-Host "Cleaning up previous test installation..." -ForegroundColor Yellow
    Remove-Item $TestInstallPath -Recurse -Force
}

# Verify source exists
if (-not (Test-Path $SourcePath)) {
    throw "Source path not found: $SourcePath. Run package-intune.ps1 first."
}

Write-Host "Source: $SourcePath" -ForegroundColor White
Write-Host "Test Install Path: $TestInstallPath" -ForegroundColor White
Write-Host ""

# Test the installation steps
try {
    Write-Host "[1/4] Creating installation directory..." -ForegroundColor Cyan
    New-Item -ItemType Directory -Path $TestInstallPath -Force | Out-Null
    Write-Host "    Created: $TestInstallPath" -ForegroundColor Green
    
    Write-Host "[2/4] Copying application files..." -ForegroundColor Cyan
    $filesToCopy = Get-ChildItem -Path $SourcePath -Exclude "Install.ps1", "Uninstall.ps1"
    foreach ($file in $filesToCopy) {
        Copy-Item -Path $file.FullName -Destination $TestInstallPath -Recurse -Force
        Write-Host "    Copied: $($file.Name)" -ForegroundColor Gray
    }
    
    Write-Host "[3/4] Verifying critical files..." -ForegroundColor Cyan
    $exePath = Join-Path $TestInstallPath "SecurityBadgePrinter.exe"
    if (Test-Path $exePath) {
        $version = (Get-Item $exePath).VersionInfo.FileVersion
        Write-Host "    Found: SecurityBadgePrinter.exe (v$version)" -ForegroundColor Green
    } else {
        throw "SecurityBadgePrinter.exe not found in test installation"
    }
    
    $configPath = Join-Path $TestInstallPath "appsettings.json"
    if (Test-Path $configPath) {
        Write-Host "    Found: appsettings.json" -ForegroundColor Green
    } else {
        Write-Host "    WARNING: appsettings.json not found" -ForegroundColor Yellow
    }
    
    Write-Host "[4/4] Testing shortcut creation..." -ForegroundColor Cyan
    $testShortcutPath = Join-Path $env:TEMP "TestShortcut_SecurityBadgePrinter.lnk"
    $WshShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WshShell.CreateShortcut($testShortcutPath)
    $Shortcut.TargetPath = $exePath
    $Shortcut.WorkingDirectory = $TestInstallPath
    $Shortcut.Description = "Siderise Security Badge Printer"
    $Shortcut.Save()
    
    if (Test-Path $testShortcutPath) {
        Write-Host "    Shortcut created successfully" -ForegroundColor Green
        Remove-Item $testShortcutPath -Force
    }
    
    Write-Host ""
    Write-Host "SUCCESS - Installation test completed" -ForegroundColor Green
    Write-Host ""
    Write-Host "Test installation location: $TestInstallPath" -ForegroundColor White
    Write-Host "You can manually test the application from this location." -ForegroundColor White
    Write-Host ""
    
    # Ask if user wants to clean up
    $cleanup = Read-Host "Clean up test installation? (Y/N)"
    if ($cleanup -eq 'Y' -or $cleanup -eq 'y') {
        Remove-Item $TestInstallPath -Recurse -Force
        Write-Host "Test installation cleaned up." -ForegroundColor Green
    }
    
} catch {
    Write-Host ""
    Write-Host "FAILED - Installation test failed" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    
    # Show what files are in the source
    Write-Host "Files in source directory:" -ForegroundColor Yellow
    Get-ChildItem $SourcePath | ForEach-Object {
        Write-Host "  - $($_.Name) ($($_.Length) bytes)" -ForegroundColor Gray
    }
    
    exit 1
}
