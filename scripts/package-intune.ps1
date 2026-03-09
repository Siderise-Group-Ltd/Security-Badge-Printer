<#
.SYNOPSIS
Package Siderise Security Badge Printer for Microsoft Intune deployment

.DESCRIPTION
This script:
1. Builds the application in Release mode (self-contained, win-x64)
2. Packages it using IntuneWinAppUtil.exe
3. Creates a version-based detection script
4. Outputs both files to the Output folder for Intune deployment

.PARAMETER ProjectRoot
Root directory of the project. Defaults to parent of scripts folder.

.PARAMETER Configuration
Build configuration. Default: Release

.PARAMETER Runtime
Target runtime. Default: win-x64

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\scripts\package-intune.ps1
#>

[CmdletBinding()]
param(
  [string]$ProjectRoot,
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64"
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ProjectRoot {
  return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

# --- Begin ---
Write-Host "=== Siderise Security Badge Printer - Intune Packaging ===" -ForegroundColor Cyan
Write-Host ""

$root = if ($ProjectRoot) { $ProjectRoot } else { Resolve-ProjectRoot }
$proj = Join-Path $root "SecurityBadgePrinter\SecurityBadgePrinter.csproj"
if (-not (Test-Path $proj)) { throw "Project file not found: $proj" }

# Output directories
$publishDir = Join-Path $root "out\intune-publish"
$outputDir = Join-Path $root "Output"
$sourceDir = Join-Path $publishDir "Source"

# Clean and create directories
if (Test-Path $publishDir) { Remove-Item $publishDir -Recurse -Force }
if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force }
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null

Write-Host "[1/5] Publishing application ($Configuration, $Runtime)..." -ForegroundColor Cyan
& dotnet publish $proj -c $Configuration -r $Runtime --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false `
  -o $sourceDir

if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }

$exePath = Join-Path $sourceDir "SecurityBadgePrinter.exe"
if (-not (Test-Path $exePath)) { throw "Built EXE not found: $exePath" }

# Get version from the built executable
Write-Host "[2/5] Reading version information..." -ForegroundColor Cyan
$ver = (Get-Item $exePath).VersionInfo.FileVersion
if (-not $ver) { $ver = "1.0.0.0" }
$ver3 = ($ver -replace '^(\d+\.\d+\.\d+).*','$1')
Write-Host "    Version: $ver3" -ForegroundColor Green

# Create install script
Write-Host "[3/5] Creating install script..." -ForegroundColor Cyan
$installScript = @"
# Siderise Security Badge Printer - Install Script
`$ErrorActionPreference = 'Stop'
`$logPath = "`$env:ProgramData\Siderise\Logs"
`$logFile = Join-Path `$logPath "SecurityBadgePrinter_Install.log"

function Write-Log {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    `$logMessage = "[`$timestamp] `$Message"
    Write-Host `$logMessage
    if (-not (Test-Path `$logPath)) {
        New-Item -ItemType Directory -Path `$logPath -Force | Out-Null
    }
    Add-Content -Path `$logFile -Value `$logMessage
}

try {
    Write-Log "Starting installation of Siderise Security Badge Printer"
    
    # Installation directory
    `$installPath = "`$env:ProgramFiles\Siderise\Security Badge Printer"
    Write-Log "Installation path: `$installPath"
    
    # Create directory
    if (-not (Test-Path `$installPath)) {
        Write-Log "Creating installation directory..."
        New-Item -ItemType Directory -Path `$installPath -Force | Out-Null
    } else {
        Write-Log "Installation directory already exists"
    }
    
    # Copy files (exclude install/uninstall scripts)
    Write-Log "Copying application files..."
    Get-ChildItem -Path "`$PSScriptRoot" -Exclude "Install.ps1", "Uninstall.ps1" | ForEach-Object {
        Copy-Item -Path `$_.FullName -Destination `$installPath -Recurse -Force
        Write-Log "  Copied: `$(`$_.Name)"
    }
    
    # Verify main executable
    `$exePath = Join-Path `$installPath "SecurityBadgePrinter.exe"
    if (-not (Test-Path `$exePath)) {
        throw "SecurityBadgePrinter.exe not found after copy"
    }
    `$version = (Get-Item `$exePath).VersionInfo.FileVersion
    Write-Log "Verified: SecurityBadgePrinter.exe (v`$version)"
    
    # Create Start Menu shortcut
    Write-Log "Creating Start Menu shortcut..."
    `$startMenuPath = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Siderise"
    if (-not (Test-Path `$startMenuPath)) {
        New-Item -ItemType Directory -Path `$startMenuPath -Force | Out-Null
    }
    
    `$WshShell = New-Object -ComObject WScript.Shell
    `$Shortcut = `$WshShell.CreateShortcut("`$startMenuPath\Security Badge Printer.lnk")
    `$Shortcut.TargetPath = `$exePath
    `$Shortcut.WorkingDirectory = `$installPath
    `$Shortcut.Description = "Siderise Security Badge Printer"
    `$Shortcut.Save()
    Write-Log "Shortcut created successfully"
    
    Write-Log "Installation completed successfully"
    Exit 0
    
} catch {
    Write-Log "ERROR: Installation failed - `$(`$_.Exception.Message)"
    Write-Log "Stack trace: `$(`$_.ScriptStackTrace)"
    Exit 1
}
"@

$installScript | Set-Content -Path (Join-Path $sourceDir "Install.ps1") -Encoding UTF8

# Create uninstall script
$uninstallScript = @"
# Siderise Security Badge Printer - Uninstall Script
`$ErrorActionPreference = 'Stop'
`$logPath = "`$env:ProgramData\Siderise\Logs"
`$logFile = Join-Path `$logPath "SecurityBadgePrinter_Uninstall.log"

function Write-Log {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    `$logMessage = "[`$timestamp] `$Message"
    Write-Host `$logMessage
    if (-not (Test-Path `$logPath)) {
        New-Item -ItemType Directory -Path `$logPath -Force | Out-Null
    }
    Add-Content -Path `$logFile -Value `$logMessage
}

try {
    Write-Log "Starting uninstallation of Siderise Security Badge Printer"
    
    # Installation directory
    `$installPath = "`$env:ProgramFiles\Siderise\Security Badge Printer"
    
    # Remove Start Menu shortcut
    `$shortcut = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Siderise\Security Badge Printer.lnk"
    if (Test-Path `$shortcut) {
        Write-Log "Removing Start Menu shortcut..."
        Remove-Item `$shortcut -Force
        Write-Log "Shortcut removed"
    } else {
        Write-Log "Start Menu shortcut not found (already removed or never created)"
    }
    
    # Remove Siderise Start Menu folder if empty
    `$startMenuPath = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Siderise"
    if (Test-Path `$startMenuPath) {
        `$items = Get-ChildItem `$startMenuPath
        if (`$items.Count -eq 0) {
            Write-Log "Removing empty Siderise Start Menu folder..."
            Remove-Item `$startMenuPath -Force
        }
    }
    
    # Remove installation directory
    if (Test-Path `$installPath) {
        Write-Log "Removing installation directory: `$installPath"
        Remove-Item `$installPath -Recurse -Force
        Write-Log "Installation directory removed"
    } else {
        Write-Log "Installation directory not found (already removed)"
    }
    
    # Remove parent Siderise folder if empty
    `$parentPath = "`$env:ProgramFiles\Siderise"
    if (Test-Path `$parentPath) {
        `$items = Get-ChildItem `$parentPath
        if (`$items.Count -eq 0) {
            Write-Log "Removing empty Siderise folder..."
            Remove-Item `$parentPath -Force
        }
    }
    
    Write-Log "Uninstallation completed successfully"
    Exit 0
    
} catch {
    Write-Log "ERROR: Uninstallation failed - `$(`$_.Exception.Message)"
    Write-Log "Stack trace: `$(`$_.ScriptStackTrace)"
    Exit 1
}
"@

$uninstallScript | Set-Content -Path (Join-Path $sourceDir "Uninstall.ps1") -Encoding UTF8

# Package with IntuneWinAppUtil
Write-Host "[4/5] Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Cyan
$intuneToolPath = "C:\Users\JEJ\OneDrive - SIDERISE\IT & Systems - IT Infrastructure and Operations - IT Infrastructure and Operations\Intune\Intune App Prep Tool\IntuneWinAppUtil.exe"

if (-not (Test-Path $intuneToolPath)) {
    throw "IntuneWinAppUtil.exe not found at: $intuneToolPath"
}

& $intuneToolPath -c $sourceDir -s "SecurityBadgePrinter.exe" -o $publishDir -q
if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil packaging failed." }

# Move .intunewin to Output folder
$intunewinFile = Get-ChildItem $publishDir -Filter "*.intunewin" | Select-Object -First 1
if (-not $intunewinFile) { throw ".intunewin file not created" }

Move-Item $intunewinFile.FullName (Join-Path $outputDir "SecurityBadgePrinter.intunewin") -Force

# Create detection script
Write-Host "[5/5] Creating detection script..." -ForegroundColor Cyan
$detectionScript = @"
<#
.SYNOPSIS
Detection script for Siderise Security Badge Printer

.DESCRIPTION
Checks if the application is installed and verifies the version.
Returns exit code 0 if the application is installed with version $ver3 or higher.
Returns exit code 1 if not installed or version is lower.
#>

`$installPath = "`$env:ProgramFiles\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
`$requiredVersion = [Version]"$ver3"

if (Test-Path `$installPath) {
    `$installedVersion = (Get-Item `$installPath).VersionInfo.FileVersion
    if (`$installedVersion) {
        `$installedVer = [Version](`$installedVersion -replace '^(\d+\.\d+\.\d+).*','`$1')
        if (`$installedVer -ge `$requiredVersion) {
            Write-Host "Siderise Security Badge Printer version `$installedVer is installed (required: `$requiredVersion)"
            Exit 0
        } else {
            Write-Host "Installed version `$installedVer is lower than required version `$requiredVersion"
            Exit 1
        }
    }
}

Write-Host "Siderise Security Badge Printer is not installed"
Exit 1
"@

$detectionScript | Set-Content -Path (Join-Path $outputDir "Detect-SecurityBadgePrinter.ps1") -Encoding UTF8

# Create deployment instructions
$instructions = @"
# Siderise Security Badge Printer - Intune Deployment Instructions

## Version: $ver3

## Files in this package:
- SecurityBadgePrinter.intunewin - The packaged application
- Detect-SecurityBadgePrinter.ps1 - Detection script for version-based deployment

## Deployment Steps:

1. **Upload to Intune**
   - Navigate to: Microsoft Intune admin center > Apps > Windows apps
   - Click: Add > Windows app (Win32)
   - Upload: SecurityBadgePrinter.intunewin

2. **App Information**
   - Name: Siderise Security Badge Printer
   - Description: Professional security badge printing application for Siderise
   - Publisher: Siderise
   - Version: $ver3

3. **Program Configuration**
   - Install command: 
     ``````
     powershell.exe -ExecutionPolicy Bypass -File ".\Install.ps1"
     ``````
   
   - Uninstall command:
     ``````
     powershell.exe -ExecutionPolicy Bypass -File "%ProgramFiles%\Siderise\Security Badge Printer\Uninstall.ps1"
     ``````
   
   - Install behavior: System
   - Device restart behavior: Determine behavior based on return codes

4. **Requirements**
   - Operating system architecture: 64-bit
   - Minimum operating system: Windows 10 1809

5. **Detection Rules**
   - Rule type: Use a custom detection script
   - Script file: Upload Detect-SecurityBadgePrinter.ps1
   - Run script as 32-bit process: No
   - Enforce script signature check: No

6. **Assignments**
   - Assign to appropriate groups (Required/Available)

## Updating the Application:

To deploy a new version:
1. Update the version in SecurityBadgePrinter.csproj
2. Run this packaging script again
3. Upload the new .intunewin file to Intune
4. The detection script will automatically detect version differences
5. Intune will update devices to the new version

## Notes:
- The application requires Azure AD configuration in appsettings.json
- Ensure the Zebra ZC300 printer driver is installed on target devices
- Users need appropriate Graph API permissions (User.Read, User.Read.All, Group.Read.All)
"@

$instructions | Set-Content -Path (Join-Path $outputDir "DEPLOYMENT_INSTRUCTIONS.md") -Encoding UTF8

Write-Host ""
Write-Host "SUCCESS" -ForegroundColor Green
Write-Host "Output folder: $outputDir" -ForegroundColor Green
Write-Host "  - SecurityBadgePrinter.intunewin" -ForegroundColor White
Write-Host "  - Detect-SecurityBadgePrinter.ps1" -ForegroundColor White
Write-Host "  - DEPLOYMENT_INSTRUCTIONS.md" -ForegroundColor White
Write-Host ""
Write-Host "Version: $ver3" -ForegroundColor Cyan
Write-Host ""
