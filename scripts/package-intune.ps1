<#
.SYNOPSIS
Package Siderise Security Badge Printer for Microsoft Intune deployment

.DESCRIPTION
This script:
1. Builds the application in Release mode (self-contained, win-x64)
2. Optionally signs the executables with a code signing certificate
3. Packages it using IntuneWinAppUtil.exe
4. Creates a version-based detection script
5. Outputs both files to the Output folder for Intune deployment

.PARAMETER ProjectRoot
Root directory of the project. Defaults to parent of scripts folder.

.PARAMETER Configuration
Build configuration. Default: Release

.PARAMETER Runtime
Target runtime. Default: win-x64

.PARAMETER PfxPath
Path to the code signing certificate (PFX file). If not provided, looks for SideriseBadgePrinterCert.pfx in project root.

.PARAMETER PfxPassword
Password for the PFX file. If not provided, will prompt.

.PARAMETER SkipSigning
Skip code signing (not recommended for production).

.PARAMETER TimestampUrl
Timestamp server URL for code signing. Default: http://timestamp.digicert.com

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\scripts\package-intune.ps1

.EXAMPLE
powershell -ExecutionPolicy Bypass -File .\scripts\package-intune.ps1 -PfxPath "C:\certs\mycert.pfx"
#>

[CmdletBinding()]
param(
  [string]$ProjectRoot,
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64",
  [string]$PfxPath,
  [SecureString]$PfxPassword,
  [switch]$SkipSigning,
  [string]$TimestampUrl = "http://timestamp.digicert.com"
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ProjectRoot {
  return (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

function Get-SignTool {
  $ci = Get-Command signtool.exe -ErrorAction SilentlyContinue
  if ($ci) {
    $p = $ci.Path
    if (-not $p -and ($ci.PSObject.Properties["Definition"])) { $p = $ci.Definition }
    if ($p) { return $p }
  }
  $probable = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Filter signtool.exe -Recurse -ErrorAction SilentlyContinue |
              Where-Object { $_.FullName -match "\\x64\\signtool\.exe$" } |
              Sort-Object FullName -Descending | Select-Object -First 1
  if ($probable) { return $probable.FullName }
  return $null
}

function Import-PfxToUserStore {
  param([string]$Path,[SecureString]$Password)
  if (-not (Test-Path $Path)) { throw "PFX not found: $Path" }
  if (-not $Password) { $Password = Read-Host -Prompt "Enter PFX password" -AsSecureString }
  $cert = Import-PfxCertificate -FilePath $Path -Password $Password -CertStoreLocation Cert:\CurrentUser\My
  if (-not $cert) { throw "Failed to import PFX certificate." }
  return $cert.Thumbprint
}

function Invoke-Sign {
  param([string]$SignTool,[string]$Thumb,[string]$Path,[string]$Timestamp)
  $output = & $SignTool sign /fd sha256 /sha1 $Thumb $Path 2>&1
  if ($LASTEXITCODE -ne 0) { 
    Write-Host "SignTool output: $output" -ForegroundColor Red
    throw "SignTool failed for: $Path" 
  }
}

# --- Begin ---
Write-Host "=== Siderise Security Badge Printer - Intune Packaging ===" -ForegroundColor Cyan
Write-Host ""

$root = if ($ProjectRoot) { $ProjectRoot } else { Resolve-ProjectRoot }
$proj = Join-Path $root "SecurityBadgePrinter\SecurityBadgePrinter.csproj"
if (-not (Test-Path $proj)) { throw "Project file not found: $proj" }

# Output directories - use temp to avoid OneDrive locking
$tempBase = Join-Path $env:TEMP "SideriseBadgePrinter-Build"
$publishDir = Join-Path $tempBase "intune-publish"
$sourceDir = Join-Path $publishDir "Source"
$outputDir = Join-Path $root "Output"

# Clean and create directories
if (Test-Path $tempBase) { Remove-Item $tempBase -Recurse -Force }
if (Test-Path $outputDir) { Remove-Item $outputDir -Recurse -Force }
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null
New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
New-Item -ItemType Directory -Path $sourceDir -Force | Out-Null

Write-Host "[1/7] Publishing application ($Configuration, $Runtime)..." -ForegroundColor Cyan
& dotnet publish $proj -c $Configuration -r $Runtime --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false `
  -o $sourceDir

if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }

$exePath = Join-Path $sourceDir "SecurityBadgePrinter.exe"
if (-not (Test-Path $exePath)) { throw "Built EXE not found: $exePath" }

# Get version from the built executable
Write-Host "[2/7] Reading version information..." -ForegroundColor Cyan
$ver = (Get-Item $exePath).VersionInfo.FileVersion
if (-not $ver) { $ver = "1.0.0.0" }
$ver3 = ($ver -replace '^(\d+\.\d+\.\d+).*','$1')
Write-Host "    Version: $ver3" -ForegroundColor Green

# Code signing
$thumbprint = $null
if (-not $SkipSigning) {
    Write-Host "[3/7] Code signing application..." -ForegroundColor Cyan
    
    # Find or prompt for PFX
    if (-not $PfxPath) {
        $candidate = Join-Path $root "SideriseBadgePrinterCert.pfx"
        if (Test-Path $candidate) { 
            $PfxPath = $candidate 
        } else {
            Write-Host "    WARNING: No PFX certificate found. Use -PfxPath to specify one." -ForegroundColor Yellow
            Write-Host "    Skipping code signing..." -ForegroundColor Yellow
        }
    }
    
    if ($PfxPath -and (Test-Path $PfxPath)) {
        $signTool = Get-SignTool
        if (-not $signTool) {
            Write-Host "    WARNING: signtool.exe not found. Install Windows SDK." -ForegroundColor Yellow
            Write-Host "    Skipping code signing..." -ForegroundColor Yellow
        } else {
            Write-Host "    SignTool: $signTool" -ForegroundColor Gray
            Write-Host "    Certificate: $PfxPath" -ForegroundColor Gray
            
            # Import certificate
            $thumbprint = Import-PfxToUserStore -Path $PfxPath -Password $PfxPassword
            Write-Host "    Thumbprint: $thumbprint" -ForegroundColor Gray
            
            # Sign all EXE and DLL files
            $filesToSign = Get-ChildItem $sourceDir -Include *.exe,*.dll -Recurse
            $signedCount = 0
            foreach ($file in $filesToSign) {
                try {
                    Invoke-Sign -SignTool $signTool -Thumb $thumbprint -Path $file.FullName -Timestamp $TimestampUrl
                    $signedCount++
                } catch {
                    Write-Host "    WARNING: Failed to sign $($file.Name): $_" -ForegroundColor Yellow
                }
            }
            Write-Host "    Signed $signedCount files" -ForegroundColor Green
        }
    }
} else {
    Write-Host "[3/7] Skipping code signing (-SkipSigning specified)..." -ForegroundColor Yellow
}

# Create install script
Write-Host "[4/7] Creating install script..." -ForegroundColor Cyan
$installScript = @"
# Siderise Security Badge Printer - Install Script
# Run with: powershell.exe -ExecutionPolicy Bypass -File Install.ps1

`$ErrorActionPreference = 'Continue'

# Force 64-bit Program Files path (avoid x86 redirection)
`$programFiles = "C:\Program Files"
`$installPath = "`$programFiles\Siderise\Security Badge Printer"
`$logPath = "`$env:ProgramData\Siderise\Logs"
`$logFile = "`$logPath\SecurityBadgePrinter_Install.log"

# Ensure log directory exists
try {
    if (-not (Test-Path `$logPath)) { New-Item -ItemType Directory -Path `$logPath -Force | Out-Null }
} catch {
    `$logFile = "`$env:TEMP\SecurityBadgePrinter_Install.log"
}

function Write-Log {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    `$logMessage = "[`$timestamp] `$Message"
    try { Add-Content -Path `$logFile -Value `$logMessage -ErrorAction SilentlyContinue } catch {}
    Write-Output `$logMessage
}

Write-Log "=== Siderise Security Badge Printer Installation ==="
Write-Log "Log file: `$logFile"
Write-Log "Install path: `$installPath"

# Create installation directory
try {
    if (-not (Test-Path `$installPath)) {
        New-Item -ItemType Directory -Path `$installPath -Force | Out-Null
        Write-Log "Created directory: `$installPath"
    } else {
        Write-Log "Directory exists: `$installPath"
    }
} catch {
    Write-Log "ERROR creating directory: `$(`$_.Exception.Message)"
    Exit 1
}

# Copy application files
try {
    Write-Log "Copying files from `$PSScriptRoot..."
    `$filesCopied = 0
    Get-ChildItem -Path "`$PSScriptRoot" -File | Where-Object { `$_.Name -notin @("Install.ps1", "Uninstall.ps1") } | ForEach-Object {
        try {
            Copy-Item -Path `$_.FullName -Destination `$installPath -Force -ErrorAction Stop
            Write-Log "  Copied: `$(`$_.Name)"
            `$filesCopied++
        } catch {
            Write-Log "  ERROR copying `$(`$_.Name): `$(`$_.Exception.Message)"
        }
    }
    Write-Log "Total files copied: `$filesCopied"
} catch {
    Write-Log "ERROR during file copy: `$(`$_.Exception.Message)"
    Exit 1
}

# Verify executable
`$exePath = Join-Path `$installPath "SecurityBadgePrinter.exe"
if (-not (Test-Path `$exePath)) {
    Write-Log "ERROR: SecurityBadgePrinter.exe not found at `$exePath"
    Exit 1
}

`$version = (Get-Item `$exePath).VersionInfo.FileVersion
Write-Log "Verified: SecurityBadgePrinter.exe (v`$version)"

# Set permissions for all users to read and execute
try {
    Write-Log "Setting file permissions..."
    `$acl = Get-Acl `$installPath
    `$rule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "ReadAndExecute", "ContainerInherit,ObjectInherit", "None", "Allow")
    `$acl.SetAccessRule(`$rule)
    Set-Acl -Path `$installPath -AclObject `$acl
    Write-Log "Permissions set successfully"
} catch {
    Write-Log "WARNING: Failed to set permissions: `$(`$_.Exception.Message)"
}

# Create Start Menu shortcut
try {
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
    Write-Log "Shortcut created"
} catch {
    Write-Log "WARNING: Shortcut failed: `$(`$_.Exception.Message)"
}

Write-Log "=== Installation completed successfully ==="
Exit 0
"@

$installScript | Set-Content -Path (Join-Path $sourceDir "Install.ps1") -Encoding UTF8

# Create uninstall script
$uninstallScript = @"
# Siderise Security Badge Printer - Uninstall Script
`$ErrorActionPreference = 'Continue'

# Check both possible installation paths
`$installPaths = @(
    "C:\Program Files\Siderise\Security Badge Printer",
    "C:\Program Files (x86)\Siderise\Security Badge Printer"
)

`$logPath = "`$env:ProgramData\Siderise\Logs"
`$logFile = "`$logPath\SecurityBadgePrinter_Uninstall.log"

try {
    if (-not (Test-Path `$logPath)) { New-Item -ItemType Directory -Path `$logPath -Force | Out-Null }
} catch {
    `$logFile = "`$env:TEMP\SecurityBadgePrinter_Uninstall.log"
}

function Write-Log {
    param([string]`$Message)
    `$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    `$logMessage = "[`$timestamp] `$Message"
    try { Add-Content -Path `$logFile -Value `$logMessage -ErrorAction SilentlyContinue } catch {}
    Write-Output `$logMessage
}

Write-Log "=== Uninstallation Started ==="

# Remove shortcuts
`$shortcut = "`$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Siderise\Security Badge Printer.lnk"
if (Test-Path `$shortcut) {
    Remove-Item `$shortcut -Force
    Write-Log "Removed Start Menu shortcut"
}

# Remove from both possible paths
foreach (`$path in `$installPaths) {
    if (Test-Path `$path) {
        try {
            Remove-Item `$path -Recurse -Force
            Write-Log "Removed: `$path"
        } catch {
            Write-Log "ERROR removing `$path: `$(`$_.Exception.Message)"
        }
    }
}

# Clean up empty parent folders
`$parentPaths = @("C:\Program Files\Siderise", "C:\Program Files (x86)\Siderise")
foreach (`$parent in `$parentPaths) {
    if (Test-Path `$parent) {
        `$items = Get-ChildItem `$parent -ErrorAction SilentlyContinue
        if (`$items.Count -eq 0) {
            Remove-Item `$parent -Force
            Write-Log "Removed empty folder: `$parent"
        }
    }
}

Write-Log "=== Uninstallation completed ==="
Exit 0
"@

$uninstallScript | Set-Content -Path (Join-Path $sourceDir "Uninstall.ps1") -Encoding UTF8

# Package with IntuneWinAppUtil
Write-Host "[5/7] Packaging with IntuneWinAppUtil.exe..." -ForegroundColor Cyan
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
Write-Host "[6/7] Creating detection script..." -ForegroundColor Cyan
$detectionScript = @"
<#
.SYNOPSIS
Detection script for Siderise Security Badge Printer

.DESCRIPTION
Checks if the application is installed and verifies the version.
Returns exit code 0 if the application is installed with version $ver exactly.
Returns exit code 1 if not installed or version does not match.
#>

`$requiredVersion = [Version]"$ver"

# Check both possible installation paths (64-bit and 32-bit redirected)
`$installPaths = @(
    "C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe",
    "C:\Program Files (x86)\Siderise\Security Badge Printer\SecurityBadgePrinter.exe",
    "`$env:ProgramFiles\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
)

`$foundPath = `$null
`$foundVersion = `$null

foreach (`$path in `$installPaths) {
    if (Test-Path `$path) {
        `$foundPath = `$path
        `$foundVersion = (Get-Item `$path).VersionInfo.FileVersion
        if (`$foundVersion) {
            try {
                `$installedVer = [Version]`$foundVersion
                if (`$installedVer -eq `$requiredVersion) {
                    Write-Host "Siderise Security Badge Printer version `$installedVer is installed at `$path"
                    Exit 0
                }
            } catch {
                Write-Host "Could not parse version: `$foundVersion"
            }
        }
        break
    }
}

# If we found it but version is wrong
if (`$foundPath) {
    Write-Host "Installed version `$foundVersion does not match required version `$requiredVersion"
    Exit 1
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

## IMPORTANT: Defender ASR Exclusion Required

If your organization uses Microsoft Defender Attack Surface Reduction (ASR) rules, you must add an exclusion for this application. Without this, Defender will block the executable from running.

**Option 1: Via Intune Policy (Recommended for deployment)**
1. Navigate to: Microsoft Intune admin center > Endpoint security > Attack surface reduction
2. Create or edit an ASR policy
3. Add exclusion for: ``C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe``

**Option 2: Via PowerShell (Local testing only)**
```powershell
Add-MpPreference -AttackSurfaceReductionOnlyExclusions "C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
```

## Notes:
- The application is code-signed with a Siderise self-signed certificate
- The application requires Azure AD configuration in appsettings.json
- Ensure the Zebra ZC300 printer driver is installed on target devices
- Users need appropriate Graph API permissions (User.Read, User.Read.All, Group.Read.All)
"@

$instructions | Set-Content -Path (Join-Path $outputDir "DEPLOYMENT_INSTRUCTIONS.md") -Encoding UTF8

# Clean up certificate from store
if ($thumbprint) {
    Write-Host "[7/7] Cleaning up certificate from store..." -ForegroundColor Cyan
    $certPath = "Cert:\CurrentUser\My\$thumbprint"
    if (Test-Path $certPath) { 
        Remove-Item $certPath -Force 
        Write-Host "    Certificate removed" -ForegroundColor Gray
    }
}

Write-Host ""
Write-Host "SUCCESS" -ForegroundColor Green
Write-Host "Output folder: $outputDir" -ForegroundColor Green
Write-Host "  - SecurityBadgePrinter.intunewin" -ForegroundColor White
Write-Host "  - Detect-SecurityBadgePrinter.ps1" -ForegroundColor White
Write-Host "  - DEPLOYMENT_INSTRUCTIONS.md" -ForegroundColor White
Write-Host ""
Write-Host "Version: $ver3" -ForegroundColor Cyan
Write-Host ""
