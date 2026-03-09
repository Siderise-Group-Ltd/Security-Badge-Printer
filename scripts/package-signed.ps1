<#
Package and sign Siderise Security Badge Printer
- Publishes a self-contained Release (win-x64)
- Signs EXE/DLLs with a PFX imported to CurrentUser store
- Builds an Inno Setup installer EXE
- Signs the installer

Usage examples:
  powershell -ExecutionPolicy Bypass -File .\scripts\package-signed.ps1 \
    -PfxPath "C:\Users\JEJ\OneDrive - SIDERISE\Documents\Projects\Card Printer\SideriseBadgePrinterCert.pfx"

  # If you want to keep the imported cert in your user store:
  powershell -ExecutionPolicy Bypass -File .\scripts\package-signed.ps1 -KeepImportedCert 
#>

[CmdletBinding()]
param(
  [string]$ProjectRoot,
  [string]$PfxPath,
  [SecureString]$PfxPassword,
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64",
  [string]$TimestampUrl = "http://timestamp.digicert.com",
  [switch]$KeepImportedCert
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ProjectRoot {
  # default: parent of this script folder
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
  throw "signtool.exe not found. Install Windows 10/11 SDK or ensure signtool is on PATH."
}

function Get-ISCC {
  $ci = Get-Command iscc.exe -ErrorAction SilentlyContinue
  if ($ci) {
    $p = $ci.Path
    if (-not $p -and ($ci.PSObject.Properties["Definition"])) { $p = $ci.Definition }
    if ($p) { return $p }
  }
  $default = "C:\\Program Files (x86)\\Inno Setup 6\\ISCC.exe"
  if (Test-Path $default) { return $default }
  throw "Inno Setup compiler (iscc.exe) not found. Install Inno Setup (winget install JRSoftware.InnoSetup) or add to PATH."
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
  & $SignTool sign /fd sha256 /tr $Timestamp /td sha256 /sha1 $Thumb $Path | Out-Null
  if ($LASTEXITCODE -ne 0) { throw "SignTool failed for: $Path" }
}

# --- Begin ---
$root = Resolve-ProjectRoot
$proj = Join-Path $root "SecurityBadgePrinter\SecurityBadgePrinter.csproj"
if (-not (Test-Path $proj)) { throw "Project file not found: $proj" }

if (-not $PfxPath) {
  $candidate = Join-Path $root "SideriseBadgePrinterCert.pfx"
  if (Test-Path $candidate) { $PfxPath = $candidate }
  else { throw "Specify -PfxPath to your code-signing PFX." }
}

$OutDir = Join-Path $root "out\publish-$Runtime"
if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutDir | Out-Null

Write-Host "[1/7] Importing certificate..." -ForegroundColor Cyan
$thumb = Import-PfxToUserStore -Path $PfxPath -Password $PfxPassword
Write-Host "    Thumbprint: $thumb"

Write-Host "[2/7] Publishing $Configuration ($Runtime, self-contained)..." -ForegroundColor Cyan
& dotnet publish $proj -c $Configuration -r $Runtime --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false `
  -o $OutDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }

$exePath = Join-Path $OutDir "SecurityBadgePrinter.exe"
if (-not (Test-Path $exePath)) { throw "Built EXE not found: $exePath" }

Write-Host "[3/7] Locating SignTool..." -ForegroundColor Cyan
$signTool = Get-SignTool
Write-Host "    SignTool: $signTool"

Write-Host "[4/7] Code-signing app binaries..." -ForegroundColor Cyan
Get-ChildItem $OutDir -Include *.exe,*.dll -Recurse | ForEach-Object {
  Invoke-Sign -SignTool $signTool -Thumb $thumb -Path $_.FullName -Timestamp $TimestampUrl
}

Write-Host "[5/7] Building Inno Setup installer..." -ForegroundColor Cyan
$ISCC = Get-ISCC
$installerDir = Join-Path $root "installer"
New-Item -ItemType Directory -Path $installerDir -Force | Out-Null

# Derive 3-part version from file version
$ver = (Get-Item $exePath).VersionInfo.FileVersion
if (-not $ver) { $ver = "1.0.0.0" }
$ver3 = ($ver -replace '^(\d+\.\d+\.\d+).*','$1')

$iss = Join-Path $installerDir "setup.iss"
$issContent = @"
[Setup]
AppId={{A3E94F4F-476C-4E07-8661-5E2A0C76E8B3}
AppName=Siderise Security Badge Printer
AppVersion=$ver3
AppPublisher=Siderise
DefaultDirName={commonpf}\Siderise\Security Badge Printer
DefaultGroupName=Siderise
OutputBaseFilename=Setup_Siderise_SecurityBadgePrinter
OutputDir=$installerDir\Output
ArchitecturesInstallIn64BitMode=x64
Compression=lzma2
SolidCompression=yes
DisableProgramGroupPage=yes
WizardStyle=modern

[Files]
Source: "$OutDir\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion

[Icons]
Name: "{commonprograms}\Siderise\Security Badge Printer"; Filename: "{app}\SecurityBadgePrinter.exe"

[Run]
Filename: "{app}\SecurityBadgePrinter.exe"; Description: "Launch Siderise Security Badge Printer"; Flags: nowait postinstall skipifsilent
"@
$issContent | Set-Content -Path $iss -Encoding UTF8

& $ISCC $iss
if ($LASTEXITCODE -ne 0) { throw "Inno Setup compilation failed." }

$setupPath = Join-Path $installerDir "Output\Setup_Siderise_SecurityBadgePrinter.exe"
if (-not (Test-Path $setupPath)) { throw "Installer not found: $setupPath" }

Write-Host "[6/7] Code-signing installer..." -ForegroundColor Cyan
Invoke-Sign -SignTool $signTool -Thumb $thumb -Path $setupPath -Timestamp $TimestampUrl

if (-not $KeepImportedCert) {
  Write-Host "[7/7] Removing imported certificate from user store..." -ForegroundColor Cyan
  $certPath = "Cert:\CurrentUser\My\$thumb"
  if (Test-Path $certPath) { Remove-Item $certPath -Force }
}

Write-Host "\nSUCCESS" -ForegroundColor Green
Write-Host "Installer: $setupPath" -ForegroundColor Green
