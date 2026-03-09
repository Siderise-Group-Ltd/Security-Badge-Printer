<#
Build a signed portable ZIP for Siderise Security Badge Printer (no installer)
- Publishes self-contained Release (win-x64 by default)
- Signs EXE/DLLs using a PFX imported to CurrentUser store
- Zips the output to installer/Output/Siderise_SecurityBadgePrinter_Portable.zip

Usage:
  powershell -ExecutionPolicy Bypass -File .\scripts\package-portable-signed.ps1 `
    -PfxPath "C:\Users\JEJ\OneDrive - SIDERISE\Documents\Projects\Card Printer\SideriseBadgePrinterCert.pfx"

Optional:
  -Configuration Release|Debug
  -Runtime win-x64|win-x86|win-arm64
  -OutputZip "C:\path\to\custom.zip"
  -KeepImportedCert  # keep the imported cert in your user store after signing
#>

[CmdletBinding()]
param(
  [string]$ProjectRoot,
  [string]$PfxPath,
  [SecureString]$PfxPassword,
  [string]$Configuration = "Release",
  [string]$Runtime = "win-x64",
  [string]$TimestampUrl = "http://timestamp.digicert.com",
  [string]$OutputZip,
  [switch]$KeepImportedCert
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

function Resolve-ProjectRoot {
  if ($ProjectRoot) { return (Resolve-Path $ProjectRoot).Path }
  return (Resolve-Path (Join-Path $PSScriptRoot ".." )).Path
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
  if (Test-Path $candidate) { $PfxPath = $candidate } else { throw "Specify -PfxPath to your code-signing PFX." }
}

$assetsDir = Join-Path $root "SecurityBadgePrinter\Assets"
$icoLogo = Join-Path $assetsDir "logo.ico"
$icoPrimary = Join-Path $assetsDir "AppIcon.ico"
$icoAlt = Join-Path $assetsDir "favicon.ico"
$pngFavicon = Join-Path $assetsDir "Favicon.png"

# If no .ico exists but a PNG favicon does, try to generate an .ico via ImageMagick (magick.exe)
if (-not (Test-Path $icoLogo) -and -not (Test-Path $icoPrimary) -and -not (Test-Path $icoAlt) -and (Test-Path $pngFavicon)) {
  $magick = (Get-Command magick.exe -ErrorAction SilentlyContinue).Path
  if ($magick) {
    Write-Host "[0/5] Generating AppIcon.ico from Favicon.png..." -ForegroundColor Cyan
    & $magick $pngFavicon -define icon:auto-resize=256,128,64,48,32,16 $icoPrimary
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path $icoPrimary)) {
      Write-Warning "Could not generate ICO automatically. Continuing without EXE icon. Place an .ico in Assets and re-run."
    }
  } else {
    Write-Warning "ImageMagick not found (magick.exe). To embed an EXE icon, install ImageMagick or place an .ico in SecurityBadgePrinter/Assets."
  }
}

$OutDir = Join-Path $root "out\publish-$Runtime"
if (Test-Path $OutDir) { Remove-Item $OutDir -Recurse -Force }
New-Item -ItemType Directory -Path $OutDir | Out-Null

if (-not $OutputZip) {
  $OutputZip = Join-Path $root "installer\Output\Siderise Security Badge Printer.zip"
}
$zipDir = Split-Path $OutputZip -Parent
New-Item -ItemType Directory -Force -Path $zipDir | Out-Null

Write-Host "[1/5] Importing certificate..." -ForegroundColor Cyan
$thumb = Import-PfxToUserStore -Path $PfxPath -Password $PfxPassword
Write-Host "    Thumbprint: $thumb"

Write-Host "[2/5] Publishing $Configuration ($Runtime, self-contained)..." -ForegroundColor Cyan
$appIconProp = @()
if (Test-Path $icoLogo) { $appIconProp = @('-p:ApplicationIcon=' + $icoLogo) }
elseif (Test-Path $icoPrimary) { $appIconProp = @('-p:ApplicationIcon=' + $icoPrimary) }
elseif (Test-Path $icoAlt) { $appIconProp = @('-p:ApplicationIcon=' + $icoAlt) }
else { Write-Warning "No .ico found in Assets. EXE will use default icon. Place logo.ico, AppIcon.ico or favicon.ico and re-run." }

& dotnet publish $proj -c $Configuration -r $Runtime --self-contained true `
  -p:PublishSingleFile=true `
  -p:IncludeNativeLibrariesForSelfExtract=true `
  -p:PublishTrimmed=false `
  @appIconProp `
  -o $OutDir
if ($LASTEXITCODE -ne 0) { throw "dotnet publish failed." }

$desiredExeName = "Security Badge Printer.exe"
$desiredExePath = Join-Path $OutDir $desiredExeName

# Try to rename the primary exe to the desired name
$defaultExePath = Join-Path $OutDir "SecurityBadgePrinter.exe"
if (Test-Path $defaultExePath) {
  Rename-Item -Path $defaultExePath -NewName $desiredExeName -Force
}
elseif (-not (Test-Path $desiredExePath)) {
  # Fallback: pick the first top-level exe in OutDir (single-file publish produces one)
  $anyExe = Get-ChildItem $OutDir -File -Filter *.exe | Select-Object -First 1
  if ($anyExe) { Rename-Item -Path $anyExe.FullName -NewName $desiredExeName -Force }
}

$exePath = $desiredExePath
if (-not (Test-Path $exePath)) { throw "Built EXE not found after publish: $exePath" }

Write-Host "[3/5] Locating SignTool..." -ForegroundColor Cyan
$signTool = Get-SignTool
Write-Host "    SignTool: $signTool"

Write-Host "[4/5] Code-signing app binaries..." -ForegroundColor Cyan
Get-ChildItem $OutDir -Include *.exe,*.dll -Recurse | ForEach-Object {
  Invoke-Sign -SignTool $signTool -Thumb $thumb -Path $_.FullName -Timestamp $TimestampUrl
}

Write-Host "[5/5] Creating portable ZIP..." -ForegroundColor Cyan
if (Test-Path $OutputZip) { Remove-Item $OutputZip -Force }

# Copy to a staging folder to avoid any transient file locks in the publish dir
$stageDir = Join-Path $root "out\stage-portable-$Runtime"
if (Test-Path $stageDir) { Remove-Item $stageDir -Recurse -Force }
New-Item -ItemType Directory -Path $stageDir | Out-Null
Copy-Item -Path (Join-Path $OutDir "*") -Destination $stageDir -Recurse -Force

# Zip from staging
Compress-Archive -Path (Join-Path $stageDir "*") -DestinationPath $OutputZip -Force

# Clean up staging
Remove-Item $stageDir -Recurse -Force

if (-not $KeepImportedCert) {
  $certPath = "Cert:\CurrentUser\My\$thumb"
  if (Test-Path $certPath) { Remove-Item $certPath -Force }
}

Write-Host "\nSUCCESS" -ForegroundColor Green
Write-Host "Portable ZIP: $OutputZip" -ForegroundColor Green
