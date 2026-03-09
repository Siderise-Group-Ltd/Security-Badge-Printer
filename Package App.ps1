# Package App.ps1
# Run from: C:\Users\JEJ\OneDrive - SIDERISE\Documents\Projects\Card Printer
# Purpose: Executes package-portable-signed.ps1 with the certificate password securely loaded

# ------------------------------------------
# INITIAL SETUP (run once before using):
# ------------------------------------------
# In PowerShell, run this ONCE on the same machine & user account:
# Read-Host "Enter certificate password" -AsSecureString | ConvertFrom-SecureString | Out-File "$env:USERPROFILE\badge_cert_password.txt"
#
# This will create an encrypted password file specific to your Windows user profile.
# Only that same user on that same machine can decrypt it.
# ------------------------------------------

# Resolve script directory (handles spaces in path)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Paths
$pfxPath    = Join-Path $scriptDir "SideriseBadgePrinterCert.pfx"
$packagePs1 = Join-Path $scriptDir "scripts\package-portable-signed.ps1"

# Password file path
$pwdFile = Join-Path $env:USERPROFILE "badge_cert_password.txt"

# Validate password file existence
if (-not (Test-Path $pwdFile)) {
    Write-Host "❌ Password file not found at '$pwdFile'." -ForegroundColor Red
    Write-Host "Run the setup command shown at the top of this script to create it." -ForegroundColor Yellow
    exit 1
}

# Load and decrypt password
$pfxPassword = Get-Content $pwdFile | ConvertTo-SecureString

# Run the packaging script in the SAME PowerShell session
Write-Host "🚀 Running package-portable-signed.ps1..." -ForegroundColor Cyan
& $packagePs1 -PfxPath $pfxPath -PfxPassword $pfxPassword -KeepImportedCert

# Check exit code and report
if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Packaging completed successfully." -ForegroundColor Green
} else {
    Write-Host "❌ Packaging failed with exit code $LASTEXITCODE." -ForegroundColor Red
    exit $LASTEXITCODE
}
