<#
.SYNOPSIS
Detection script for Siderise Security Badge Printer

.DESCRIPTION
Checks if the application is installed and verifies the version.
Returns exit code 0 if the application is installed with version 1.2.0.0 exactly.
Returns exit code 1 if not installed or version does not match.
#>

$requiredVersion = [Version]"1.2.0.0"

# Check both possible installation paths (64-bit and 32-bit redirected)
$installPaths = @(
    "C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe",
    "C:\Program Files (x86)\Siderise\Security Badge Printer\SecurityBadgePrinter.exe",
    "$env:ProgramFiles\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
)

$foundPath = $null
$foundVersion = $null

foreach ($path in $installPaths) {
    if (Test-Path $path) {
        $foundPath = $path
        $foundVersion = (Get-Item $path).VersionInfo.FileVersion
        if ($foundVersion) {
            try {
                $installedVer = [Version]$foundVersion
                if ($installedVer -eq $requiredVersion) {
                    Write-Host "Siderise Security Badge Printer version $installedVer is installed at $path"
                    Exit 0
                }
            } catch {
                Write-Host "Could not parse version: $foundVersion"
            }
        }
        break
    }
}

# If we found it but version is wrong
if ($foundPath) {
    Write-Host "Installed version $foundVersion does not match required version $requiredVersion"
    Exit 1
}

Write-Host "Siderise Security Badge Printer is not installed"
Exit 1
