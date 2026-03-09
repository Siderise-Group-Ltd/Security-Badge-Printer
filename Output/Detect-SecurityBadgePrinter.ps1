<#
.SYNOPSIS
Detection script for Siderise Security Badge Printer

.DESCRIPTION
Checks if the application is installed and verifies the version.
Returns exit code 0 if the application is installed with version 1.1.3 or higher.
Returns exit code 1 if not installed or version is lower.
#>

$installPath = "$env:ProgramFiles\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
$requiredVersion = [Version]"1.1.3"

if (Test-Path $installPath) {
    $installedVersion = (Get-Item $installPath).VersionInfo.FileVersion
    if ($installedVersion) {
        $installedVer = [Version]($installedVersion -replace '^(\d+\.\d+\.\d+).*','$1')
        if ($installedVer -ge $requiredVersion) {
            Write-Host "Siderise Security Badge Printer version $installedVer is installed (required: $requiredVersion)"
            Exit 0
        } else {
            Write-Host "Installed version $installedVer is lower than required version $requiredVersion"
            Exit 1
        }
    }
}

Write-Host "Siderise Security Badge Printer is not installed"
Exit 1
