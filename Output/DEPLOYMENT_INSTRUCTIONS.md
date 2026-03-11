# Siderise Security Badge Printer - Intune Deployment Instructions

## Version: 1.2.5

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
   - Version: 1.2.5

3. **Program Configuration**
   - Install command: 
     ```
     powershell.exe -ExecutionPolicy Bypass -File ".\Install.ps1"
     ```
   
   - Uninstall command:
     ```
     powershell.exe -ExecutionPolicy Bypass -File "%ProgramFiles%\Siderise\Security Badge Printer\Uninstall.ps1"
     ```
   
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
3. Add exclusion for: `C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe`

**Option 2: Via PowerShell (Local testing only)**
`powershell
Add-MpPreference -AttackSurfaceReductionOnlyExclusions "C:\Program Files\Siderise\Security Badge Printer\SecurityBadgePrinter.exe"
`

## Notes:
- The application is code-signed with a Siderise self-signed certificate
- The application requires Azure AD configuration in appsettings.json
- Ensure the Zebra ZC300 printer driver is installed on target devices
- Users need appropriate Graph API permissions (User.Read, User.Read.All, Group.Read.All)
