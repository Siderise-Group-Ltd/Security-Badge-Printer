# Manual Certificate Installation for Siderise Security Badge Printer

## Certificate File
`SideriseBadgePrinter.cer` (located in project root)

## Option 1: PowerShell Command (Run as Administrator)

```powershell
# Import to Trusted Root Certification Authorities
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2("C:\Users\JEJ\OneDrive - SIDERISE\Documents\Projects\Card Printer\SideriseBadgePrinter.cer")
$rootStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("Root", "LocalMachine")
$rootStore.Open("ReadWrite")
$rootStore.Add($cert)
$rootStore.Close()

# Import to Trusted Publishers
$pubStore = New-Object System.Security.Cryptography.X509Certificates.X509Store("TrustedPublisher", "LocalMachine")
$pubStore.Open("ReadWrite")
$pubStore.Add($cert)
$pubStore.Close()

Write-Host "Certificate installed successfully"
```

## Option 2: GUI Method

1. **Right-click** `SideriseBadgePrinter.cer` → **Install Certificate**
2. Select **Local Machine** → Click **Next**
3. Select **Place all certificates in the following store** → Click **Browse**
4. Select **Trusted Root Certification Authorities** → Click **OK** → **Next** → **Finish**
5. Repeat steps 1-4, but in step 4 select **Trusted Publishers** instead

## Option 3: Group Policy / Intune

Deploy the certificate via:
- **Intune**: Devices > Configuration profiles > Create profile > Templates > Trusted certificate
- **GPO**: Computer Configuration > Policies > Windows Settings > Security Settings > Public Key Policies > Trusted Root/Publishers

## Verification

After installation, verify with:
```powershell
Get-ChildItem Cert:\LocalMachine\Root | Where-Object { $_.Subject -like "*Siderise*" }
Get-ChildItem Cert:\LocalMachine\TrustedPublisher | Where-Object { $_.Subject -like "*Siderise*" }
```

Both commands should return the certificate with subject: `CN=Siderise Security Badge Printer`
