# Siderise Security Badge Printer — Project Overview

This repository contains a Windows WPF (.NET 8) desktop application that integrates with Microsoft Graph (Azure AD / Entra ID) to search and select users, render Siderise-branded security badges, and print to a Zebra ZC300 card printer.

The application is located in `SecurityBadgePrinter/` and is designed for the PC connected to the Zebra printer.

## Tech Stack
- **Framework**: .NET 8, WPF
- **Identity/Graph**: `Azure.Identity`, `Microsoft.Graph`
- **Rendering**: `SkiaSharp`, `QRCoder`
- **Printing**: `System.Drawing.Printing`
- **Config & Logging**: `Microsoft.Extensions.Configuration.*`, `Serilog`

## Repository Structure
- `SecurityBadgePrinter/`
  - `App.xaml`, `App.xaml.cs`: app bootstrap and configuration loading
  - `MainWindow.xaml`, `MainWindow.xaml.cs`: UI, search/filter logic, preview, print actions
  - `Models/AppConfig.cs`: strongly-typed configuration (`AzureAd`, `Printer`)
  - `Services/GraphServiceFactory.cs`: builds a `GraphServiceClient` using `InteractiveBrowserCredential`
  - `Services/BadgeRenderer.cs`: renders branded badge images (CR80)
  - `Services/PrinterService.cs`: prints rendered badges to Zebra ZC300
  - `appsettings.json`: tenant, client ID, and optional group whitelist
  - `SecurityBadgePrinter.csproj`: dependencies and resources
  - `installer/SecurityBadgePrinter.iss`: Inno Setup script (optional packaging)
  - `README.md`: short app-level readme (this root readme is the comprehensive overview)

## How It Works (High-level)
```mermaid
flowchart LR
  U[User] --> WPF[WPF App]
  WPF -->|Interactive login (MSAL via Azure.Identity)| AAD[Azure AD]
  WPF -->|Graph queries| GRAPH[Microsoft Graph]
  GRAPH -->|Users & Photos| WPF
  WPF -->|Render| R[BadgeRenderer (SkiaSharp)]
  R -->|PNG| P[PrinterService]
  P -->|Print Job| ZC[Zebra ZC300]
```

## Configuration
- File: `SecurityBadgePrinter/appsettings.json`
- Types: `SecurityBadgePrinter.Models.AppConfig`

```json
{
  "AzureAd": {
    "TenantId": "<tenant-guid>",
    "ClientId": "<app-client-id>",
    "AllowedGroupIds": [ "<optional-group-guid>", "<optional-group-guid>" ]
  },
  "Printer": {
    "Name": "Zebra - Security Badge"
  }
}
```

Notes:
- If `AzureAd.AllowedGroupIds` is populated, the app operates in "group whitelist mode" and pulls users from those security groups.
- If empty, the app queries the tenant directly using Microsoft Graph with server-side OData filters.
- The printer is selected by name. Ensure the Zebra ZC300 driver exposes the printer as configured here (default `Zebra - Security Badge`).

## Authentication & Token Cache
- `Services/GraphServiceFactory.cs` constructs a `GraphServiceClient` using `InteractiveBrowserCredential`.
- Redirect URI: `http://localhost` (Desktop/Mobile platform).
- A persistent auth record is stored at `%LOCALAPPDATA%\SecurityBadgePrinter\authRecord.json` to avoid sign-in prompts on subsequent runs. Delete this file to force re-auth.
- Scoped permissions requested: `User.Read`, `User.Read.All`, `Group.Read.All` (grant admin consent for org-wide usage).

## Search & Filtering
- Domain restriction: Only `@siderise.com` identities are considered (`MainWindow.xaml.cs` constant `AllowedDomain`).
- UI filter fields in `MainWindow.xaml`: `Name`, `Department`, `Job Title`, `Office`.
- Entry points:
  - `LoadFilterValuesAsync()`: populates initial filter options and loads the first user list.
  - `QueryAndBindUsersAsync(query)`: executes the current filter set, fetches users, and binds to the results list.

Two operating modes:

1) Non-group mode (no `AllowedGroupIds`)
- Uses Graph `Users.GetAsync()` with OData filters for domain, account status, exclusion keywords, and optional department/title/office startswith filters.

2) Group whitelist mode (`AllowedGroupIds` present)
- Aggregates members from listed security groups and applies client-side filtering.
- Office options are inferred from group `displayName` formatted as `"Office - Type"` (e.g., `Maesteg - Staff` → office `Maesteg`).
- `GetFilteredUsersAsync(selectedOffice, selectedDepartment, selectedTitle, token)` centralizes this logic.
- `UpdateFilterOptionsAsync()` updates the Department/Title options dynamically based on current Office and cross-filter selections.

## Badge Rendering
- File: `Services/BadgeRenderer.cs`
- Output size: CR80 at 300 DPI → `1011 x 638` pixels, landscape.
- Branding (per Siderise Brand Guidelines 2025):
  - Navy Primary `#002751` for name, department, QR code, borders
  - Blue Primary `#0090D0` for job title and accents
  - Arial typeface with bold weights for hierarchy
- Photo: Center-cropped with rounded-corner frame and a brand-blue border. If no photo, renders light-gray with initials.
- Text fitting: Adaptive single-line preference with intelligent wrapping and ellipsizing for long names/titles/departments.
- QR code: `BadgeRenderer.BuildQr(upn)` encodes as `{"C" + <UPN local-part> + "}"`, e.g., `john.smith@domain` → `{Cjohn.smith}`. Rendered with QRCoder and embedded into the badge.

## Printing
- File: `Services/PrinterService.cs`
- Renders PNG to a `PrintDocument` targeting the configured printer.
- Landscape, zero margins; attempts to force a CR80 physical size using a custom `PaperSize("CR80 (custom)", 337, 213)` (hundredths of an inch).
- Maps the 300-DPI design to device DPI at runtime and centers/clamps the image to avoid distortion.
- Uses `StandardPrintController` to suppress UI. Set the printer’s driver media to CR80 cards.

## UI & Branding
- File: `MainWindow.xaml`
- Implements Siderise’s 2025 brand guidelines across:
  - Color resources: Navy/Blue primaries plus light tints for backgrounds and states
  - Typography: Arial with bold headers and clear hierarchy
  - Buttons: modern primary/secondary styles with hover/press states
  - Panels: shadowed cards for Search/Results/Preview sections
- The banner displays logo (if present under `Assets/`) and app version.

## Build & Run
1. Install .NET 8 SDK and the Zebra ZC300 printer driver.
2. Azure Portal: Register an app (Single-tenant). Desktop/Mobile platform. Redirect URI `http://localhost`.
3. API permissions (Delegated): `User.Read`, `User.Read.All`, optionally `Group.Read.All`. Grant admin consent.
4. Configure `SecurityBadgePrinter/appsettings.json` with `TenantId`, `ClientId`, and optionally `AllowedGroupIds`.
5. Build and run `SecurityBadgePrinter`. Sign in when prompted.
6. Search for users, select one, preview the badge, and print.

## Packaging & Deployment

### Standard Installer (Inno Setup)
- `SecurityBadgePrinter/installer/SecurityBadgePrinter.iss` is an Inno Setup script to create a Windows installer.
- Adjust app/version metadata in the script before building.

### Intune Deployment
The application can be deployed via Microsoft Intune using the packaged `.intunewin` file:

1. **Package the application**: Run `scripts/package-intune.ps1` to create the Intune deployment package
2. **Output location**: `Output/` folder contains:
   - `SecurityBadgePrinter.intunewin` - The packaged application
   - `Detect-SecurityBadgePrinter.ps1` - Version-based detection script
3. **Upload to Intune**: 
   - Navigate to Intune > Apps > Windows apps > Add
   - Upload the `.intunewin` file
   - Configure install/uninstall commands (provided in the package)
   - Use the detection script for version-based deployment updates

The detection script checks for the installed version, making it easy to push updates by incrementing the version number in the project file.

### GitHub Actions CI/CD
Automated builds and tests run on every push via `.github/workflows/build.yml`:
- Builds the application for Release configuration
- Runs unit tests (when available)
- Creates build artifacts
- Validates code quality

## Known Issues / Next Steps
- Office/Department interactive filtering in group whitelist mode depends on group naming conventions (`"Office - Type"`). If groups don’t follow this format, Office derivation and cross-filtering will be incomplete.
- If some departments don’t appear after choosing an Office, verify the user memberships in the whitelisted groups and the group display names. The function to review is `UpdateFilterOptionsAsync()` and its calls to `GetFilteredUsersAsync()`.
- Future enhancements:
  - Duplex and overlays using Zebra Card SDK
  - Back-of-card layout templating
  - Admin panel for mapping offices to groups without relying on naming conventions

## Troubleshooting
- To force re-authentication: delete `%LOCALAPPDATA%\SecurityBadgePrinter\authRecord.json`.
- Ensure printer name matches `Printer.Name` in `appsettings.json`.
- If badges print slightly scaled/misaligned, check the printer driver’s card size and orientation; the app uses a custom CR80 size and device-DPI mapping.
- If logos don’t render: ensure `Assets/Logo.png` (or `Assets/Logo - White.png`) is present and included as a WPF Resource (see `SecurityBadgePrinter.csproj`).

## Dependencies (see `SecurityBadgePrinter.csproj`)
- Microsoft.Graph 5.x
- Azure.Identity 1.x
- Microsoft.Extensions.Configuration (core/json/binder) 8.x
- SkiaSharp 2.88.x (+ NativeAssets.Win32)
- QRCoder 1.4.x
- Serilog 3.x (+ Serilog.Sinks.File)
- System.Drawing.Common 8.x

## Versioning
- Assembly version is set in `SecurityBadgePrinter.csproj`. Update before packaging releases.

## Contact / Handover Notes
- The application implements Siderise brand visuals and end-to-end badge printing. Configuration and identity details are in `appsettings.json` and persisted auth records under `%LOCALAPPDATA%`.
- Start with the search and filtering flow in `MainWindow.xaml.cs`, then follow rendering in `BadgeRenderer.cs`, and printing in `PrinterService.cs`.
