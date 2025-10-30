# Microsoft 365 License Compliance Report
MSGraph_M365_licenses.ps1

## Synopsis

This script generates a Microsoft 365 license compliance report by retrieving user and license data from Microsoft Graph and mapping selected license SKUs (e.g., E3, Visio, Copilot) to users.  
The report is exported to CSV, and execution details are logged with automatic log rotation.

## Description

- **Retrieve license data via Microsoft Graph**: Collects license assignments for all enabled users in the tenant.
- **License-to-user mapping**: Matches each configured SKU with the list of users holding that license.
- **CSV output**: Creates a structured report showing which users have which licenses.
- **Logging with rotation**: All actions and errors are logged, with automatic log cleanup after the configured retention period.
- **Customizable SKU list**: Update `$licenseSkuMap` in the script to include additional licenses to track.

## Requirements

- PowerShell 5.1 or later
- Microsoft.Graph PowerShell SDK (at least Microsoft.Graph.Users module)
- Microsoft Graph API permissions:
  - Directory.Read.All
- Azure Attribute Assignment Administrator role needed for security attributes (GA is not enough)

## Usage

1. Install Microsoft Graph PowerShell SDK if not already installed:
    ```powershell
    Install-Module Microsoft.Graph -Scope CurrentUser
    ```

2. Folder structure (created automatically if missing):
    ```
    C:\Machine\Scripts\Licensereport\
    ├─ Export-LicenseReport.ps1
    ├─ logs\
    │  ├─ [date]_ReportAction.log
    ├─ [date]_Report.csv
    ```

3. Run the script:
    ```powershell
    .\Export-LicenseReport.ps1
    ```

4. Authenticate to Microsoft Graph when prompted. Ensure the account has permissions to read user and license data.

## Output

- **CSV File** – named `[yyyyMMdd]_Report.csv`, includes:
  - DisplayName  
  - UserPrincipalName  
  - UserID  
  - UserType  
  - CompanyName  
  - Columns for each monitored license with `True`/`False` values

- **Log File** – stored in `/logs`, includes script activity, errors, and cleanup actions.

## Variables

- `$runas` – user running the script
- `$ScriptDirectory` – script location
- `$LogRetention` – days to retain log files
- `$licenseSkuMap` – SKUs and friendly names to track
- `$File_cloud` – output CSV file path
- `$propertiesToGet` – Microsoft Graph user properties retrieved

## Functions

- `Write-Log` – logs messages with timestamp and severity
- `Test-LicensePresence` – checks if a given UPN has a specified license

## Example

```powershell
.\Export-LicenseReport.ps1
```
## Notes
Version: 1.0
Author: petri.nieminen91@gmail.com
