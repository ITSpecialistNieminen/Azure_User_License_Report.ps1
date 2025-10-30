 <#
  .SYNOPSIS
    Generates a Microsoft 365 license compliance report by combining Active Directory and Microsoft Graph data.

  .DESCRIPTION
    This script generates a Microsoft 365 license compliance report by retrieving user data 
    from Microsoft Graph and local Active Directory. It maps selected license SKUs (such as E3, Visio, Copilot, etc.)
    to users.

    Results are exported to a CSV file with detailed license presence per user. The script logs key actions and errors, handles log rotation, 
    and provides structured outputs for auditing and reporting.

   .REQUIREMENTS
        - PowerShell 5.1 or later
        - Microsoft.Graph PowerShell SDK module

   .PERMISSIONS
    Microsoft Graph API:
        - Directory.Read.All
    Azure Attribute Assignment Administrator role needed for security attributes (GA is not enough)

  .LINK
    License plan SKUs: 
    https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

  .OUTPUTS
    CSV file written to script directory with license and user details.
    Log file stored in /logs subfolder for tracking execution history and issues.

  .EXAMPLE
    .\Export-LicenseReport.ps1
    Runs the script and generates a license compliance CSV report.
    
    Authors:
        petri.nieminen91@gmail.com
#>

##### VARIABLES #####

# User who run script
$runas=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$ScriptDirectory = split-path -parent $MyInvocation.MyCommand.Definition
$LogDirectory = Join-Path $ScriptDirectory "logs"

# Check if the main folder exists
if (-Not (Test-Path -Path $ScriptDirectory)) {
    # Create the main folder
    New-Item -ItemType Directory -Path $ScriptDirectory
    Write-Host "Folder created: $ScriptDirectory"
} else {
    Write-Host "Folder already exists: $ScriptDirectory"
}

# Check if the 'logs' subfolder exists
if (-Not (Test-Path -Path $LogDirectory)) {
    # Create the 'logs' subfolder
    New-Item -ItemType Directory -Path $LogDirectory
    Write-Host "Subfolder 'logs' created: $LogDirectory"
} else {
    Write-Host "Subfolder 'logs' already exists: $LogDirectory"
}

$date = (Get-Date -Format 'yyyyMMdd')

# Log file custom name
$LogFileName = "ReportAction.log"
# Log filename based on date and label
$LogFileNameFull = $date + "_" + $LogFileName
# Log file location
$loglocation = "$ScriptDirectory\logs"
$Logfile = (Join-Path $loglocation $LogFileNameFull)
# Log retention policy (days)
$LogRetention = 180
$logmaxage = (Get-Date).AddDays(-$LogRetention)

# ResultDir
$File_cloud = "$ScriptDirectory\$($date)_Report.csv"

# Build the list of properties to retrieve
$propertiesToGet = "DisplayName,UserPrincipalName,AssignedLicenses,OnPremisesSyncEnabled,Id,userType,CompanyName"

# Initialize empty CSV file with header to ensure correct column order
$HeaderObjectCloud = [PSCustomObject]@{
    DisplayName          = ""
    UserPrincipalName    = ""
    UserID               = ""
    UserType             = ""
    CompanyName          = ""
}

##### FUNCTIONS #####

# Logging function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$False)]
        [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
        [String]
        $Level = "INFO",
        [Parameter(Mandatory=$True)]
        [string]
        $Message,
        [Parameter(Mandatory=$False)]
        [string]
        $Logfile = (Join-Path $loglocation $LogFileNameFull)
    )
 
    $timestamp = (Get-Date).toString("dd.MM.yyyy HH:mm:ss")
    $line = "$timestamp $level $message"
    if($logfile) {
        Add-Content $logfile -Value $line
    } else {
        Write-Output $line
    }
}

# Helper function
function Test-LicensePresence {
    param(
        [string]$licenseKey,
        [string]$upn
    )
    return $licenseHolders[$licenseKey] -contains $upn
}

##### SCRIPT #####

Write-Log -Level INFO "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<START>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"

# Log rotate
try {
    $oldLogs = Get-ChildItem "$loglocation\*.log" | Where-Object { $_.LastWriteTime -lt $logmaxage }

    if ($oldLogs.Count -gt 0) {
        $oldLogs | Remove-Item -Force
        Write-Log -Level INFO -Message "Deleted $($oldLogs.Count) log file(s) older than $LogRetention days."
    } else {
        Write-Log -Level INFO -Message "No log files older than $LogRetention days to clean up."
    }
}
catch {
    Write-Log -Level WARN -Message "Log cleanup failed: $($_.Exception.Message)"
}

Write-Log -Level INFO "Script run as $runas."

try {
    # Check if the Microsoft.Graph module is already installed
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
        Write-Log -Level WARN "Microsoft.Graph module is not installed."
    }

    # Attempt to import the module
    Import-Module Microsoft.Graph.Users -ErrorAction Stop
    Write-Log -Level INFO "Microsoft.Graph.Subscription and Microsoft.Graph.Users module successfully imported."
}
catch {
    Write-Log -Level FATAL "Failed to import Microsoft.Graph module. Error:`n$($_.Exception.Message)"
    exit 1
}

# Connect to Microsoft Graph
Write-Log -Level INFO -Message "Connecting to Microsoft Graph..."
try {
    Connect-MgGraph -ErrorAction Stop
    Write-Log -Level INFO -Message "Connection to Microsoft Graph established."
} catch {
    Write-Log -Level FATAL -Message "Failed to connect to Microsoft Graph. Error: $($_.Exception.Message)"
    exit
}

# Retrieve all licensed users from Microsoft Graph
Write-Log -Level INFO -Message "Retrieving Graph users with selected properties..."
try {
    $GraphUsers = Get-MgUser -All -Filter "accountEnabled eq true" -Property $propertiesToGet
    Write-Log -Level INFO "Retrieved licensed users from Microsoft Graph"
}
catch {
    Write-Log -Level ERROR -Message "Failed to retrieve Microsoft Graph users. Error:`n$($_.Exception.Message)"
    exit
}


# Build a map of SkuId → SkuPartNumber
$skuMap = @{}
try {
    $subscribedSkus = Get-MgSubscribedSku
    foreach ($sku in $subscribedSkus) {
        if ($sku.SkuId -and $sku.SkuPartNumber) {
            $skuMap[$sku.SkuId] = $sku.SkuPartNumber
        }
    }
    Write-Log -Level INFO "Mapped SKU IDs to SKU part numbers"
}
catch {
    Write-Log -Level ERROR -Message "Failed to retrieve or map subscribed SKUs. Error:`n$($_.Exception.Message)"
    exit
}

# Build a hashtable of UPNs per license SKU part number
$licenseSkuMap = @{
    Visio                   = "VISIOCLIENT"                         # Visio Plan 2
    Copilot                 = "MICROSOFT_365_COPILOT"               # Microsoft 365 Copilot
    AudioConference         = "MCOMEETADV"                          # Microsoft 365 Audio Conferencing
    TeamsEEA                = "MICROSOFT_TEAMS_EEA_NEW"             # Microsoft Teams EEA
    E3                      = "SPE_E3"                               # Microsoft 365 E3
    ATPEnterprise           = "ATP_ENTERPRISE"                      # Microsoft Defender for Office 365 (Plan 1)
    E1                      = "STANDARDPACK"                        # Office 365 E1
    O365NoTeams             = "O365_W/O TEAMS BUNDLE_M3"            # Microsoft 365 E3 EEA (no Teams)
    CopilotStudioTrial      = "MICROSOFT_COPILOT_STUDIO_TRIAL"      # Microsoft Copilot Studio Viral Trial
    PowerAutomateFree       = "FLOW_FREE"                           # Microsoft Power Automate Free
    TeamsRoomsPro           = "MTR_PRO"                             # Microsoft Teams Rooms Pro
    FabricFree              = "FABRIC_FREE"                         # Microsoft Fabric (Free)
    PowerAppsDeveloper      = "POWERAPPS_DEV"                       # Microsoft Power Apps for Developer
    PowerAppsPlan2Trial     = "POWERAPPS_VIRAL"                     # Microsoft Power Apps Plan 2 Trial
    AppConnect              = "APP_CONNECT"                         # App Connect 
    StreamTrial             = "STREAM_O365_E3_VIRAL"                # Microsoft Stream Trial
    TeamsPremiumDept        = "TEAMS_PREMIUM_FOR_DEPARTMENTS"       # Teams Premium (for Departments)
}
#Full list: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

# Build a list of UPNs per SKU
$licenseHolders = @{}
foreach ($key in $licenseSkuMap.Keys) {
    $licenseHolders[$key] = @()
}

try {
    foreach ($user in $GraphUsers) {
        foreach ($assigned in $user.AssignedLicenses) {
            $skuId = $assigned.SkuId
            if ($skuMap.ContainsKey($skuId)) {
                $skuPart = $skuMap[$skuId]
                foreach ($key in $licenseSkuMap.Keys) {
                    if ($skuPart -eq $licenseSkuMap[$key]) {
                        $licenseHolders[$key] += $user.UserPrincipalName
                    }
                }
            }
        }
    }
    Write-Log -Level INFO "Built license mapping for users"
}
catch {
    Write-Log -Level ERROR -Message "Failed to map licenses to users. Error:`n$($_.Exception.Message)"
    exit
}

# Prepare result file with headers
try {
    # Add license SKU keys as properties to header object BEFORE exporting
    foreach ($key in $licenseSkuMap.Keys) {
        $HeaderObjectCloud | Add-Member -NotePropertyName $key -NotePropertyValue ""
    }

    # Now write header to CSV file
    $HeaderObjectCloud | Export-Csv -Path $File_cloud -NoTypeInformation -Encoding UTF8
    }
catch {
    Write-Log -Level ERROR -Message "Failed to write header to file. Error:`n$($_.Exception.Message)"
    exit
}

# Loop through cloud users and append license information
try {
    $cloudUsers = $GraphUsers #| Where-Object { $_.OnPremisesSyncEnabled -ne $true }

    foreach ($user in $cloudUsers) {

        # Build hashtable for the output object
        $outputProps = @{
            DisplayName          = $user.DisplayName
            UserPrincipalName    = $user.UserPrincipalName
            UserID               = $user.Id
            UserType             = $user.UserType
            CompanyName          = $user.CompanyName
        }

        # Add license presence info to output properties
        foreach ($key in $licenseSkuMap.Keys) {
            $outputProps[$key] = if (Test-LicensePresence -licenseKey $key -upn $user.UserPrincipalName) { $true } else { $false }
        }

        # Create PSCustomObject with all properties at once
        $outputObject = [PSCustomObject]$outputProps

        # Append the object to CSV
        $outputObject | Export-Csv -Path $File_cloud -Append -NoTypeInformation -Encoding UTF8
    }

    Write-Log -Level INFO "Processed cloud-only users"
}
catch {
    Write-Log -Level ERROR -Message "Error while processing cloud-only users. Error:`n$($_.Exception.Message)"
    exit
}
Write-Log -Level INFO "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>" 
