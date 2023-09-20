# Script:	Azure_license_monitoring.ps1
# Purpose:	Monitor assigned Azure licenses and email if license count drop below threshold
# Author:	Petri Nieminen 
# Github:	https://github.com/ITSpecialistNieminen/Azure_License_monitoring.ps1.git


##### VARIABLES #####

# User who run script
$runas=[System.Security.Principal.WindowsIdentity]::GetCurrent().Name

#$ScriptDirectory = split-path -parent $MyInvocation.MyCommand.Definition
$ScriptDirectory = "C:\Monitoring\Azure"

# Log file custom name
$LogFileName = "O365_licenses.log"
# Log filename based on date and label
$LogFileNameFull = (Get-Date -Format yyyyMMdd) + "_" + $LogFileName
# Log file location
$loglocation = "$ScriptDirectory\logs"
$Logfile = (Join-Path $loglocation $LogFileNameFull)
# Log retention policy (90 days)
$LogRetention = 90
$logmaxage = (Get-Date).AddDays(-$LogRetention)

$TenantLicenses = @()  # Initialize an empty array to store license information

# Specify the license SKU you want to check
$licenseIds = "EMS", #Enterprise Mobility + Security E3
"EMSPREMIUM", #Enterprise Mobility + Security E5
"EXCHANGESTANDARD", #Exchange Online (Plan 1)
"ENTERPRISEPACK", #Office 365 E3
"STANDARDPACK", #Office 365 E1
"SPE_E3", #Microsoft 365 E3
"IDENTITY_THREAT_PROTECTION" #Microsoft 365 E5 Security

#Full list: https://learn.microsoft.com/en-us/azure/active-directory/enterprise-users/licensing-service-plan-reference

# Email settings
$from = ''
$to = ''
$bcc = ''
$SmtpServer = ''

# Azure login creds
$TenantUname = ""
$TenantPass = Get-Content -Path "$ScriptDirectory\Passwords\Azure_Monitoring_pwd.txt" | ConvertTo-SecureString
$TenantCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $TenantUname, $TenantPass

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

# License function
Function Get-Licenses {
    param (
        [Parameter(Mandatory = $true)]
        [string]$LicenseId
    )

    # SKU to product name mapping table
    $productNames = @{
        "EMS" = "Enterprise Mobility + Security E3"
        "EMSPREMIUM" = "Enterprise Mobility + Security E5"
        "EXCHANGESTANDARD" = "Exchange Online (Plan 1)" 
        "ENTERPRISEPACK" = "Office 365 E3"
        "STANDARDPACK" = "Office 365 E1"
        "SPE_E3" = "Microsoft 365 E3"
        "IDENTITY_THREAT_PROTECTION" = "Microsoft 365 E5 Security"
        # Add more mappings if needed
    }

    # Retrieve license information for the tenant
    $TenantLicenses = Get-MsolAccountSku | Where-Object {$_.SkuPartNumber -eq $LicenseId} |
        Select-Object -Property SkuPartNumber, ActiveUnits, ConsumedUnits |
        ForEach-Object {
            $unitsLeft = $_.ActiveUnits - $_.ConsumedUnits
            [PSCustomObject]@{
                SkuPartNumber = $_.SkuPartNumber
                ProductName = $productNames[$_.SkuPartNumber]
                ActiveUnits = $_.ActiveUnits
                ConsumedUnits = $_.ConsumedUnits
                UnitsLeft = $unitsLeft
            }
        }

    # Return the license information
    return $TenantLicenses
}


##### SCRIPT #####

# Log rotate
try {
    Get-ChildItem $loglocation\*$LogFileName | Where-Object {$_.LastWriteTime -lt $logmaxage} -ErrorAction Stop | Remove-Item -ErrorAction Continue
}
catch {
    Write-Log -Level WARN -Message "Log rotate failed. Error:`r`n$_.Exception.Message"
}

Write-Log -Level INFO "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<START>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
Write-Log -Level INFO "Script run as $runas."

#Check if MSOnline PowerShell module has been installed
Try {
    Write-Log -Level INFO "Checking if MSOnline PowerShell Module is Installed..."
    $MSOnlineModule = Get-Module -ListAvailable "MSOnline"
 
    #Check if MSOnline  Module is installed
    If(!$MSOnlineModule)
    {
        Write-Log -Level INFO "MSOnline Module not found" 
 
        #Check if script is executed under elevated user permissions - Run as Administrator
        If (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))
        {  
            Write-Log -Level INFO "Please Run this script in elevated mode (Run as Administrator)! "
            Exit
        }
 
        Write-Log -Level INFO "Installing MSOnline PowerShell Module..."
        Install-Module MSOnline -Force -Confirm:$False
        Write-Log -Level INFO "MSOnline Module installed!"
    }
    Else
    {
        Write-Log -Level INFO "MSOnline Module found"
        #sharepoint online powershell module import
        Write-Log -Level INFO "Importing MSOnline PowerShell Module..."
        Import-Module MSOnline
        Write-Log -Level INFO "MSOnline Module imported successfully"
    }
}
Catch{
    Write-Log -Level ERROR "MSOnline Module Error:`r`n$($_.Exception.Message)"
}

# Set the security protocol to TLS 1.2
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Connect to MSOnline service
try {
    Connect-MsolService  -Credential $TenantCredentials -ErrorAction Stop
    Write-Log -Level INFO "Connection to MsolService established"
}
catch {
    Write-Log -Level ERROR -Message "Failed to connect to MSOnline service. Error:`r`n$($_.Exception.Message)"
    exit
}

foreach ($licenseId in $licenseIds) {
    try 
        {
        $licenses = Get-Licenses -LicenseId $licenseId # Function call
        $licensesUnitsLeft = $licenses.UnitsLeft
        $licensesProductName = $licenses.ProductName
        }
    catch 
        {
        Write-Log -Level ERROR -Message "Failed to retrieve license information for SKU '$licenseId'. Error:`r`n$($_.Exception.Message)"
        }
        if ($licensesUnitsLeft -lt 3) {
            try
                {
                Write-Log -Level WARN "There are only $licensesUnitsLeft units left on the license $licensesProductName."
             
                $date = Get-Date
                $body = "Product $licensesProductName have only $licensesUnitsLeft licenses left."
                $Subject = "Azure licenses monitoring: License count below threshold!."       
                Send-MailMessage -Encoding 'UTF8' -From $from -To $to -Subject $Subject -Body $body -SmtpServer $SmtpServer -Bcc $bcc
                Write-Log -Level INFO -Message "Email notification sent to: $to"
                }
            catch
                {
                Write-Log -Level ERROR -Message "Failed to send email. Error:`r`n$($_.Exception.Message)"
                }
        }
        $TenantLicenses += $licenses  # Append license information to the array (include all unit information)
}
 
Write-Log -Level INFO "Current status of licenses:"
# Iterate over each license and write to the log file
$TenantLicenses | ForEach-Object {
    $line = "SkuPartNumber: $($_.SkuPartNumber), ProductName: $($_.ProductName), ActiveUnits: $($_.ActiveUnits), ConsumedUnits: $($_.ConsumedUnits), UnitsLeft: $($_.UnitsLeft)"
    Write-Log -Level INFO -Message $line
}
 
# Disconnect to MsolService
try {
    [Microsoft.Online.Administration.Automation.ConnectMsolService]::ClearUserSessionState()
    Write-Log -Level INFO "Connection to MsolService disconnected"
}
catch {
    Write-Log -Level ERROR -Message "Failed to disconnect to MsolService. Error:`r`n$($_.Exception.Message)"
    exit
} 
 
Write-Log -Level INFO "<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<END>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
