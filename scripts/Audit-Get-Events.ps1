# Export script for Microsoft Copilot Interactions events from Purview Audit Log
# Special thanks goes to https://github.com/12Knocksinna/Office365itpros team

# Check if Exchange Online module is already installed
$module = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'ExchangeOnlineManagement' }

if ($module -eq $null) {
    try {
        Write-Host "Installing module..."
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser
    } 
    catch {
        Write-Host "Failed to install module: $_"
        exit
    }
}

Import-Module ExchangeOnlineManagement

# Connect to Exchange Online
try {
    Connect-ExchangeOnline
    Write-Host "Connected to Exchange Online."
} catch {
    Write-Host "Failed to connect to Exchange Online: $_"
}

# CSV Folder path  
$csvpath = "C:\M365CopilotReport\Copilot_Events.csv"

[array]$Records = Search-UnifiedAuditLog -StartDate (Get-Date).Adddays(-120) -EndDate (Get-Date).AddDays(1) -Formatted -ResultSize 5000 -SessionCommand ReturnLargeSet -Operations CopilotInteraction
If (!($Records)) {
    Write-Host "No Copilot audit records found - exiting"
    Break
} Else {
    # Remove any duplicate records and make sure that everything is sorted in date order
    $Records = $Records | Sort-Object Identity -Unique 
    $Records = $Records | Sort-Object {$_.CreationDate -as [datetime]}
    Write-Host ("{0} Copilot audit records found. Now analyzing the content" -f $Records.count)
}

$Report = [System.Collections.Generic.List[Object]]::new()
ForEach ($Rec in $Records) {
    $AuditData = $Rec.AuditData | ConvertFrom-Json
    $CopilotApp = 'Copilot for M365'; $Context = $null; $CopilotLocation = $null
    
    Switch ($Auditdata.copiloteventdata.contexts.type) {
        "xlsx" {
            $CopilotApp = "Excel"
        }
        "docx" {
            $CopilotApp = "Word"
        }
        "pptx" {
            $CopilotApp = "PowerPoint"
        }
        "TeamsMeeting" {
            $CopilotApp = "Teams"
            $CopilotLocation = "Teams meeting"
        }
        "whiteboard" {
            $CopilotApp = "Whiteboard"    
        }
        "loop"{
            $CopilotApp = "Loop"
        }
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotApp = "Teams"
    } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat") {
        $CopilotApp = "Copilot for M365 Chat"
    }

    If ($Auditdata.copiloteventdata.contexts.id) {
        $Context = $Auditdata.copiloteventdata.contexts.id
    } ElseIf ($Auditdata.copiloteventdata.threadid) {
        $Context = $Auditdata.copiloteventdata.threadid
        # $CopilotApp = "Teams"
    }

    If ($Auditdata.copiloteventdata.contexts.id -like "*/sites/*") {
        $CopilotLocation = "SharePoint Online"
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
        $CopilotLocation = "Teams"
        If ($Auditdata.copiloteventdata.contexts.id -like "*ctx=channel*") {
            $CopilotLocation = "Teams Channel"
        } Else {
            $CopilotLocation = "Teams Chat"
        }
    } ElseIf ($Auditdata.copiloteventdata.contexts.id -like "*/personal/*") {
        $CopilotLocation = "OneDrive for Business"
    } 
    # Make sure that we report the resources used by Copilot
    $AccessedResources = $AuditData.copiloteventdata.accessedResources.name -join ", "
    $AccessedResourceLocations = $AuditData.copiloteventdata.accessedResources.id -join ", "

    $ReportLine = [PSCustomObject][Ordered]@{
        TimeStamp                       = (Get-Date $Rec.CreationDate -format "dd-MMM-yyyy HH:mm:ss")
        User                            = $Rec.UserIds
        App                             = $CopilotApp
        Location                        = $CopilotLocation
        'App context'                   = $Context   
        'Accessed Resources'            = $AccessedResources
        'Accessed Resource Locations'   = $AccessedResourceLocations
    }
    $Report.Add($ReportLine)
}

$Report | ConvertTo-Csv -NoTypeInformation | Out-File $csvpath
