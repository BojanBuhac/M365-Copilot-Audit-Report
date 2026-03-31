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

# Get tenant domain for connection
$tenantDomain = Read-Host "Enter your tenant domain (e.g., contoso.onmicrosoft.com or contoso.com)"

# Connect to Exchange Online
try {
    Connect-ExchangeOnline -UserPrincipalName "admin@$tenantDomain" -ShowProgress $true
    Write-Host "Connected to Exchange Online for tenant: $tenantDomain"
} catch {
    Write-Host "Failed to connect to Exchange Online: $_"
    exit
}

#Modify the values for the following variables to configure the audit log search.
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$reportDir = Join-Path $scriptPath "M365CopilotReport"

# Create output directory if it doesn't exist
if (!(Test-Path $reportDir)) {
    New-Item -ItemType Directory -Path $reportDir -Force
    Write-Host "Created output directory: $reportDir"
}

$logFile = Join-Path $reportDir "AuditScriptLog.txt"
$outputFile = Join-Path $reportDir "Copilot_Events.csv"
If(Test-Path $outputFile -PathType Leaf)
    {
        $lastEvent = Get-Content $outputFile -ErrorAction SilentlyContinue | Select-Object -Last 1
        $start = [DateTime]$lastEvent.Substring(1,20)
    } 
    else
    {
        [DateTime]$start = [DateTime]::UtcNow.AddDays(-365)
    }
[DateTime]$end = [DateTime]::UtcNow
$record = "CopilotInteraction"
$resultSize = 5000
$intervalMinutes = 1440

#Start script
[DateTime]$currentStart = $start
[DateTime]$currentEnd = $end

Function Write-LogFile ([String]$Message)
{
    $final = [DateTime]::Now.ToUniversalTime().ToString("s") + ":" + $Message
    $final | Out-File $logFile -Append
}

Write-LogFile "BEGIN: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$Recordsize."
Write-Host "Retrieving audit records for the date range between $($start) and $($end), RecordType=$record, ResultsSize=$Recordsize"

$totalCount = 0
while ($true)
{
    $currentEnd = $currentStart.AddMinutes($intervalMinutes)
    if ($currentEnd -gt $end)
    {
        $currentEnd = $end
    }

    if ($currentStart -eq $currentEnd)
    {
        break
    }

    $sessionID = [Guid]::NewGuid().ToString() + "_" +  "ExtractLogs" + (Get-Date).ToString("yyyyMMddHHmmssfff")
    Write-LogFile "INFO: Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    Write-Host "Retrieving audit records for activities performed between $($currentStart) and $($currentEnd)"
    $currentCount = 0

    $sw = [Diagnostics.StopWatch]::StartNew()

    do
    {
        $Records = Search-UnifiedAuditLog -StartDate $currentStart -EndDate $currentEnd -RecordType $record -SessionId $sessionID -SessionCommand ReturnLargeSet -ResultSize $resultSize

        if (($Records | Measure-Object).Count -ne 0)
        {
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
                    "StreamVideo" {
                        $CopilotApp = "Stream"
                        $CopilotLocation = "Stream video player"
                    }
                }
            
                If ($Auditdata.copiloteventdata.contexts.id -like "*https://teams.microsoft.com/*") {
                    $CopilotApp = "Teams"
                } ElseIf ($AuditData.CopiloteventData.AppHost -eq "bizchat") {
                    $CopilotApp = "Copilot for M365 Chat"
                }  ElseIf ($AuditData.CopiloteventData.AppHost -eq "Outlook") {
                    $CopilotApp = "Outlook"
                }   ElseIf ($AuditData.CopiloteventData.AppHost -eq "Copilot Studio") {
                    $CopilotApp = "Copilot Studio Agent"
                }
            
                If ($Auditdata.copiloteventdata.contexts.id) {
                    $Context = $Auditdata.copiloteventdata.contexts.id
                } ElseIf ($Auditdata.copiloteventdata.threadid) {
                    $Context = $Auditdata.copiloteventdata.threadid
                    # $CopilotApp = "Teams"
                }

                $AgentName = ""
                if ($CopilotApp -eq "Copilot Studio Agent" -and $AuditData.AppIdentity -match '.*_(.+?)$') {
                    $AgentName = $Matches[1]
                } elseif ($CopilotApp -eq "Copilot Studio Agent" -and $AuditData.AppIdentity -match '.*-(.+?)$') {
                    $AgentName = $Matches[1]
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
                 # Make sure that we report the resources used by Copilot and the action (like read) used to access the resource
                [array]$AccessedResources = $AuditData.copiloteventdata.accessedResources.name | Sort-Object -Unique
                [string]$AccessedResources = $AccessedResources -join ", "
                [array]$AccessedResourceLocations = $AuditData.copiloteventdata.accessedResources.id | Sort-Object -Unique
                [string]$AccessedResourceLocations = $AccessedResourceLocations -join ", "
                [array]$AccessedResourceActions = $AuditData.copiloteventdata.accessedResources.action | Sort-Object -Unique
                [string]$AccessedResourceActions = $AccessedResourceActions -join ", "
    
                $ReportLine = [PSCustomObject][Ordered]@{
                    TimeStamp                       = (Get-Date $Rec.CreationDate -format "dd-MMM-yyyy HH:mm:ss")
                    User                            = $Rec.UserIds
                    App                             = $CopilotApp
                    Location                        = $CopilotLocation
                    'App context'                   = $Context   
                    'Accessed Resources'            = $AccessedResources
                    'Accessed Resource Locations'   = $AccessedResourceLocations
                    Action                          = $AccessedResourceActions
                    AgentName                       = $AgentName
                }
                $Report.Add($ReportLine)
            }

            $Report | export-csv -Path $outputFile -Append -NoTypeInformation

            $currentTotal = $Records[0].ResultCount
            $totalCount += $Records.Count
            $currentCount += $Records.Count
            Write-LogFile "INFO: Retrieved $($currentCount) audit records out of the total $($currentTotal)"

            if ($currentTotal -eq $Records[$Records.Count - 1].ResultIndex)
            {
                $message = "INFO: Successfully retrieved $($currentTotal) audit records for the current time range. Moving on!"
                Write-LogFile $message
                Write-Host "Successfully retrieved $($currentTotal) audit records for the current time range. Moving on to the next interval." -foregroundColor Yellow
                ""
                break
            }
        }
    }
    while (($Records | Measure-Object).Count -ne 0)

    $currentStart = $currentEnd
}

Write-LogFile "END: Retrieving audit records between $($start) and $($end), RecordType=$record, PageSize=$Recordsize, total count: $totalCount."
Write-Host "Script complete! Finished retrieving audit records for the date range between $($start) and $($end). Total count: $totalCount" -foregroundColor Green


