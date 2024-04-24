# Export script for Microsoft Copilot Interactions events from Purview Audit Log

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

# Folder path, CSV and timestamp for files
$path = "C:\M365CopilotReport\"
$csv = $path + "Copilot_Events.csv"
$log = $path + "Copilot_TimeStamp.txt"
$pathExists = Test-Path $csv

# Get the last processed timestamp from the Copilot_TimeStamp.txt and set the date range
$lastTimeStamp = Get-Content $log -ErrorAction SilentlyContinue | Select-Object -Last 1
$start = [DateTime]$lastTimeStamp
$end = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

#If path doesn't exist create it with headers
if (! $pathExists) {
        $date = $start
        $results = Search-UnifiedAuditLog -StartDate $date -EndDate $date.AddDays(1) -RecordType CopilotInteraction -ResultSize 5000
        if ($results) {
            $results | ConvertTo-Csv -NoTypeInformation | Select-Object -First 1 | Out-File $csv
        }
}

# Collect data for each event from timestamp to today without headers
    for ($date = $start; $date -le $end; $date = $date.AddDays(1)) {
        $results = Search-UnifiedAuditLog -StartDate $date -EndDate $date.AddDays(1) -RecordType CopilotInteraction -ResultSize 5000
            if ($results) {
                $results | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File $csv -Append
            }
        }

# Write the time of completion to Copilot_TimeStamp.txt
$today = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
Add-Content -Path $log -Value $today
Write-Host "The Copilot_TimeStamp.txt file was last updated at: $today"
