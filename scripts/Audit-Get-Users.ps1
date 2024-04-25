# Check if MG Online module is already installed
$module = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'Microsoft.Graph' }

if ($module -eq $null) {
    try {
        Write-Host "Installing module..."
        Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
    } 
    catch {
        Write-Host "Failed to install module: $_"
        exit
    }
}
# Connect to Microsoft Graph
try {
    Connect-mggraph -Scopes "User.Read.All" -NoWelcome
    Write-Host "Connected to Microsoft Graph."
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_"
}

# CSV File path  
$csvUserspath = "C:\M365CopilotReport\Copilot_Users.csv"
Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $('639dec6b-bb19-468b-871c-c5c441c4b0cb'))" -ConsistencyLevel eventual -CountVariable CopilotLicensedUserCount -All -Property DisplayName, UserPrincipalName, jobTitle, City, Country, UsageLocation | Select-Object DisplayName, UserPrincipalName, jobTitle, City, Country, UsageLocation | Export-csv $csvUserspath
