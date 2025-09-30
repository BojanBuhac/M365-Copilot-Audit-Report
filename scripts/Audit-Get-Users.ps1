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
    Connect-MgGraph -Scopes "User.Read.All","User.ReadBasic.All","Directory.Read.All" -NoWelcome
    Write-Host "Connected to Microsoft Graph."
} catch {
    Write-Host "Failed to connect to Microsoft Graph: $_"
}

# CSV File path  
$csvUserspath = "C:\M365CopilotReport\Copilot_Users.csv"

# Replace with actual Copilot SKU ID(s) from your tenant
$copilotSkuIds = "639dec6b-bb19-468b-871c-c5c441c4b0cb"

# Get users with job titles
$users = Get-MgUser -Filter "JobTitle ne null" -ConsistencyLevel eventual -CountVariable CopilotLicensedUserCount -All -Property Id, DisplayName, UserPrincipalName, JobTitle, Department, City, Country, UsageLocation, AssignedLicenses

# Build enriched objects with manager info and license check
$results = foreach ($user in $users) {

    $managerName = ""
    $managerUPN = ""
    $hasCopilot = $false

    # Get manager info
    try {
        $manager = Get-MgUserManager -UserId $user.Id -ErrorAction Stop
        $managerData = Get-MgUser -UserId $manager.Id -ErrorAction Stop
        $managerName = $managerData.DisplayName
        $managerUPN = $managerData.UserPrincipalName
    } catch {
        $managerName = ""
        $managerUPN = ""
    }

     # Check for Copilot license
    if ($user.AssignedLicenses) {
        foreach ($lic in $user.AssignedLicenses) {
            $skuIdString = "$($lic.SkuId)"
            if ($copilotSkuIds -contains $skuIdString) {
                $hasCopilot = $true
                break
            }
        }
    }

    [PSCustomObject]@{
        EntraID           = $user.Id
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        JobTitle          = $user.JobTitle
        Department        = $user.Department
        City              = $user.City
        Country           = $user.Country
        UsageLocation     = $user.UsageLocation
        ManagerName       = $managerName
        ManagerUPN        = $managerUPN
        HasCopilotLicense = $hasCopilot
    }
}

# Export to CSV
$results | Export-Csv $csvUserspath -NoTypeInformation -Encoding UTF8
Write-Host "Report exported to $csvUserspath"

