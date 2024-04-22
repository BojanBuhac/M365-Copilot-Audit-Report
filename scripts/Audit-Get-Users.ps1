# Audit log export script for Copilot Interactions Events

# Check if AzureAD Online module is already installed
$moduleAzureADInstalled = Get-Module -ListAvailable | Where-Object { $_.Name -eq 'AzureAD' }

if ($moduleAzureADInstalled -eq $null) {
    # AzureAD module is not installed, attempt to install it
    try {
        Write-Host "Installing AzureAD module..."
        Install-Module -Name AzureAD -Force -AllowClobber -Scope CurrentUser
    } catch {
        Write-Host "Failed to install AzureAD module: $_"
        exit
    }
}

# Import the AzureAD module
Import-Module AzureAD

# Connect to AzureAD
try {
    Connect-AzureAD
    Write-Host "Connected to AzureAD."
} catch {
    Write-Host "Failed to connect to AzureAD: $_"
}

# Define the folder path
$folderPath = "C:\M365CopilotReport\"

# Path to the output CSV file
$outputCsv = $folderPath + "Copilot_Users.csv"

Get-AzureADUser -All $true | Select DisplayName, UserPrincipalName, jobTitle, City, Country, UsageLocation -ExpandProperty AssignedLicenses | Where-Object {$_.SkuID -eq '639dec6b-bb19-468b-871c-c5c441c4b0cb'} | Export-csv $outputCsv