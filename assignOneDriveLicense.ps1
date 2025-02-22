# Check if the Microsoft Graph module is installed; if not, install it
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Microsoft.Graph module not found. Installing..."
    Install-Module -Name Microsoft.Graph -Force -AllowClobber -Scope CurrentUser
}

# Import Microsoft Graph PowerShell Module
Import-Module Microsoft.Graph

# Authenticate (if not already authenticated)
Connect-MgGraph

# Function to assign OneDrive for Business Plan 1
function Assign-OneDriveLicense {
    param (
        [string]$userId
    )

    # Retrieve the SkuId for OneDrive for Business Plan 1 (WACONEDRIVESTANDARD)
    $skuId = (Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq "WACONEDRIVESTANDARD"}).SkuId

    # Check if the user already has the license assigned
    $userLicenses = Get-MgUserLicenseDetail -UserId $userId
    $hasLicense = $userLicenses | Where-Object { $_.SkuId -eq $skuId }

    if ($hasLicense) {
        Write-Host "User $userId already has OneDrive for Business Plan 1 assigned."
    } else {
        # Create an array of licenses to add
        $addLicenses = @([Microsoft.Graph.PowerShell.Models.MicrosoftGraphAssignedLicense]@{SkuId = $skuId})

        # Create an empty array for licenses to remove
        $removeLicenses = @()

        # Assign the license with both AddLicenses and RemoveLicenses
        Set-MgUserLicense -UserId $userId -AddLicenses $addLicenses -RemoveLicenses $removeLicenses

        Write-Host "License assigned successfully to user: $userId"
    }
}

# Take User Principal Name as input
$userUPN = Read-Host "Enter the User Principal Name (UPN)"
Assign-OneDriveLicense -userId $userUPN
