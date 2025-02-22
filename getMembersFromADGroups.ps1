# Define the group names and output file path
$groupNames = @("BH_Strawhats_RW", "Bh_OnePiece_RW")
$outputFile = "C:\Users\komminenis\Documents\temp\RW-Members.xlsx"

# Initialize an empty array to store results
$allGroupMembers = @()

# Loop through each group and fetch members
foreach ($groupName in $groupNames) {
    $groupMembers = Get-ADGroupMember -Identity $groupName | ForEach-Object {
        if ($_.ObjectClass -eq 'User') {
            $userDetails = Get-ADUser -Identity $_.DistinguishedName -Properties DisplayName, SamAccountName, EmailAddress
            [PSCustomObject]@{
                GroupName       = $groupName
                DisplayName     = $userDetails.DisplayName
                SamAccountName  = $userDetails.SamAccountName
                EmailAddress    = $userDetails.EmailAddress
            }
        }
    }
    $allGroupMembers += $groupMembers
}

# Export all results to Excel
$allGroupMembers | Export-Excel -Path $outputFile -AutoSize

Write-Host "Group members exported to $outputFile"
