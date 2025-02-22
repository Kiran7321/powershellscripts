# Before executing the script run  Connect-AzureAD in the command line to connect to the azure ad


# Define variables
$GroupEmail = "onepiece.onmicrosoft.com"  # Replace with your group's email
$OutputFilePath = "C:\Users\komminenis\Documents\onepieceGrpMembers.csv"  # Replace with your desired file path

# Retrieve the group using the email
$group = Get-AzureADGroup -Filter "Mail eq '$GroupEmail'"

# Check if the group exists
if ($group -eq $null) {
    Write-Host "Group with email '$GroupEmail' not found."
} else {
    $groupId = $group.ObjectId

    # Retrieve group members
    $members = Get-AzureADGroupMember -ObjectId $groupId

    # Export the members to a CSV file
    $members | Select-Object DisplayName, UserPrincipalName | Export-Csv -Path $OutputFilePath -NoTypeInformation

    Write-Host "Members of group '$GroupEmail' have been exported to '$OutputFilePath'."
}

# Disconnect from Azure AD
Disconnect-AzureAD
