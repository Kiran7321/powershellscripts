﻿# Define the OU for users
$ou = "OU=pirates,DC=onepiece,DC=strawhats"

# Get all users in the specified OU
$users = Get-ADUser -Filter * -SearchBase $ou -Properties Manager, DisplayName

# Define the username or email of the new manager
$newManagerUsername = "golddr"  # Update with the username or email of the new manager

# Get the user object of the new manager
$newManager = Get-ADUser -Filter "SamAccountName -eq '$newManagerUsername' -or EmailAddress -eq '$newManagerUsername'"

if ($newManager) {
    # Loop through each user and update the Manager field
    foreach ($user in $users) {
        $managerFieldValue = $user.Manager

        # If Manager field matches the old manager's canonical name, update it
        if ($managerFieldValue -eq "CN=Luffy\, D monkey,OU=pirates,DC=onepiece,DC=strawhats") {
            Set-ADUser -Identity $user.DistinguishedName -Manager $newManager.DistinguishedName
            Write-Host "Updated manager for $($user.DisplayName) successfully."
        }
    }
} else {
    Write-Host "New manager not found."
}
