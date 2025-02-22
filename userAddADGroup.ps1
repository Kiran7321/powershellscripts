#########################

# Author : Kommineni, Sai
# Date : 12/11/2024
# About: This script adds an AD group to provided list of users with their display names in an excel file & creates a log file in .csv format
# Note: It seraches for exact match, so make sure that display name is exactly the same as AD, other wise you can email or employee number as inputs with some additional changes

#########################

# Import necessary modules
Import-Module ActiveDirectory
Import-Module ImportExcel  # Ensure you have this module installed

# Define the path to the Excel file, AD group name, and the log file path
$excelFilePath = "C:\Users\komminenis\Documents\userAddADGroup.xlsx"
$groupName = "Imp_SSO_Users"
$logFilePath = "C:\Users\komminenis\Documents\userAddADGroupLog.csv"

# Create an empty array to store log results
$logResults = @()

# Load the list of users from the Excel file
$usersList = Import-Excel -Path $excelFilePath

# Loop through each user in the Excel file and attempt to add to the AD group
foreach ($user in $usersList) {
    # Get the display name from the Excel file (adjust this to match your column name)
    $displayName = $user."DisplayName"

    # Find the user in AD using the display name
    $adUser = Get-ADUser -Filter { DisplayName -eq $displayName } -ErrorAction SilentlyContinue

    if ($adUser) {
        try {
            # Attempt to add the user to the AD group
            Add-ADGroupMember -Identity $groupName -Members $adUser.SamAccountName
            # If successful, add a success entry to the log
            $logResults += [PSCustomObject]@{
                DisplayName = $displayName
                Status      = "Added Successfully"
            }
            Write-Output "$displayName has been added to the group $groupName."
        } catch {
            # If there's an error adding the user, log the failure
            $logResults += [PSCustomObject]@{
                DisplayName = $displayName
                Status      = "Failed to Add - Error Adding to Group"
            }
            Write-Output "Error adding $displayName to the group $groupName."
        }
    } else {
        # Log if the user was not found in AD
        $logResults += [PSCustomObject]@{
            DisplayName = $displayName
            Status      = "Failed to Add - User Not Found"
        }
        Write-Output "User with display name $displayName not found in Active Directory."
    }
}

# Export the log results to a CSV file
$logResults | Export-Csv -Path $logFilePath -NoTypeInformation -Force
Write-Output "Log file created at $logFilePath"
