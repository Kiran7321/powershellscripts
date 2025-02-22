# Import the Active Directory module
Import-Module ActiveDirectory

# Import the Import-Excel module (if not already installed)
# Import-Module ImportExcel  # Commented out as it's not used in the script

# Read input CSV file
$inputFile = "C:\Users\komminenis\Documents\existUsers.csv"
$outputFile = "C:\Users\komminenis\Documents\outputExistUsers2.csv"
$data = Import-Csv $inputFile

# Define the search base (the root of your AD tree, adjust as needed)
$searchBase = "DC=baysidehealth,DC=intra"

# Loop through each row in the input data
foreach ($row in $data) {
    $fullName = $row.FullName.Trim()  # Ensure there are no extra spaces

    # use it for debugging - Write-Output "Checking user: $fullName"

    # Check if user exists in Active Directory
    $user = Get-ADUser -Filter {DisplayName -eq $fullName} -SearchBase $searchBase

    # Update AccountExists column and UserName column
    if ($user) {
        $row.AccountExists = "Yes"
        $row.UserName = $user.SamAccountName

        # Retrieve description from user account in Active Directory
        $row.Description = Get-AdUser -Identity $user -Properties Description | Select-Object -ExpandProperty Description
    } else {
        $row.AccountExists = "No"
        $row.UserName = ""
        $row.Description = ""
    }
}

# Export updated data to CSV
$data | Export-Csv $outputFile -NoTypeInformation

Write-Output "Script completed. Output written to $outputFile"
