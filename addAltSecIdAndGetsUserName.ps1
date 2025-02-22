# Import the required module
Import-Module ImportExcel

# Define the path to the Excel file
$excelPath = "C:\Users\komminenis\Documents\positionChange.xlsx"

# Read the Excel file into a variable
$data = Import-Excel -Path $excelPath

# Iterate through each row in the Excel data
foreach ($row in $data) {
    $userEmail = $row.userEmail
    $trackingValue = $row.trackingValue

    # Find the user in Active Directory using their email
    $user = Get-ADUser -Filter "Mail -eq '$userEmail'" -Properties altSecurityIdentities, SamAccountName

    if ($user) {
        # Set the altSecurityIdentities attribute
        Set-ADUser -Identity $user -Add @{altSecurityIdentities=$trackingValue}

        # Update the userName column with the user's logon name
        $row.userName = $user.SamAccountName
    }
}

# Export the updated data back to the Excel file
$data | Export-Excel $excelPath 
