# Import the necessary module for Active Directory
Import-Module ActiveDirectory

# Define the OU for new users and the log file (Excel format)
$ou = "OU=Agency Nurses,OU=Users,OU=External, OU= Baysidehealth, DC=baysidehealth,DC=intra"
$logFile = "C:\Users\komminenis\Documents\agency account logs\Agency Nurse Account Log.xlsx"

# Initialize the log data array to store the logging information
$logData = @()

# Get the current date and calculate the expiry date (6 months from now)
$expiryDate = (Get-Date).AddMonths(6)

# Import the Excel file containing user details
$excelFilePath = "C:\Users\komminenis\Documents\Agency Nurse Accounts.xlsx"
$userData = Import-Excel -Path $excelFilePath

# Function to check if the altSecurityIdentities contains the AHPRA value
function AccountExists($ahpraValue) {
    $existingUsers = Get-ADUser -Filter * -Properties altSecurityIdentities | 
        Where-Object { $_.altSecurityIdentities -like "*$ahpraValue*" }
    
    return $existingUsers.Count -gt 0
}

# Function to generate a unique username
function GenerateUsername($firstName, $lastName) {
    $username = $lastName + $firstName.Substring(0, 1).ToLower()
    $index = 1

    while (Get-ADUser -Filter {sAMAccountName -eq $username}) {
        $username = $lastName + $firstName.Substring(0, $index + 1).ToLower()
        $index++
    }
    
    return $username
}

# Function to create a new user
function CreateUser($firstName, $lastName, $ahpra, $srNumber, $dob, $password) {
    $username = GenerateUsername $firstName $lastName
    $password = ConvertTo-SecureString $password -AsPlainText -Force
    $formattedAHPRA = "AHPRA=$ahpra"

    # Create the user without setting altSecurityIdentities or HomePage initially
    New-ADUser `
        -SamAccountName $username `
        -UserPrincipalName "$username@baysidehealth.intra" `
        -Name "$lastName, $firstName" `
        -GivenName $firstName `
        -Surname $lastName `
        -DisplayName "$lastName, $firstName" `
        -Description "Agency Nurse ; SR $srNumber" `
        -AccountExpirationDate $expiryDate `
        -Path $ou `
        -Enabled $true `
        -ChangePasswordAtLogon $true `
        -AccountPassword $password

    # Now, set the altSecurityIdentities attribute and HomePage (DOB), these can't be set using New-ADUser
    Set-ADUser -Identity $username -Add @{altSecurityIdentities=$formattedAHPRA}
    Set-ADUser -Identity $username -Replace @{HomePage=$dob}  # Setting DOB in HomePage

    # Add the creation log to the array
    $logData += [PSCustomObject]@{
        Status = "Created"
        Username = $username
        AHPRA = $ahpra
        Password = $password
        DOB = $dob
        Timestamp = Get-Date
    }
}

# Loop through each row in the Excel file
foreach ($row in $userData) {
    $firstName = $row.FirstName
    $lastName = $row.LastName
    $ahpra = $row.AHPRA
    $srNumber = $row.SRNumber
    $dob = $row.DOB  # Adding the DOB from the Excel file
    $password = $row.Password  # Replace with your preferred password policy

    # Check if the account already exists by checking altSecurityIdentities contains AHPRA
    if (AccountExists $ahpra) {
        echo "yes"
        $logData += [PSCustomObject]@{
            Status = "Skipped"
            Username = "N/A"
            AHPRA = $ahpra
            Password = "N/A"
            DOB = $dob
            Timestamp = Get-Date
        }
    } #else {CreateUser $firstName $lastName $ahpra $srNumber $dob $password}
}

# Export the log data to an Excel file
$logData | Export-Excel -Path $logFile -AutoSize
