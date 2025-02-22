# Import the Excel module if needed (ensure the module is installed)
Import-Module ImportExcel

# Define the path to the Excel file
$ExcelFilePath = "C:\Users\komminenis\Documents\accCreationInputFile.xlsx"

# Read the Excel file
$users = Import-Excel -Path $ExcelFilePath

# Initialize the log file
$LogFilePath = "C:\Users\komminenis\Documents\account creation logs\ADAccountCreationLog.csv"
$LogData = @()

# Iterate through each user entry
foreach ($user in $users) {
    # Validate required fields
    if (-not ($user.'First Name' -and $user.'Last Name' -and $user.Password -and $user.DOB -and $user.OU -and $user.'End Date')) {
        $LogData += [PSCustomObject]@{
            DisplayName = "$($user.'Last Name'), $($user.'First Name')"
            Status      = "Skipped"
            Reason      = "Missing required data"
        }
        continue
    }

    # Initialize values
    $FirstName = $user.'First Name'
    $LastName = $user.'Last Name'
    $DisplayName = "$LastName, $FirstName"
    $Password = $user.Password
    $DOB = $user.DOB
    $OU = $user.OU
    $Incident = $user.Incident
    $EndDate = [datetime]::Parse($user.'End Date')
    $Description = "GE HealthCare ; $Incident"
    $BaseUsername = ($LastName + $FirstName.Substring(0, 1)).ToLower()
    $CurrentUsername = $BaseUsername
    $UsernameConflict = $false

    # Check for existing accounts
    $counter = 1
    while ($true) {
        $existingAccount = Get-ADUser -Filter {SamAccountName -eq $CurrentUsername} -Properties wWWHomePage
        if ($existingAccount) {
            # Check DOB conflict
            if ($existingAccount.wWWHomePage -eq $DOB) {
                # Skip user creation
                $LogData += [PSCustomObject]@{
                    DisplayName = $DisplayName
                    Status      = "Skipped"
                    Reason      = "Username exists with matching DOB"
                }
                $UsernameConflict = $true
                break
            } else {
                # Increment username with next character or skip if exhausted
                if ($FirstName.Length -ge $counter + 1) {
                    $CurrentUsername = ($BaseUsername + $FirstName.Substring($counter, 1)).ToLower()
                    $counter++
                } else {
                    $LogData += [PSCustomObject]@{
                        DisplayName = $DisplayName
                        Status      = "Skipped"
                        Reason      = "Unresolved username conflict"
                    }
                    $UsernameConflict = $true
                    break
                }
            }
        } else {
            break
        }
    }

    # If no conflict, create the account
    if (-not $UsernameConflict) {
        try {
            New-ADUser -Name $DisplayName `
                       -GivenName $FirstName `
                       -Surname $LastName `
                       -DisplayName $DisplayName `
                       -SamAccountName $CurrentUsername `
                       -UserPrincipalName "$CurrentUsername@domain.com" `
                       -Path $OU `
                       -AccountPassword (ConvertTo-SecureString $Password -AsPlainText -Force) `
                       -Enabled $true `
                       -ChangePasswordAtLogon $true `
                       -Description $Description `
                       -AccountExpirationDate $EndDate `
                       -OtherAttributes @{
                           wWWHomePage = $DOB
                           postOfficeBox = "Non-Staff"
                       }

            $LogData += [PSCustomObject]@{
                DisplayName = $DisplayName
                Status      = "Created"
                Username    = $CurrentUsername
                Password    = $Password
            }
        } catch {
            $LogData += [PSCustomObject]@{
                DisplayName = $DisplayName
                Status      = "Error"
                Reason      = $_.Exception.Message
            }
        }
    }
}

# Export the log to CSV
$LogData | Export-Csv -Path $LogFilePath -NoTypeInformation -Force

Write-Host "Script execution completed. Log saved to $LogFilePath."
