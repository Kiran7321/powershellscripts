# Import the Active Directory module
Import-Module ActiveDirectory

# Import the Import-Excel module (if not already installed)
Import-Module ImportExcel

# Define the path to the Excel file
$excelFilePath = "C:\Users\komminenis\Documents\Account Extensions.xlsm"

# Creates the output log file in the given folder with current date at the end of the file name
$outputFile = "C:\Users\komminenis\Documents\account extension logs\Account Extensions Output$(Get-Date -Format 'ddMM').csv"

# Import data from the Excel file
$data = Import-Excel -Path $excelFilePath

# Define the search base (the root of your AD tree, adjust as needed)
$searchBase = "DC=onepiece,DC=luffytaro"

# Loop through the data in the Excel file
foreach ($row in $data) {
    $employeeNumber = $row."Employee Num"
    $expirationDate = $row."Expiry Date 1"

    if ($employeeNumber) {
        $employeeNumber = "000" + $employeeNumber
        $filter = "(&(objectClass=user)(objectCategory=person)(postOfficeBox=$employeeNumber))"
        $user = Get-ADUser -LDAPFilter $filter -SearchBase $searchBase -Properties Description, Office

        if ($user) {
            $fullName = $user.Name

            if ($expirationDate -eq '01/01/2101') {
                # Clear the account expiration if expiration date is '01/01/2101'
                Clear-ADAccountExpiration -Identity $user.DistinguishedName
                Write-Host "Account for user $fullName set to never expire (expiration date is '01/01/2101')."
            } else {
                # Set the account expiration date to the exact date from the Excel file
                Set-ADAccountExpiration -Identity $user.DistinguishedName -DateTime $expirationDate
                Write-Host "Account expiration date for user $fullName set to $expirationDate."
            }

            # Retrieve username, fullname, and description from user account in Active Directory

            $row.UserName = $user.SamAccountName
            $row.UserFullName = $fullName
            $row.UserDescription = Get-AdUser -Identity $user -Properties Description | Select-Object -ExpandProperty Description

        } else {
            Write-Host "User with employee number $employeeNumber not found in Active Directory."
        }
    } else {
        Write-Host "Invalid data for employee number."
    }
}
 
$data | Export-Csv $outputFile -NoTypeInformation

# Send-MailMessage -From "s.kommineni@alfred.org.au" -To "its.helpdesk@alfred.org.au" -Subject "test mail from powershell" -Body "Successful" -Credential "s.kommineni@alfred.org.au" -Attachments "C:\Users\komminenis\Documents\account extension logs\Account Extensions Output0911.csv" -SmtpServer "smtp.office365.com" -Port 587 -UseSsl 
