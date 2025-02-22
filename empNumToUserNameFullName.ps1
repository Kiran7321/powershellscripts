# Import the Active Directory module
Import-Module ActiveDirectory

# Import the Import-Excel module (if not already installed)
Import-Module ImportExcel

# Define the path to the Excel file
$excelFilePath = "C:\Users\komminenis\Documents\details.xlsx"
$outputFile = "C:\Users\komminenis\Documents\employee num details\detailsOutput1$(Get-Date -Format 'ddMM').csv"

# Import data from the Excel file
$data = Import-Excel -Path $excelFilePath

# Define the search base (the root of your AD tree, adjust as needed)
$searchBase = "DC=onepiece,DC=luffytaro"

# Loop through the data in the Excel file
foreach ($row in $data) {
    $employeeNumber = $row."Employee Num"  

    if ($employeeNumber) {
        $employeeNumber = "000" + $employeeNumber
        $filter = "(&(objectClass=user)(objectCategory=person)(postOfficeBox=$employeeNumber))" # assuming that the org uses po box to store employee number
        $user = Get-ADUser -LDAPFilter $filter -SearchBase $searchBase -Properties Description, Office

        if ($user) {
  
            $row.UserName = $user.SamAccountName
            $row.UserFullName = $user.Name

        } else {
            Write-Host "User with employee number $employeeNumber not found in Active Directory."
        }
    } else {
        Write-Host "Invalid data for employee number."
    }
}
 
$data | Export-Csv $outputFile -NoTypeInformation
