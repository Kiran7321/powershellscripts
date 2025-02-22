# Import the Active Directory module
Import-Module ActiveDirectory

# Define the path to your Excel file
$excelFilePath = "C:\users\komminenis\Documents\detailsV1.xlsx"

# Load the Excel file
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelFilePath)
$worksheet = $workbook.Sheets.Item(1) # Assuming data is in the first sheet

# Get the range of the usernames column
$usernameColumn = $worksheet.Range("B2:B" + $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row) # Adjust the range as needed
$descriptionColumn = $worksheet.Range("F2:F" + $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row) # Adjust the range as needed

# Iterate through each username and update description
for ($i = 1; $i -le $usernameColumn.Rows.Count; $i++) {
    $username = $usernameColumn.Cells.Item($i, 1).Value2
    if ($username) {
        try {
            # Query Active Directory for the description
            $user = Get-ADUser -Identity $username -Properties Description
            if ($user) {
                $description = $user.Description
                $descriptionColumn.Cells.Item($i, 1).Value2 = $description
            }
        } catch {
            # Log or handle the error if needed
            Write-Output "Error retrieving data for username: $username"
        }
    }
}

# Save and close the Excel file
$workbook.Save()
$workbook.Close()
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null