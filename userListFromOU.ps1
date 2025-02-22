# Import Active Directory module
Import-Module ActiveDirectory

# Define the Organizational Unit (OU) to search
$OU = "OU=pirates,DC=onepiece,DC=strawhats"

# Define the output Excel file path
$ExcelFilePath = "C:\Users\komminenis\Documents\userListFromOU.xlsx"

# Query Active Directory to get users' DisplayName and Description
$users = Get-ADUser -Filter * -SearchBase $OU -Property DisplayName,Description | 
         Select-Object DisplayName, Description

# Create a new Excel Application COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Add a new workbook
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Set headers in Excel
$worksheet.Cells.Item(1,1) = "DisplayName"
$worksheet.Cells.Item(1,2) = "Description"

# Start adding data from the second row
$row = 2

# Loop through each user and write data to Excel
foreach ($user in $users) {
    $worksheet.Cells.Item($row, 1) = $user.DisplayName
    $worksheet.Cells.Item($row, 2) = $user.Description
    $row++
}

# Auto-fit the columns for a better view
$worksheet.Columns.Item("A:B").AutoFit()

# Save the Excel file
$workbook.SaveAs($ExcelFilePath)

# Clean up
$workbook.Close($true)
$excel.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# Garbage collection
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "User information has been exported to $ExcelFilePath"
