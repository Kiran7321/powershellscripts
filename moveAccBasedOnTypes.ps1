# Load Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Define the path to the Excel file
$ExcelFilePath = "C:\Users\komminenis\Documents\userListFromOU.xlsx"

# Open the workbook
$workbook = $excel.Workbooks.Open($ExcelFilePath)
$worksheet = $workbook.Sheets.Item(1)

# Find the last used row
$lastRow = $worksheet.Cells($worksheet.Rows.Count, 1).End(-4162).Row  # -4162 is equivalent to xlUp

# Set the column indices for the fields
$displayNameCol = 1  # Assuming DisplayName is in column A
$descriptionCol = 2  # Assuming Description is in column B
$typeCol = 3         # Assuming Type is in column C
$cernerPositionCol = 4 # Assuming Cerner position is in column D
$moveCol = 5         # New column "Move" will be in column E

# Write the header for the new "Move" column
$worksheet.Cells.Item(1, $moveCol) = "Move"

# Define the OUs
$PiratesOU = "OU=pirates,DC=onepiece,DC=strawhats"
$SnipersOU = "OU=snipers,DC=onepiece,DC=strawhats"

# Iterate through each row
for ($row = 2; $row -le $lastRow; $row++) {
    # Read the values
    $displayName = $worksheet.Cells.Item($row, $displayNameCol).Value2
    $type = $worksheet.Cells.Item($row, $typeCol).Value2

    $user = Get-ADUser -Filter { DisplayName -eq $displayName }

    if ($user) {
        try {
            if ($type -eq "Pirates") {
                Move-ADObject -Identity $user.DistinguishedName -TargetPath $PiratesOU-ErrorAction Stop
                $worksheet.Cells.Item($row, $moveCol) = "Yes"
            } elseif ($type -eq "Snipers") {
                Move-ADObject -Identity $user.DistinguishedName -TargetPath $SnipersOU -ErrorAction Stop
                $worksheet.Cells.Item($row, $moveCol) = "Yes"
            } else {
                $worksheet.Cells.Item($row, $moveCol) = "No - Type Not Specified"
            }
        } catch {
            # Handle specific error cases
            if ($_.Exception.Message -match "name that is already in use") {
                $worksheet.Cells.Item($row, $moveCol) = "Can't move - Name already exists in target OU"
            } else {
                $worksheet.Cells.Item($row, $moveCol) = "Can't move - " + $_.Exception.Message
            }
        }
    } else {
        $worksheet.Cells.Item($row, $moveCol) = "No - User Not Found"
    }
}

# Save and close the workbook
$workbook.Save()
$workbook.Close()

# Clean up
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Processing completed. The Excel file has been updated."
