############################################

# Author         :   Kommineni, Sai
# Date           :   23/11/2024
# Functionality  :   This script will take computer names either full / last few digits as input and add a specified AD group to them

#################################################

# Load the ImportExcel module
Import-Module ImportExcel

# Path to the Excel file containing computer names
$excelFilePath = "C:\Users\komminenis\Documents\addADGroupToPCs.xlsx" # Update the path to your Excel file

# Sheet name in the Excel file - Update the sheet name if different
$sheetName = "Sheet1"

# Column name in the Excel file that contains the computer names
$columnName = "Last6Digit" 

# AD group name - Replace with your actual AD group name
$groupName = "BH_Onepiece_RW" 

# Path to save the log CSV file - Update the path to your log CSV file
$logFilePath = "C:\Users\komminenis\Documents\app installtion logs\addADGroupToPCsLog.csv" 

# Initialize an array to hold the log results
$logResults = @()

# Read computer names from the Excel file
$computers = Import-Excel -Path $excelFilePath -WorksheetName $sheetName | Select-Object -ExpandProperty $columnName

# Iterate over each computer name and add it to the AD group
foreach ($computer in $computers) {
    try {
        # Get the computer object from AD
        $pc = Get-ADComputer -Filter "Name -like '*$computer'" -Properties Name
        
        if ($pc) {
            # Add the computer to the specified AD group
            Add-ADGroupMember -Identity $groupName -Members $pc
            $logResults += [PSCustomObject]@{
                ComputerName = $pc.Name
                Status = "Added"
            }
        } else {
            $logResults += [PSCustomObject]@{
                ComputerName = $computer
                Status = "Not Found"
            }
        }
    } catch {
        $logResults += [PSCustomObject]@{
            ComputerName = $computer
            Status = "Error: $_"
        }
    }
}

# Export the log results to a CSV file
$logResults | Export-Csv -Path $logFilePath -NoTypeInformation

Write-Output "Operation complete. Log file saved to $logFilePath."
