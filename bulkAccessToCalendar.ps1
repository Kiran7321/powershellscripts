#This script takes input of list of emails and give them view access to a calendar
# Run  Connect-ExchangeOnline before running the script


# Define the path to your Excel file
$ExcelFilePath = "C:\Users\komminenis\Documents\calendarAccess.xlsx"

# Import the Excel data
$ExcelData = Import-Excel -Path $ExcelFilePath

# Loop through each row and set permissions
foreach ($row in $ExcelData) {
    $UserEmail = $row.Email.Trim()
    Add-MailboxFolderPermission -Identity calendarnameoremail:\Calendar -User $UserEmail -AccessRights Reviewer
}
