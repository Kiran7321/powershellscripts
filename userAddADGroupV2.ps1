# Define the users to add
$users = @("vokounb", "wisemand", "burkitts")

# Get all security groups that match the filter and store them in an array
$groups = Get-ADObject -Filter {Name -like "607*RW"} | Select-Object -ExpandProperty DistinguishedName

# Loop through each group and add each user
foreach ($group in $groups) {
    foreach ($user in $users) {
        Add-ADGroupMember -Identity $group -Members $user
    }
}
