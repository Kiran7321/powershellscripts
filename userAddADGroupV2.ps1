# Define the users to add
$users = @("luffy", "chopper", "nami")

# Get all security groups that match the filter and store them in an array
$groups = Get-ADObject -Filter {Name -like "onepiece*RW"} | Select-Object -ExpandProperty DistinguishedName

# Loop through each group and add each user
foreach ($group in $groups) {
    foreach ($user in $users) {
        Add-ADGroupMember -Identity $group -Members $user
    }
}
