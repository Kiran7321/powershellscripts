

# Fetch the distribution group members
$members = Get-DistributionGroupMember -Identity "PSYCHIATRIC - Registrars"

# Sort the members alphabetically
$sortedMembers = $members | Sort-Object Name

# Add the "GG_PSYCHIATRIC" AD group to each member
foreach ($member in $sortedMembers) {
    Add-ADGroupMember -Identity "GG_PSYCHIATRIC" -Members $member
}

# Confirm the addition
Get-ADGroupMember -Identity "GG_PSYCHIATRIC"
