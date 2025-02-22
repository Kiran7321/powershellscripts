

# Fetch the distribution group members
$members = Get-DistributionGroupMember -Identity "Strawhats"

# Sort the members alphabetically
$sortedMembers = $members | Sort-Object Name

# Add the AD group to each member
foreach ($member in $sortedMembers) {
    Add-ADGroupMember -Identity "gg_winners" -Members $member
}

# Confirm the addition
Get-ADGroupMember -Identity "gg_winners"
