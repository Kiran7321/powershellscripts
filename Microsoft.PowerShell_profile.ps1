function usergroup {
    param (
        [string]$username
    )
    $groups = (Get-ADUser -Identity $username -Property MemberOf).MemberOf | ForEach-Object { (Get-ADGroup $_).Name }
    $groups -join "`r`n"
}


function get_ewsmapi {
    param (
        [string]$emailAddress
    )

    # Connect to Exchange Online (will prompt for credentials if needed)
    Connect-ExchangeOnline

    # Get CAS mailbox details for the provided email address
    Get-CASMailbox $emailAddress | Select-Object EwsEnabled, MapiEnabled

}

function set_ewsmapi {
    param (
        [string]$emailAddress
    )

    # Get CAS mailbox details for the provided email address
    Set-casmailbox gcabi@alfred.org.au -MAPIEnabled $true -EwsEnabled $true

}

function contains {
    param (
        [string]$ContainsText
    )
    Get-ADObject -Filter {Name -like "*$ContainsText*"} | Select-Object Name
}
