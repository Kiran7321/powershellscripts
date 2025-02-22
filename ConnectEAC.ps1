param([switch]$Elevated)
if (((New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)) -eq $false)  {
    if ($elevated) {
        # tried to elevate, did not work, aborting
    } else {
        
        Start-Process powershell.exe -Verb RunAs -ArgumentList ('-noprofile -noexit -file "{0}" -elevated' -f ($myinvocation.MyCommand.Definition))
    }
    exit
}

Register-EngineEvent PowerShell.Exiting -Action {
    Remove-PSSession $eacSession 
}

function Show-Help(){
    Write-Host ""
    Write-Host "To see this message again use the command: Show-Help" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Delegate Access:" -ForegroundColor Cyan
    Write-Host "----------------"
    Write-Host "Note: AD groups need to be mail-enabled universal security groups in order to be assigned delegated permissions" -ForegroundColor Yellow
    Write-Host "To mail-enable a universal security group:" -ForegroundColor Cyan
    Write-Host "Enable-DistributionGroup -Identity `"Group Name`" -Alias `“groupalias`”" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""
    Write-Host "To mail-disable an existing mail-enabled group:" -ForegroundColor Cyan
    Write-Host "Disable-DistributionGroup -Identity `"Group Name`"" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""

    Write-Host "Assigning a group SendAs permission to an Onprem mailbox:" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------------"
    Write-Host "Note:  This needs to be done on both On-Prem and O365 Exchange with different commands" -ForegroundColor Yellow
    Write-Host "From the On-Prem Exchange Management:" -ForegroundColor Cyan
    Write-Host "Add-ADPermission -Identity `“onpremmbx`” -user `“Username`” -AccessRights ExtendedRight -ExtendedRights `"Send As`"" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""

    Write-Host "To confirm permissions applied:"
    Write-Host "Get-ADPermission -Identity `"onpremmbx`" |   where {$_.ExtendedRights -like 'Send*'} | Format-Table -Auto User,Deny,ExtendedRights" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""

    Write-Host "Assigning a group `“Send On Behalf`” permission to an Onprem mailbox:" -ForegroundColor Cyan
    Write-Host "---------------------------------------------------------------------"
    Write-Host "Set-Mailbox -Identity `"onpremmbx`" -GrantSendOnBehalfTo `“Groupname`”" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""
    
    Write-Host "To check its Send On Behalf permission:"
    Write-Host "Get-Mailbox -Identity `"onpremmbx`" |   Format-List GrantSendOnBehalfTo" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""
    
    Write-Host "Converting a   mailbox type:" -ForegroundColor Cyan
    Write-Host "----------------------------"
    Write-Host "Set-Mailbox -Identity `“MailboxIdentity`” -Type <Regular | Room |   Equipment |   Shared>" -ForegroundColor Green -BackgroundColor Black
    Write-Host ""
}

if($elevated){
        $Host.UI.RawUI.WindowTitle = "On-Prem Exchange Admin Powershell"
        Write-Host "Connecting to On-Prem Exchange Admin Powershell" 
        $Credentials = Get-Credential -message "Please login to the On-Prem Exchange Admin Console:"
        $eacSession  = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://aex03/PowerShell/ -Authentication Kerberos -Credential $Credentials
        Import-PSSession $eacSession
        Show-Help
}

