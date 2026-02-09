<#
.SYNOPSIS
    Audit user SafeSenders lists for BullPhish domains
.DESCRIPTION
    Audit all mail-enabled users' individual SafeSenders blocklists, checking for BullPhish ID sending domains. Outputs results to terminal.
    PLEASE NOTE: You MUST be connected to Exchange Online via PowerShell BEFORE running this script.
    See https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/connect-exchangeonline?view=exchange-ps for help on connecting to Exchange Online
#>


function Start-Audit {
    $users = Get-EXOMailbox -ResultSize Unlimited

$sendingDomains = @(
    "*@bp-service-support.com",
    "*@bp-securityawareness.com",
    "*@online-account.info",
    "*@myonlinesecuritysupport.com",
    "*@service-noreply.info",
    "*@banking-alerts.info",
    "*@bullphish.com",
    "*@verifyaccount.help",
    "*@suspected-fraud.info"

)

$formattedResults = @()

# Gathers the blocked senders and domain list from each user in the tenant
foreach ($i in $users){
    $results = Get-MailboxJunkEmailConfiguration $i.UserPrincipalName | where-object {$_.BlockedSendersAndDomains -ne ""}
    $formattedResults += $results | Select-Object Identity,BlockedSendersAndDomains
}

# Processes the data into individual results by user and creates a custom table out of it
$flattened = foreach ($row in $formattedResults){

    $individualBlocks = $row.BlockedSendersAndDomains -split '\s+'

    foreach ($entry in $individualBlocks){
        [PSCustomObject]@{
            User = $row.Identity
            BlockedEntry = $entry
        }
    }
}

# Filters the above results by picking out entries matching BullPhish's sending domains
foreach ($i in $sendingDomains){
    $flattened | Select-Object User,BlockedEntry | where-object {$_.BlockedEntry -like $i}
}
}


## ---- MAIN ----
try {
    Start-Audit

}
catch {
    Write-Error "Not connected to Exchange Online"
}