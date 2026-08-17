<#
.SYNOPSIS
Adds trusted sender domains to the Exchange Online external sender AllowList.

.DESCRIPTION
Connects to an existing Exchange Online PowerShell session and prompts the user to enter one or
more domains to add to the External Sender identification AllowList (Get/Set-ExternalInOutlook).
Displays the current allow list, collects new domain entries interactively, confirms with the user
before applying changes, and updates the tenant's AllowList accordingly. Exits early if no active
Exchange Online session is detected.

.PARAMETER CurrentAllowlist
Displays the current AllowList and exits without prompting for new domains.

.EXAMPLE
PS C:\> .\Modify-OrgTrustedSenders.ps1

Runs the script, checks for an active Exchange Online session, and prompts for domains to add.

.EXAMPLE
PS C:\> .\Modify-OrgTrustedSenders.ps1 -CurrentAllowlist

Displays the current AllowList without prompting to add new domains.

.INPUTS
None. Domain entries are collected interactively via Read-Host.

.OUTPUTS
None. Writes status and the updated AllowList to the console.

.NOTES
Requires ExchangeOnlineManagement PS module and an active Exchange Online PowerShell session (Connect-ExchangeOnline) with permissions
to run Get-ExternalInOutlook and Set-ExternalInOutlook.

.LINK
https://learn.microsoft.com/powershell/module/exchange/set-externalinoutlook
#>

[CmdletBinding()]
param(
    # Switch param to list current allowlist instead of the addition loop
    [switch]$CurrentAllowlist
)

# ----FUNCTIONS----

function Get-CurrentAllowlist {
    $currentAllowList = (Get-ExternalInOutlook).AllowList
    Write-Host ($currentAllowList -join "`n")
}

function Add-TrustedSender {
    # Empty array to store entered domains later
    $newDomains = @()
    
    # Loop to add domains to the array
    while ($true) {
        $newEntry = Read-Host -Prompt "Enter domain to add to allow list (leave blank to proceed)"
        if ([string]::IsNullOrWhiteSpace($newEntry)) { break }
        $newDomains += $newEntry
    }

    # Exits if there's no domains as Set-ExternalInOutlook doesn't like null values
    if ($newDomains.Count -eq 0) {
        Write-Host "No domains entered. Exiting..." -ForegroundColor Red
        return
    }

    # Confirmation check to make sure none of the domains are misspelled or anything 
    Write-Host "`nAdding the following domains. Please double check for errors: `n" -ForegroundColor Yellow
    Write-Host ($newDomains -join "`n")

    # Confirmation logic. Adds the new domain array to the allow list if "y". Exits the script if "n"    
    while ($true) {
            $confirmMarker = Read-Host -Prompt "`nPlease confirm if this list is correct (y/n)"
        switch($confirmMarker) {
            "y" {
                Write-Host "Adding new domains..."
                Set-ExternalInOutlook -AllowList @{Add=$newDomains}
                Write-Host "`nUpdated allowlist:" -ForegroundColor Cyan
                Get-CurrentAllowlist
                return
            }
            "n" {
                Write-Host "Exiting..." -ForegroundColor Red
                return
            }

            default {
                Write-Host "Invalid option, please enter 'y' or 'n'." -ForegroundColor Red
                
            }
        }
    }
    }


# ----MAIN----

if (-not (Get-ConnectionInformation)) {
    Write-Host "Not connected to Exchange Online! Run Connect-ExchangeOnline and sign in with admin credentials to proceed." -ForegroundColor Red
    exit
}

if ($CurrentAllowlist) {
    Get-CurrentAllowlist
    exit
}

Add-TrustedSender