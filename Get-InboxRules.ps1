# Script to get inbox rules from Exchange Online
param (
    [Parameter(Mandatory=$true)]
    [string]$EmailAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$RuleName
)

# Function to check if ExchangeOnlineManagement module is installed
function Check-ExchangeModule {
    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        Write-Host "ExchangeOnlineManagement module not found. Installing it now..." -ForegroundColor Yellow
        Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
    }
}

try {
    # Check and import ExchangeOnlineManagement module
    Check-ExchangeModule
    Import-Module ExchangeOnlineManagement

    # Connect to Exchange Online
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Connect-ExchangeOnline -ShowBanner:$false

    # Get all inbox rules for the mailbox
    Write-Host "`nAll Inbox Rules for ${EmailAddress}:" -ForegroundColor Green
    Get-InboxRule -Mailbox $EmailAddress | Format-List

    # If RuleName is provided, get detailed info for that specific rule
    if ($RuleName) {
        Write-Host "`nDetailed Information for Rule: ${RuleName}" -ForegroundColor Green
        Get-InboxRule -Mailbox $EmailAddress | 
            Where-Object {$_.Name -eq $RuleName} | 
            Format-List *
    }
}
catch {
    Write-Host "An error occurred: $_" -ForegroundColor Red
}
finally {
    # Disconnect from Exchange Online
    Write-Host "`nDisconnecting from Exchange Online..." -ForegroundColor Cyan
    Disconnect-ExchangeOnline -Confirm:$false
}