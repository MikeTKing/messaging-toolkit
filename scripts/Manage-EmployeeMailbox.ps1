<#
.SYNOPSIS
    Employee mailbox lifecycle functions for a sandbox Microsoft 365 tenant:
    New-EmployeeMailbox (provisioning) and Disable-EmployeeMailbox (offboarding).

.DESCRIPTION
    Requires:
      - ExchangeOnlineManagement module, connected via Connect-ExchangeOnline
      - Microsoft.Graph module, connected via Connect-MgGraph
        (scopes: User.ReadWrite.All, Directory.ReadWrite.All)

    Both functions log every attempt (success or failure) via the shared
    Write-Log function to a single timestamped log file for the run.

.EXAMPLE
    . .\Manage-EmployeeMailbox.ps1
    New-EmployeeMailbox -DisplayName "Alice Nguyen" -UserPrincipalName "anguyen@contoso.onmicrosoft.com" -MailboxSizeGB 25 -FullAccessDelegate "manager@contoso.onmicrosoft.com"

.EXAMPLE
    . .\Manage-EmployeeMailbox.ps1
    Disable-EmployeeMailbox -UserPrincipalName "anguyen@contoso.onmicrosoft.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$LogDirectory
)

# ---------------------------------------------------------------------------
# Shared logging (same pattern as Process-Users.ps1)
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
    $LogDirectory = Join-Path $baseDir "logs"
}
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = Join-Path $LogDirectory "MailboxLifecycle_$RunTimestamp.log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format s), $Level, $Message
    try {
        Add-Content -Path $LogPath -Value $line -ErrorAction Stop
    } catch {
        Write-Warning "Could not write to log file '$LogPath': $($_.Exception.Message)"
    }
    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
}

# ---------------------------------------------------------------------------
# New-EmployeeMailbox: provisions a mailbox, sets starting size + permission
# ---------------------------------------------------------------------------
function New-EmployeeMailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DisplayName,

        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName,

        # Starting mailbox quota in GB. Adjust to your tenant's default plan.
        [Parameter(Mandatory = $false)]
        [int]$MailboxSizeGB = 50,

        # Optional user/group to grant Full Access delegate permission
        # (e.g. a manager or shared-support account).
        [Parameter(Mandatory = $false)]
        [string]$FullAccessDelegate,

        # Optional license SKU ID to assign (e.g. from Get-MgSubscribedSku).
        # If omitted, the mailbox is created but left unlicensed — Exchange
        # will disable it after the grace period until one is assigned.
        [Parameter(Mandatory = $false)]
        [string]$LicenseSkuId
    )

    try {
        Write-Log "Creating mailbox for $UserPrincipalName ($DisplayName)"

        # Generate a temporary password the user will be forced to change.
        # Swap this for your tenant's actual password policy / secure generation.
        $tempPassword = ConvertTo-SecureString ("Temp" + (Get-Random -Minimum 10000 -Maximum 99999) + "!") -AsPlainText -Force

        # NOTE: Exchange Online's New-Mailbox does not accept -UserPrincipalName
        # (that parameter is on-premises only). In the cloud, use
        # -MicrosoftOnlineServicesID to set the UPN instead.
        New-Mailbox -Name $DisplayName `
            -DisplayName $DisplayName `
            -MicrosoftOnlineServicesID $UserPrincipalName `
            -Password $tempPassword `
            -ResetPasswordOnNextLogon $true `
            -ErrorAction Stop | Out-Null

        # Set starting mailbox size. Exchange requires IssueWarningQuota <=
        # ProhibitSendQuota <= ProhibitSendReceiveQuota, so all three must be
        # set explicitly or a default plan value can violate the ordering.
        $quota = "${MailboxSizeGB}GB"
        $warningQuota = "{0}GB" -f [math]::Round($MailboxSizeGB * 0.9, 1)
        $sendQuota = "{0}GB" -f [math]::Round($MailboxSizeGB * 0.95, 1)
        Set-Mailbox -Identity $UserPrincipalName `
            -ProhibitSendReceiveQuota $quota `
            -ProhibitSendQuota $sendQuota `
            -IssueWarningQuota $warningQuota `
            -UseDatabaseQuotaDefaults $false `
            -ErrorAction Stop

        Write-Log "Set mailbox quota for $UserPrincipalName to $quota (warning at $warningQuota)"

        if (-not [string]::IsNullOrWhiteSpace($FullAccessDelegate)) {
            Add-MailboxPermission -Identity $UserPrincipalName `
                -User $FullAccessDelegate `
                -AccessRights FullAccess `
                -InheritanceType All `
                -AutoMapping $true `
                -ErrorAction Stop | Out-Null

            Write-Log "Granted FullAccess on $UserPrincipalName to $FullAccessDelegate"
        }

        if (-not [string]::IsNullOrWhiteSpace($LicenseSkuId)) {
            # New cloud objects can take a few seconds to become visible to
            # Graph right after New-Mailbox creates them; retry briefly
            # instead of failing on a race condition.
            $mgUser = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                try {
                    $mgUser = Get-MgUser -UserId $UserPrincipalName -ErrorAction Stop
                    break
                } catch {
                    Start-Sleep -Seconds 3
                }
            }
            if (-not $mgUser) {
                throw "Could not find $UserPrincipalName in Graph after mailbox creation (replication delay)"
            }

            Set-MgUserLicense -UserId $mgUser.Id `
                -AddLicenses @(@{ SkuId = $LicenseSkuId }) `
                -RemoveLicenses @() `
                -ErrorAction Stop | Out-Null

            Write-Log "Assigned license $LicenseSkuId to $UserPrincipalName"
        } else {
            Write-Log "No LicenseSkuId provided for $UserPrincipalName - mailbox will be disabled after the grace period until one is assigned" "WARN"
        }

        Write-Log "Successfully provisioned mailbox for $UserPrincipalName"
        return $true
    } catch {
        Write-Log "FAILED: New-EmployeeMailbox for $UserPrincipalName - $($_.Exception.Message)" "ERROR"
        return $false
    }
}

# ---------------------------------------------------------------------------
# Disable-EmployeeMailbox: converts to shared, strips licenses (offboarding)
# ---------------------------------------------------------------------------
function Disable-EmployeeMailbox {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName
    )

    try {
        Write-Log "Starting offboarding for $UserPrincipalName"

        # Convert to a shared mailbox so email history remains accessible
        # without consuming a paid license.
        Set-Mailbox -Identity $UserPrincipalName -Type Shared -ErrorAction Stop
        Write-Log "Converted $UserPrincipalName to a shared mailbox"

        # Strip all currently assigned licenses via Graph.
        $mgUser = Get-MgUser -UserId $UserPrincipalName -Property "Id,AssignedLicenses" -ErrorAction Stop
        $skuIds = $mgUser.AssignedLicenses | ForEach-Object { $_.SkuId }

        if ($skuIds -and $skuIds.Count -gt 0) {
            Set-MgUserLicense -UserId $mgUser.Id `
                -AddLicenses @() `
                -RemoveLicenses $skuIds `
                -ErrorAction Stop | Out-Null
            Write-Log "Removed $($skuIds.Count) license(s) from $UserPrincipalName"
        } else {
            Write-Log "No assigned licenses found for $UserPrincipalName" "WARN"
        }

        Write-Log "Successfully offboarded $UserPrincipalName"
        return $true
    } catch {
        Write-Log "FAILED: Disable-EmployeeMailbox for $UserPrincipalName - $($_.Exception.Message)" "ERROR"
        return $false
    }
}
