<#
.SYNOPSIS
    Pulls all users in the tenant and their license assignment status via
    Microsoft Graph, and exports the result to a timestamped CSV.

.DESCRIPTION
    Requires the Microsoft.Graph.Authentication and Microsoft.Graph.Users
    modules, connected via Connect-MgGraph with at least the User.Read.All
    scope (Organization.Read.All is also needed to resolve friendly SKU
    names instead of raw GUIDs).

    Reuses the same Write-Log pattern as the rest of the toolkit: every
    row's outcome is logged to a timestamped log file, and the whole
    Graph call is wrapped in try/catch so a partial failure doesn't lose
    the run.

.PARAMETER OutputPath
    Path for the exported CSV. Defaults to a timestamped file in the
    current directory, e.g. license-report_20260831_153000.csv

.PARAMETER LogDirectory
    Directory for the run's log file. Defaults to a "logs" folder next
    to the script.

.PARAMETER IncludeGuests
    By default, guest/external accounts (UserType -ne "Member") are
    excluded from the report. Pass this switch to include them.

.EXAMPLE
    Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All"
    .\Get-LicenseReport.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [string]$LogDirectory,

    # By default, guest/external accounts are excluded from the report -
    # they're not tenant employees, so counting them as "unlicensed"
    # alongside real staff skews the numbers. Pass this to include them.
    [Parameter(Mandatory = $false)]
    [switch]$IncludeGuests
)

# ---------------------------------------------------------------------------
# Shared logging (same pattern as the rest of the toolkit)
# ---------------------------------------------------------------------------
if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
    $LogDirectory = Join-Path $baseDir "logs"
}
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = Join-Path $LogDirectory "LicenseReport_$RunTimestamp.log"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
    $OutputPath = Join-Path $baseDir "license-report_$RunTimestamp.csv"
}

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

Write-Log "Starting license report run"

# ---------------------------------------------------------------------------
# Confirm we're connected before doing anything else
# ---------------------------------------------------------------------------
try {
    $context = Get-MgContext -ErrorAction Stop
    if (-not $context) {
        throw "Not connected. Run Connect-MgGraph -Scopes 'User.Read.All','Organization.Read.All' first."
    }
    Write-Log "Connected to tenant $($context.TenantId) as $($context.Account)"
} catch {
    Write-Log "FAILED: $($_.Exception.Message)" "ERROR"
    exit 1
}

# ---------------------------------------------------------------------------
# Build a SkuId -> friendly name lookup so the report shows plan names
# instead of raw GUIDs.
# ---------------------------------------------------------------------------
$skuLookup = @{}
try {
    Get-MgSubscribedSku -All -ErrorAction Stop | ForEach-Object {
        $skuLookup[$_.SkuId] = $_.SkuPartNumber
    }
    Write-Log "Loaded $($skuLookup.Count) SKU(s) for name resolution"
} catch {
    Write-Log "Could not load SKU names (falling back to raw SkuId GUIDs): $($_.Exception.Message)" "WARN"
}

# ---------------------------------------------------------------------------
# Pull all users and their license assignments
# ---------------------------------------------------------------------------
try {
    $users = Get-MgUser -All -Property "Id,DisplayName,UserPrincipalName,AccountEnabled,AssignedLicenses,UserType" -ErrorAction Stop
    Write-Log "Retrieved $($users.Count) user(s) from Graph"
} catch {
    Write-Log "FAILED to retrieve users: $($_.Exception.Message)" "ERROR"
    exit 1
}

if (-not $IncludeGuests) {
    $guestCount = ($users | Where-Object { $_.UserType -ne "Member" }).Count
    $users = $users | Where-Object { $_.UserType -eq "Member" }
    if ($guestCount -gt 0) {
        Write-Log "Excluded $guestCount guest/external account(s) from the report (use -IncludeGuests to include them)"
    }
}

# ---------------------------------------------------------------------------
# Build the report rows
# ---------------------------------------------------------------------------
$report = New-Object System.Collections.Generic.List[object]
$successCount = 0
$failureCount = 0

foreach ($user in $users) {
    try {
        $skuNames = $user.AssignedLicenses | ForEach-Object {
            if ($skuLookup.ContainsKey($_.SkuId)) { $skuLookup[$_.SkuId] } else { $_.SkuId }
        }

        $report.Add([PSCustomObject]@{
            DisplayName       = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled    = $user.AccountEnabled
            Licensed          = $user.AssignedLicenses.Count -gt 0
            LicenseCount      = $user.AssignedLicenses.Count
            LicenseSkus       = ($skuNames -join "; ")
        })

        Write-Log "Processed $($user.UserPrincipalName) (licenses: $($user.AssignedLicenses.Count))"
        $successCount++
    } catch {
        Write-Log "FAILED: $($user.UserPrincipalName) - $($_.Exception.Message)" "ERROR"
        $failureCount++
    }
}

# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------
try {
    $report | Export-Csv -Path $OutputPath -NoTypeInformation -ErrorAction Stop
    Write-Log "Exported report to $OutputPath"
} catch {
    Write-Log "FAILED to export CSV: $($_.Exception.Message)" "ERROR"
    exit 1
}

Write-Log "Run complete. Success: $successCount, Failed: $failureCount, Total: $($users.Count)"

# ---------------------------------------------------------------------------
# Summary rollup: quick at-a-glance counts without opening the CSV
# ---------------------------------------------------------------------------
$licensedCount = ($report | Where-Object { $_.Licensed }).Count
$unlicensedCount = $report.Count - $licensedCount
Write-Log "Summary: $licensedCount licensed, $unlicensedCount unlicensed"

$skuBreakdown = $report | Where-Object { $_.Licensed } |
    ForEach-Object { $_.LicenseSkus -split "; " } |
    Group-Object | Sort-Object Count -Descending

foreach ($sku in $skuBreakdown) {
    Write-Log "  $($sku.Name): $($sku.Count) user(s)"
}

if ($failureCount -gt 0) {
    Write-Log "Completed with errors. See log for details: $LogPath" "WARN"
    exit 2
}

exit 0
