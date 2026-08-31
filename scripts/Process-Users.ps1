<#
.SYNOPSIS
    Reads a CSV of user records, validates each row, and logs every
    success or failure to a timestamped log file.

.DESCRIPTION
    Expected CSV columns: FirstName, LastName, UserPrincipalName, Department, Email
    Each row is validated for:
      - Required fields present (FirstName, LastName, UserPrincipalName)
      - UserPrincipalName looks like a valid email/UPN format
      - Email (if present) looks like a valid email format
    Every row's outcome (success or failure) is written to a log file
    named with the run's timestamp, e.g. UserImport_20260831_142301.log

.PARAMETER CsvPath
    Path to the input CSV file.

.PARAMETER LogDirectory
    Directory where the timestamped log file will be created.
    Defaults to a "logs" folder next to the script.

.EXAMPLE
    .\Process-Users.ps1 -CsvPath .\users.csv

.EXAMPLE
    .\Process-Users.ps1 -CsvPath .\users.csv -LogDirectory C:\Logs
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [string]$LogDirectory
)

# $PSScriptRoot is only populated when the script is run as a saved .ps1 file
# in certain contexts. Fall back to the current working directory if it's
# empty so the default LogDirectory never fails to resolve.
if ([string]::IsNullOrWhiteSpace($LogDirectory)) {
    $baseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) { $PSScriptRoot } else { (Get-Location).Path }
    $LogDirectory = Join-Path $baseDir "logs"
}

# ---------------------------------------------------------------------------
# Setup: timestamped log file
# ---------------------------------------------------------------------------
if (-not (Test-Path -Path $LogDirectory)) {
    New-Item -Path $LogDirectory -ItemType Directory -Force | Out-Null
}

$RunTimestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath = Join-Path $LogDirectory "UserImport_$RunTimestamp.log"

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
        # If logging itself fails, at least surface it on the console.
        Write-Warning "Could not write to log file '$LogPath': $($_.Exception.Message)"
    }
    # Mirror to console so the run is visible interactively too.
    switch ($Level) {
        "ERROR" { Write-Host $line -ForegroundColor Red }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        default { Write-Host $line }
    }
}

# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------
$EmailPattern = '^[^@\s]+@[^@\s]+\.[^@\s]+$'

function Test-UserRow {
    param($Row)

    $errors = New-Object System.Collections.Generic.List[string]

    if ([string]::IsNullOrWhiteSpace($Row.FirstName)) {
        $errors.Add("FirstName is missing")
    }
    if ([string]::IsNullOrWhiteSpace($Row.LastName)) {
        $errors.Add("LastName is missing")
    }
    if ([string]::IsNullOrWhiteSpace($Row.UserPrincipalName)) {
        $errors.Add("UserPrincipalName is missing")
    } elseif ($Row.UserPrincipalName -notmatch $EmailPattern) {
        $errors.Add("UserPrincipalName '$($Row.UserPrincipalName)' is not a valid UPN/email format")
    }
    if (-not [string]::IsNullOrWhiteSpace($Row.Email) -and $Row.Email -notmatch $EmailPattern) {
        $errors.Add("Email '$($Row.Email)' is not a valid email format")
    }

    return $errors
}

# ---------------------------------------------------------------------------
# Load the CSV (this itself can fail: bad path, bad format, locked file, etc.)
# ---------------------------------------------------------------------------
Write-Log "Starting user import run. Source: $CsvPath"

try {
    if (-not (Test-Path -Path $CsvPath)) {
        throw "CSV file not found at path: $CsvPath"
    }
    $users = Import-Csv -Path $CsvPath -ErrorAction Stop
    Write-Log "Loaded $($users.Count) row(s) from CSV."
} catch {
    Write-Log "FAILED to load CSV: $($_.Exception.Message)" "ERROR"
    # Nothing more we can do without data — exit with a non-zero code so
    # calling scripts / schedulers can detect the failure.
    exit 1
}

# ---------------------------------------------------------------------------
# Process each row
# ---------------------------------------------------------------------------
$successCount = 0
$failureCount = 0
$rowNumber = 1  # header is row 1 in the file, data starts at row 2

foreach ($row in $users) {
    $rowNumber++
    $displayName = if ([string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
        "<row $rowNumber, no UPN>"
    } else {
        $row.UserPrincipalName
    }

    try {
        $validationErrors = Test-UserRow -Row $row

        if ($validationErrors.Count -gt 0) {
            # Treat validation failures as controlled errors, not exceptions,
            # but still route them through the same failure logging path.
            throw ($validationErrors -join "; ")
        }

        # --- Placeholder for real processing (e.g. New-MgUser / New-ADUser) ---
        # Simulate work here; replace with actual provisioning logic.
        # e.g. New-MgUser -UserPrincipalName $row.UserPrincipalName ...

        Write-Log "Processed $displayName (row $rowNumber)"
        $successCount++
    } catch {
        Write-Log "FAILED: $displayName (row $rowNumber) - $($_.Exception.Message)" "ERROR"
        $failureCount++
    }
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Log "Run complete. Success: $successCount, Failed: $failureCount, Total: $($users.Count)"

if ($failureCount -gt 0) {
    Write-Log "Completed with errors. See log for details: $LogPath" "WARN"
    exit 2
}

exit 0
