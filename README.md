# messaging-toolkit

PowerShell tooling for two common IT-admin workflows: bulk user validation/import
logging, and Microsoft 365 employee mailbox lifecycle (provisioning + offboarding).
Built and tested against a sandbox Microsoft 365 / Exchange Online / Entra ID tenant.

## Contents

| Script | Purpose |
|---|---|
| `scripts/Process-Users.ps1` | Reads a CSV of user records, validates each row, logs every success/failure to a timestamped log file. |
| `scripts/Manage-EmployeeMailbox.ps1` | Provides `New-EmployeeMailbox` (provision + license) and `Disable-EmployeeMailbox` (offboard to shared mailbox) functions for Exchange Online / Microsoft Graph. |
| `sample-data/users.csv` | Sample CSV with a mix of valid and intentionally invalid rows, for testing `Process-Users.ps1`. |

---

## Process-Users.ps1

Validates a CSV of user records (required fields, UPN/email format) and logs
every row's outcome — success or failure — to a timestamped log file. Wraps
both the CSV load and each row's processing in try/catch so a single bad
row never crashes the run.

**Requirements:** PowerShell 5.1+ (no external modules).

**Usage:**
```powershell
.\scripts\Process-Users.ps1 -CsvPath .\sample-data\users.csv
```

**Parameters:**
| Name | Required | Description |
|---|---|---|
| `-CsvPath` | Yes | Path to the input CSV. Expected columns: `FirstName, LastName, UserPrincipalName, Department, Email`. |
| `-LogDirectory` | No | Where the timestamped log is written. Defaults to a `logs` folder next to the script. |

**Output:** `logs/UserImport_<timestamp>.log`, plus a console summary. Exit code
`0` = clean run, `1` = couldn't load the CSV, `2` = completed with row-level failures.

---

## Manage-EmployeeMailbox.ps1

Two functions covering the employee mailbox lifecycle in Exchange Online /
Microsoft 365. Both share the same timestamped logging pattern as
`Process-Users.ps1`, and every external call (mailbox creation, quota,
permissions, licensing) is wrapped in try/catch.

**Requirements:**
- `ExchangeOnlineManagement` module, connected via `Connect-ExchangeOnline`
- `Microsoft.Graph.Authentication` module, connected via `Connect-MgGraph`
  (scopes: `User.ReadWrite.All`, `Directory.ReadWrite.All`)
- Both connections must be active **in the same PowerShell session** before
  calling either function.

**Usage:**
```powershell
Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All"
Connect-ExchangeOnline -DisableWAM

. .\scripts\Manage-EmployeeMailbox.ps1

New-EmployeeMailbox -DisplayName "Jane Doe" -UserPrincipalName "jdoe@contoso.onmicrosoft.com" `
    -MailboxSizeGB 25 -LicenseSkuId "f245ecc8-75af-4f8e-b61f-27d8114de5f3"

Disable-EmployeeMailbox -UserPrincipalName "jdoe@contoso.onmicrosoft.com"
```

### `New-EmployeeMailbox`
Creates the mailbox, sets a starting quota, optionally grants a delegate
Full Access, optionally assigns a license, and logs the outcome of each step.

| Parameter | Required | Description |
|---|---|---|
| `-DisplayName` | Yes | Display name for the new mailbox. |
| `-UserPrincipalName` | Yes | UPN for the new account. |
| `-MailboxSizeGB` | No (default `50`) | Mailbox quota in GB. Warning/prohibit-send thresholds are scaled from this. |
| `-FullAccessDelegate` | No | UPN of a user/group to grant Full Access delegate permission. |
| `-LicenseSkuId` | No | License SKU GUID to assign (see `Get-MgSubscribedSku`). If omitted, the mailbox is created unlicensed and will be disabled after the grace period. |

### `Disable-EmployeeMailbox`
Offboards a mailbox: converts it to shared (keeps mail accessible without
consuming a license) and strips every currently assigned license.

| Parameter | Required | Description |
|---|---|---|
| `-UserPrincipalName` | Yes | UPN of the mailbox to offboard. |

---

## Known issues / troubleshooting log

Real errors hit while building and testing this against a live sandbox
tenant, kept here so the fixes don't have to be rediscovered.

**`Connect-ExchangeOnline` → "A window handle must be configured"**
Known bug introduced when `ExchangeOnlineManagement` added WAM (Web Account
Manager) sign-in. Fix: `Connect-ExchangeOnline -DisableWAM`, run in a fresh
session (won't take effect if you already connected without the flag).

**`Import-Module Microsoft.Graph.Authentication` → `TypeLoadException` on `GetTokenAsync`**
Assembly conflict between the Graph SDK and an already-loaded Exchange Online
session in the same PowerShell process. Fix: connect Graph *before* Exchange
Online in a fresh window (`Connect-MgGraph` then `Connect-ExchangeOnline -DisableWAM`).

**`New-Mailbox : A parameter cannot be found that matches parameter name 'UserPrincipalName'`**
Exchange Online's `New-Mailbox` doesn't accept `-UserPrincipalName` — that
parameter is on-premises-only. In the cloud, use `-MicrosoftOnlineServicesID`
to set the UPN instead.

**`Set-Mailbox : The value of property 'ProhibitSendReceiveQuota' must be greater than or equal to that of property 'ProhibitSendQuota'`**
Exchange requires `IssueWarningQuota <= ProhibitSendQuota <= ProhibitSendReceiveQuota`.
Setting only `ProhibitSendReceiveQuota` leaves `ProhibitSendQuota` at the
mailbox plan's default, which can be larger than your requested quota. Fix:
set all three explicitly.

**Graph object not found immediately after `New-Mailbox`**
Newly created cloud objects can take a few seconds to replicate before
they're visible to Microsoft Graph. `New-EmployeeMailbox` retries
`Get-MgUser` briefly before failing, to absorb this delay.
