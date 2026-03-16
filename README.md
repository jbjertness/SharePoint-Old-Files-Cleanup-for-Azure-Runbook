# SharePoint Old Files Cleanup

A PowerShell script that uses **Microsoft Graph API** to recursively scan a SharePoint Online document library (or subfolder) and **permanently delete** old files older than a specified number of years — optionally limited to specific file extensions (e.g., Excel files like `.xlsx`, `.xlsm`, etc.).

Built for scenarios where you need to reclaim storage space in SharePoint sites by removing outdated documents, while respecting retention/lock policies (skips locked/checked-out files).

**Important**: Deletions are **permanent** (they go to the first-stage recycle bin, but can still be recovered there for ~93 days by default). Always test with **DryRun** first!

## Features

- Authenticates via **managed identity** (ideal for Azure Automation / Functions) or interactive login
- Targets a specific SharePoint site, document library, and optional subfolder
- Deletes files older than X years based on **last modified date**
- Optional file-type filtering (e.g., only Excel files)
- Handles pagination for large folders/libraries
- Retries on transient Graph errors (429 throttling, 5xx)
- Skips locked files (423 Locked) and reports them
- Dry-run mode to preview deletions without making changes
- Clear summary and statistics at the end

## Prerequisites

- PowerShell 7+ recommended (works in Windows PowerShell 5.1 too, but slower)
- The `Az.Accounts` module (for managed identity auth)
- Permissions:
  - **Application** permissions (for managed identity): `Sites.ReadWrite.All` (or `Files.ReadWrite.All`)
  - **Delegated** permissions (interactive): `Files.ReadWrite.All` or equivalent
- Run in an environment with access to Azure AD (e.g., Azure Runbook, local with login, VM with managed identity)

## How It Works (High-Level)

1. Authenticates and gets an access token for Graph API
2. Resolves the site → drive (library) → folder
3. Recursively walks the folder tree using /children endpoint
4. For each file:
  - Checks extension (if filtered)
  - Checks lastModifiedDateTime against cutoff period
  - Deletes via DELETE /drives/{driveId}/items/{itemId}

5. Reports locked files, skips, and totals

## Limitations & Warnings

- No version history cleanup — only deletes current files (not old versions)
- Recycle bin — deleted items go to the site's recycle bin (recoverable for ~93 days based on SharePoint defaults)
- Retention policies / legal holds — may prevent deletion (script skips locked items)
- Throttling — large libraries will take hours/days; the script will continue to retry until finished or cancelled but patient 😌
- No restore — use DryRun religiously before real runs
- Not multi-threaded (to avoid aggressive throttling)

## Installation / Setup

1. Clone or download this repository
2. Open `Cleanup-OldSharePointFiles.ps1` (or whatever you name the script)
3. Edit the **CONFIG** section at the top:

```powershell
$SiteUrlHostAndPath = "yourtenant.sharepoint.com:/sites/yoursite"   # e.g. contoso.sharepoint.com:/sites/TeamSite
$LibraryName        = "Shared Documents"                            # or "Documents", custom name, etc.
$FolderPath         = "Archives/2020"                               # leave empty "" for entire library
$KeepYears          = 2                                             # keep files modified in last 2 years
$DryRun             = $true                                         # change to $false to actually delete
$PageSize           = 200
$DeleteExtensions   = @(".xlsx", ".xlsm", ".xlsb", ".xls")          # or @() for ALL file types
```
## Usage Examples
Interactive test run (dry-run on):
```powershell
$DryRun = $true
Connect-AzAccount   # or use managed identity
.\Cleanup-OldSharePointFiles.ps13
```

Production run in Azure Automation (managed identity):
- Set $DryRun = $false
- Deploy as a Runbook
- Ensure the Automation Account has Sites.ReadWrite.All Graph permission
Target only entire "Documents" library, keep 3 years, all file types:

```powershell
$FolderPath       = ""
$KeepYears        = 3
$DeleteExtensions = @()
$DryRun           = $false
```



