Import-Module Az.Accounts
Connect-AzAccount -Identity | Out-Null

# =========================
# CONFIG (EDIT THESE)
# =========================
$SiteUrlHostAndPath = "yourtenant.sharepoint.com:/sites/ExampleSite"     
$LibraryName        = "Documents"                                    
$FolderPath         = "ExamplePath/Yourpath"
$KeepYears          = 2
$DryRun             = $true            # set to $true to simulate
$PageSize           = 200

# File types to delete (leave empty @() to target ALL file types)
$DeleteExtensions   = @(".xlsx", ".xlsm", ".xlsb", ".xls") # chose file type withL: @(".pdf", ".docx", ".xlsx")

# =========================
# AUTH
# =========================
$token   = (Get-AzAccessToken -ResourceUrl "https://graph.microsoft.com").Token
$headers = @{ Authorization = "Bearer $token" }

function Invoke-Graph {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("GET","POST","PUT","PATCH","DELETE")]$Method,
        [Parameter(Mandatory=$true)][string]$Uri,
        [byte[]]$BodyBytes = $null,
        [string]$ContentType = $null
    )

    $max = 5
    for ($i=1; $i -le $max; $i++) {
        try {
            if ($null -ne $BodyBytes) {
                return Invoke-RestMethod -Headers $headers -Uri $Uri -Method $Method -Body $BodyBytes -ContentType $ContentType
            } else {
                return Invoke-RestMethod -Headers $headers -Uri $Uri -Method $Method
            }
        } catch {
            $status = $null
            try { $status = [int]$_.Exception.Response.StatusCode } catch {}

            if ($status -in 429, 500, 502, 503, 504) {
                $sleep = [Math]::Min(60, 2 * $i)
                Write-Output "Transient Graph error ($status). Retry $i/$max in ${sleep}s: $Uri"
                Start-Sleep -Seconds $sleep
                continue
            }

            throw
        }
    }

    throw "Graph call failed after $max retries: $Method $Uri"
}

# =========================
# RESOLVE SITE
# =========================
$site   = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$SiteUrlHostAndPath"
$siteId = $site.id

# =========================
# RESOLVE DRIVE
# =========================
$drives = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/drives"
$drive  = $drives.value | Where-Object { $_.name -eq $LibraryName }

if (-not $drive) {
    throw "Document library '$LibraryName' not found on site $($site.webUrl). Available: $($drives.value.name -join ', ')"
}

$driveId = $drive.id
Write-Output "Site:             $($site.webUrl)"
Write-Output "Target library:   $LibraryName ($driveId)"
Write-Output "DeleteExtensions: $(
    if ($DeleteExtensions -and $DeleteExtensions.Count -gt 0) { $DeleteExtensions -join ', ' } else { 'ALL' }
)"

# =========================
# RESOLVE FOLDER
# =========================
if ([string]::IsNullOrWhiteSpace($FolderPath)) {
    $folderItem = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root"
    $folderId   = $folderItem.id
    $displayFolderPath = $LibraryName
} else {
    $encodedPath = [System.Uri]::EscapeDataString($FolderPath).Replace("%2F","/")
    $folderItem  = Invoke-Graph -Method GET -Uri "https://graph.microsoft.com/v1.0/drives/$driveId/root:/${encodedPath}:"
    $folderId    = $folderItem.id
    $displayFolderPath = "$LibraryName/$FolderPath"
}

Write-Output "Scope:            $displayFolderPath"
Write-Output "DryRun:           $DryRun"
Write-Output "KeepYears:        $KeepYears"

# =========================
# CUTOFF
# =========================
$cutoff = (Get-Date).AddYears(-$KeepYears)
Write-Output "Deleting items older than: $cutoff"

# =========================
# RECURSIVE WALK + DELETE
# =========================
$script:deletedCount = 0
$script:scannedCount = 0
$script:skippedCount = 0
$script:lockedCount  = 0

function Process-Folder {
    param(
        [Parameter(Mandatory=$true)][string]$ParentItemId,
        [Parameter(Mandatory=$true)][string]$ParentPath
    )

    $uri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$ParentItemId/children?`$top=$PageSize"

    while ($uri) {
        $resp = Invoke-Graph -Method GET -Uri $uri

        foreach ($child in $resp.value) {
            $script:scannedCount++
            $childName = $child.name
            $childPath = "$ParentPath/$childName"

            if ($null -ne $child.folder) {
                Process-Folder -ParentItemId $child.id -ParentPath $childPath
                continue
            }

            # =========================
            # FILE TYPE FILTER
            # =========================
            if ($DeleteExtensions -and $DeleteExtensions.Count -gt 0) {
                $ext = [System.IO.Path]::GetExtension($childName).ToLowerInvariant()
                if ([string]::IsNullOrWhiteSpace($ext) -or ($DeleteExtensions -notcontains $ext)) {
                    $script:skippedCount++
                    continue
                }
            }

            $lm = [datetime]$child.lastModifiedDateTime

            if ($lm -lt $cutoff) {
                if ($DryRun) {
                    Write-Output "[DRY RUN] Would delete: $childPath  (LastModified: $lm)"
                    $script:skippedCount++
                } else {
                    $delUri = "https://graph.microsoft.com/v1.0/drives/$driveId/items/$($child.id)"
                    try {
                        Invoke-Graph -Method DELETE -Uri $delUri
                        Write-Output "🗑️ Deleted: $childPath  (LastModified: $lm)"
                        $script:deletedCount++
                    }
                    catch {
                        $status = $null
                        try { $status = [int]$_.Exception.Response.StatusCode } catch {}

                        if ($status -eq 423) {
                            $script:lockedCount++
                            Write-Output "⏭️ Skipped (LOCKED): $childPath"
                            $script:skippedCount++
                            continue
                        }

                        throw
                    }
                }
            } else {
                $script:skippedCount++
            }
        }

        $uri = $resp.'@odata.nextLink'
    }
}

Process-Folder -ParentItemId $folderId -ParentPath $displayFolderPath

Write-Output "===================================="
Write-Output "Scanned:        $script:scannedCount"
Write-Output "Deleted:        $script:deletedCount"
Write-Output "Not deleted:    $script:skippedCount"
Write-Output "Locked skipped: $script:lockedCount"
Write-Output "===================================="
