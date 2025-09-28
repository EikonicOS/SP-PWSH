<#
Top-level folder rollup for a SharePoint "Documents" library (Microsoft Graph, app-only auth)
- Computes recursive FileCount, SubfolderCount, SizeBytes/MB/GB per TOP-LEVEL folder
- Fast: processes folders in parallel and appends rows to CSV as it goes
- Shows macro counter [i/N]; optional per-folder progress (disabled by default in parallel)

Prereqs (one-time):
  1) App registration with Graph Application permissions:
       - Sites.Read.All (or Sites.Selected with site grants)
     Admin consent required.
  2) Certificate uploaded to the app (keep private key on runner).
  3) PowerShell 7+ and:
       Install-Module Microsoft.Graph.Authentication -Scope CurrentUser

Edit the TenantId/ClientId/Thumbprint below.
#>

param(
  [string]$HostName    = "hashiragg.sharepoint.com",      # tenant host
  [string]$SitePath    = "sites/House",                   # site path
  [string]$LibraryName = "Documents",                     # library name
  [string]$OutputCsv   = "C:\Reports\Documents-TopFolders.csv",
  [string]$TopFolderName,                                 # (optional) only this top-level folder
  [int]   $Throttle    = 6                                 # parallel workers (4–8 is typical)
)

$ErrorActionPreference = 'Stop'

# ---- App-only Graph connection settings (FILL THESE IN) ----
$TenantId = "<TENANT-ID>"
$ClientId = "<APP-CLIENT-ID>"
$Thumb    = "<CERT-THUMBPRINT>"

# Core modules
Import-Module Microsoft.PowerShell.Utility -ErrorAction Stop
Import-Module Microsoft.PowerShell.Management -ErrorAction Stop
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# Connect (app-only) in the main session
if (-not (Get-MgContext)) {
  Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumb -NoWelcome | Out-Null
}

# ---------- Helpers (main session) ----------
function Resolve-SiteId {
  param([string]$HostName, [string]$SitePath)
  $cleanSitePath = $SitePath.TrimStart('/')
  $byPathUri = "/v1.0/sites/${HostName}:/$cleanSitePath"
  try {
    $site = Invoke-MgGraphRequest -Method GET -Uri $byPathUri
    if ($site.id) { return $site.id }
  } catch {
    Write-Warning "Direct site path lookup failed (${HostName}:/$cleanSitePath). Falling back to search..."
  }
  $name = ($cleanSitePath.Split('/') | Select-Object -Last 1)
  $search = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites?search=$name"
  $cand = $search.value | Where-Object { $_.webUrl -like "https://$HostName/*" } | Select-Object -First 1
  if ($cand) { return $cand.id }
  throw "Could not resolve site id for host '${HostName}' and path '$SitePath'."
}

# ---------- Resolve site + library ----------
$siteId = Resolve-SiteId -HostName $HostName -SitePath $SitePath
$drives = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/sites/$siteId/drives"
$drive  = $drives.value | Where-Object name -eq $LibraryName | Select-Object -First 1
if (-not $drive) { $available = ($drives.value.name -join ', '); throw "Library '$LibraryName' not found. Available: $available" }
$driveId = $drive.id

# Get top-level folders
$topChildren = Invoke-MgGraphRequest -Method GET -Uri "/v1.0/drives/$driveId/root/children?`$top=999&`$select=id,name,folder"
$topFolders  = $topChildren.value | Where-Object { $_.folder -ne $null }

if (-not $topFolders) {
  Write-Host "No top-level folders found in '$LibraryName'." -ForegroundColor Yellow
  New-Item -ItemType File -Force -Path $OutputCsv | Out-Null
  '"TopFolder","FolderPath","FileCount","SubfolderCount","SizeBytes","SizeMB","SizeGB"' | Out-File $OutputCsv -Encoding UTF8
  Write-Host "Wrote empty CSV: $OutputCsv"
  exit
}

# Filter to single folder if requested
if ($TopFolderName) {
  $one = $topFolders | Where-Object { $_.name -ieq $TopFolderName } | Select-Object -First 1
  if (-not $one) { $one = $topFolders | Where-Object { $_.name -like $TopFolderName } | Select-Object -First 1 }
  if (-not $one) {
    $available = ($topFolders.name -join ', ')
    throw "Top-level folder '$TopFolderName' not found. Available: $available"
  }
  $topFolders = @($one)
}

# Macro counter
$totalFolders = $topFolders.Count
$folderIndex  = 0
Write-Host ("Found {0} top-level folder(s) to scan." -f $totalFolders) -ForegroundColor Cyan

# Prepare CSV and header (so parallel workers can append)
$null = New-Item -ItemType File -Path $OutputCsv -Force
'"TopFolder","FolderPath","FileCount","SubfolderCount","SizeBytes","SizeMB","SizeGB"' | Out-File $OutputCsv -Encoding UTF8

# Share serializable inputs to runspaces
$runspaceArgs = @{
  TenantId   = $TenantId
  ClientId   = $ClientId
  Thumb      = $Thumb
  DriveId    = $driveId
  OutputCsv  = $OutputCsv
}

# Show a simple macro counter as items start
$topFolders | ForEach-Object {
  $script:folderIndex++
  Write-Host ("[{0}/{1}] Queued: {2}" -f $script:folderIndex,$totalFolders,$_.name) -ForegroundColor DarkGray
} | Out-Null
$folderIndex = 0  # reset for parallel loop display (each worker logs its own completion)

# -------- Parallel scan --------
$topFolders | ForEach-Object -Parallel {
  param($TenantId, $ClientId, $Thumb, $DriveId, $OutputCsv)

  # Each runspace: lightweight imports + app-only connect
  Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
  if (-not (Get-MgContext)) {
    Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $Thumb -NoWelcome | Out-Null
  }

  # Paged children fetcher
  function Get-ChildrenPaged {
    param([string]$DriveId, [string]$ItemId)
    $uri = "/v1.0/drives/$DriveId/items/$ItemId/children?`$top=999&`$select=id,name,folder,parentReference,size"
    while ($true) {
      $resp = Invoke-MgGraphRequest -Method GET -Uri $uri
      foreach ($c in $resp.value) { $c }
      $next = $resp.'@odata.nextLink'
      if (-not $next) { break }
      $uri = $next
    }
  }

  $folderName = $PSItem.name
  $folderId   = $PSItem.id

  # Traverse this folder
  $fileCount = 0; $subfolderCount = 0; [int64]$bytes = 0
  $stack = [System.Collections.Stack]::new()
  $stack.Push($PSItem)

  while ($stack.Count -gt 0) {
    $current = $stack.Pop()
    if ($null -ne $current.folder) {
      if ($current.id -ne $folderId) { $subfolderCount++ }
      foreach ($child in Get-ChildrenPaged -DriveId $DriveId -ItemId $current.id) {
        if ($null -ne $child.folder) {
          $stack.Push($child)
        } else {
          $fileCount++
          $bytes += [int64]$child.size
        }
      }
    }
  }

  $mb = [Math]::Round(($bytes / 1MB), 2)
  $gb = [Math]::Round(($bytes / 1GB), 4)
  $csvLine =
    '"' + ($folderName -replace '"','""') + '","' +
          ($folderName -replace '"','""') + '",' +
          $fileCount + ',' + $subfolderCount + ',' + $bytes + ',' + $mb + ',' + $gb

  # Append a CSV row atomically
  Add-Content -Path $OutputCsv -Value $csvLine -Encoding UTF8

  # Log per-folder completion (kept short to reduce overhead)
  Write-Host ("[done] {0} — Files:{1} Folders:{2} SizeMB:{3}" -f $folderName,$fileCount,$subfolderCount,$mb) -ForegroundColor Gray

} -ThrottleLimit $Throttle -ArgumentList $runspaceArgs.TenantId, $runspaceArgs.ClientId, $runspaceArgs.Thumb, $runspaceArgs.DriveId, $runspaceArgs.OutputCsv

Write-Host "Done. Wrote: $OutputCsv" -ForegroundColor Green
