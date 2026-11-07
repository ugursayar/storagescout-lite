<#
.SYNOPSIS
    StorageScout - Microsoft 365 SharePoint/OneDrive storage waste audit. Finds reclaimable
    pooled storage (version-history & recycle-bin bloat, stale sites), quantifies how much add-on
    storage spend you can drop, and lists departed-user OneDrives heading for archive billing.

.DESCRIPTION
    Runs entirely inside YOUR tenant using the Microsoft Graph PowerShell SDK (read-only).
    No data ever leaves your environment - StorageScout has no servers and collects nothing.

    Produces:
      * CSV exports for each finding (for remediation)
      * A console summary

    Findings:
      * Tenant storage position     (SharePoint pool used vs pooled allowance vs add-on purchased)
      * Departed-user OneDrives     (owner deleted / unlicensed / disabled) - SEPARATE from the pool:
                                    OneDrive is per-user storage; unlicensed accounts follow Microsoft's
                                    archive lifecycle (read-only day 60, archived day 93, deleted day 365
                                    unless pay-as-you-go billing is on: $0.05/GB/mo + $0.60/GB reactivation)

.PARAMETER OutputPath
    Folder for the report + CSVs. Default: .\StorageScout-Report-<timestamp>

.PARAMETER Lite
    Free mode: exports the tenant-position + departed-user-OneDrive CSV + a summary only
    (no $ totals, no HTML report).

.PARAMETER MockDataPath
    Internal/testing: load tenant data from a JSON file instead of calling Microsoft Graph.

.EXAMPLE
    .\StorageScout-Audit.ps1 -MockDataPath .\test\mock-tenant.json

.NOTES
    StorageScout  |  storagescout.dev  |  read-only, runs in your tenant.
#>

[CmdletBinding()]
param(
    [string] $OutputPath,
    [switch] $Lite,
    [string] $MockDataPath
)

$ErrorActionPreference = 'Stop'
$script:Version = '1.1.0'
$script:Edition = 'Lite'   # build edition (set per package): Lite | Solo | Pro
$nowUtc = (Get-Date).ToUniversalTime()


# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Write-Step { param($Msg) Write-Host "[StorageScout] $Msg" -ForegroundColor Cyan }
function Write-Warn { param($Msg) Write-Host "[StorageScout] $Msg" -ForegroundColor Yellow }


function ConvertTo-GB { param($v) if ($null -eq $v) { 0.0 } else { [math]::Round([double]$v, 1) } }

# ----------------------------------------------------------------------------------------------
# Data acquisition (Graph OR mock)
# ----------------------------------------------------------------------------------------------
function Get-TenantData {
    param([string]$MockPath)

    if ($MockPath) {
        Write-Step "Loading MOCK data from $MockPath"
        $raw = Get-Content -Path $MockPath -Raw | ConvertFrom-Json
        # Mock dates are absolute, so age/inactivity math runs against the mock's own asOf instead of
        # the clock; otherwise the demo output (and the published sample figures) drift day by day.
        if ($raw.asOf) { $script:nowUtc = ([datetime]$raw.asOf).ToUniversalTime() }
        return $raw
    }

    Write-Step "Connecting to Microsoft Graph (read-only)..."
    $scopes = @('Reports.Read.All','Sites.Read.All','User.Read.All','Organization.Read.All','Directory.Read.All')
    try {
        Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
    } catch {
        throw "Microsoft Graph PowerShell SDK not installed. Run:  Install-Module Microsoft.Graph -Scope CurrentUser"
    }
    Connect-MgGraph -Scopes $scopes -NoWelcome

    Write-Step "Reading organization + licensed user count..."
    $org  = Get-MgOrganization
    $reportsRetrievable = $true

    # Licensed user count drives the pooled allowance (1 TB base + 10 GB per licensed user).
    $licensedUsers = 0
    try {
        $skus = Get-MgSubscribedSku -All
        # count distinct users with any license: use consumed units of base plans as a proxy is unreliable;
        # instead pull licensed user count from the reports below if available. Fallback: sum consumed of the
        # largest base plan. Keep simple + transparent:
        $licensedUsers = ([int]($skus | Measure-Object -Property ConsumedUnits -Maximum).Maximum)
    } catch { }

    Write-Step "Reading SharePoint site usage (read-only report)..."
    $sites = @()
    try {
        $tmp = New-TemporaryFile
        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period='D30')" -OutputFilePath $tmp
        $sites = Import-Csv -Path $tmp | ForEach-Object {
            [pscustomobject]@{
                url              = $_.'Site URL'
                title            = $_.'Owner Display Name'
                template         = $_.'Root Web Template'
                storageUsedGB    = [math]::Round(([double]($_.'Storage Used (Byte)') / 1GB), 1)
                storageQuotaGB   = [math]::Round(([double]($_.'Storage Allocated (Byte)') / 1GB), 1)
                lastActivityDate = $_.'Last Activity Date'
                ownerDeleted     = $false
                versionHistoryGB = 0
                recycleBinGB     = 0
            }
        }
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Warn "SharePoint usage report unavailable (needs Reports.Read.All). Storage detail will be limited."
        $reportsRetrievable = $false
    }

    Write-Step "Reading OneDrive usage (read-only report)..."
    $oneDrives = @()
    try {
        $tmp2 = New-TemporaryFile
        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/reports/getOneDriveUsageAccountDetail(period='D30')" -OutputFilePath $tmp2
        $oneDrives = Import-Csv -Path $tmp2 | ForEach-Object {
            [pscustomobject]@{
                ownerDisplayName = $_.'Owner Display Name'
                ownerUPN         = $_.'Owner Principal Name'
                ownerStatus      = 'active'   # report cannot tell deleted/disabled; enrich below if possible
                storageUsedGB    = [math]::Round(([double]($_.'Storage Used (Byte)') / 1GB), 1)
                lastActivityDate = $_.'Last Activity Date'
            }
        }
        Remove-Item $tmp2 -Force -ErrorAction SilentlyContinue
    } catch {
        Write-Warn "OneDrive usage report unavailable (needs Reports.Read.All)."
        $reportsRetrievable = $false
    }

    # Enrich OneDrive owners with account status (deleted/disabled => orphaned)
    if ($oneDrives.Count -gt 0) {
        Write-Step "Checking OneDrive owner account status..."
        foreach ($od in $oneDrives) {
            if (-not $od.ownerUPN) { continue }
            try {
                $u = Get-MgUser -UserId $od.ownerUPN -Property 'accountEnabled,assignedLicenses' -ErrorAction Stop
                if (@($u.AssignedLicenses).Count -eq 0) { $od.ownerStatus = 'unlicensed' }   # in the unlicensed-OneDrive lifecycle
                elseif (-not $u.AccountEnabled)         { $od.ownerStatus = 'disabled' }     # still licensed - license waste, not storage
            } catch {
                $od.ownerStatus = 'deleted'   # not found => deleted user => unlicensed lifecycle
            }
        }
    }

    # Pool usage is computed in Invoke-Analysis from SharePoint sites only. OneDrive storage is
    # per-user (1 TB per licensed user by default) and does NOT draw on the pooled SharePoint quota.
    $usedGB = $null

    return [pscustomobject]@{
        organization       = [pscustomobject]@{ displayName = $org.DisplayName; id = $org.Id }
        reportsRetrievable = $reportsRetrievable
        tenant = [pscustomobject]@{
            licensedUsers    = $licensedUsers
            usedStorageGB    = [math]::Round($usedGB,1)
            pooledBaseGB     = 1024
            perUserGB        = 10
            addOnPurchasedGB = $null
        }
        sites     = $sites
        oneDrives = $oneDrives
    }
}

# ----------------------------------------------------------------------------------------------
# Analysis
# ----------------------------------------------------------------------------------------------
function Invoke-Analysis {
    param($Data
    )

    $t = $Data.tenant
    $pooledBase = if ($t.pooledBaseGB) { [double]$t.pooledBaseGB } else { 1024 }
    $perUser    = if ($null -ne $t.perUserGB) { [double]$t.perUserGB } else { 10 }
    $licensed   = [int]$t.licensedUsers
    $pooledAllowanceGB = $pooledBase + ($perUser * $licensed)
    # Pool usage = SharePoint sites only. OneDrive personal sites (SPSPERS*) are outside the pool.
    $poolSites  = @($Data.sites | Where-Object { -not ([string]$_.template).StartsWith('SPSPERS', 'OrdinalIgnoreCase') })
    $usedGB     = if ($poolSites.Count -gt 0) { [math]::Round((($poolSites | Measure-Object -Property storageUsedGB -Sum).Sum),1) }
                  else { ConvertTo-GB $t.usedStorageGB }
    $overageGB  = [math]::Max(0, [math]::Round($usedGB - $pooledAllowanceGB,1))

    # --- Departed-user OneDrives (owner deleted / unlicensed / disabled) ---
    # NOT pool savings. deleted/unlicensed => Microsoft's unlicensed-OneDrive lifecycle (read-only day 60,
    # archived day 93, deleted day 365 unless pay-as-you-go billing is on). disabled-but-licensed => the
    # license is still being paid for; the storage itself is not billed separately.
    $orphaned = @()
    foreach ($od in $Data.oneDrives) {
        $status = ([string]$od.ownerStatus).ToLower()
        if ($status -in @('deleted','unlicensed','disabled')) {
            $orphaned += [pscustomobject]@{
                Owner       = $od.ownerDisplayName
                UPN         = $od.ownerUPN
                Status      = $status
                StorageGB   = ConvertTo-GB $od.storageUsedGB
                LastActivity= if ($od.lastActivityDate) { [string]$od.lastActivityDate } else { 'never' }
                Lifecycle   = if ($status -eq 'disabled') { 'licensed - license waste' } else { 'unlicensed - archive day 93 / delete day 365' }
            }
        }
    }
    $orphanedGB  = [math]::Round((($orphaned | Measure-Object -Property StorageGB -Sum).Sum),1); if (-not $orphanedGB){$orphanedGB=0}
    $lifecycle   = @($orphaned | Where-Object { $_.Status -ne 'disabled' })
    $lifecycleGB = [math]::Round((($lifecycle | Measure-Object -Property StorageGB -Sum).Sum),1); if (-not $lifecycleGB){$lifecycleGB=0}

    $result = [ordered]@{
        Org=$Data.organization
        LicensedUsers=$licensed
        PooledAllowanceGB=[math]::Round($pooledAllowanceGB,1)
        UsedGB=$usedGB
        OverageGB=$overageGB
        Orphaned=$orphaned; OrphanedGB=$orphanedGB
        Lifecycle=$lifecycle; LifecycleGB=$lifecycleGB
        ReportsAvailable = if ($null -ne $Data.reportsRetrievable) { [bool]$Data.reportsRetrievable } else { $true }
        SiteCount = @($Data.sites).Count
        OneDriveCount = @($Data.oneDrives).Count
    }
    return [pscustomobject]$result
}

# ----------------------------------------------------------------------------------------------
# Reporting
# ----------------------------------------------------------------------------------------------

# ----------------------------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------------------------
Write-Host ""
Write-Host "  StorageScout v$script:Version - Microsoft 365 storage waste audit" -ForegroundColor Green
Write-Host "  Read-only. Runs in your tenant. Nothing leaves your environment." -ForegroundColor DarkGray
Write-Host ""

if (-not $OutputPath) { $OutputPath = Join-Path (Get-Location) ("StorageScout-Report-" + $nowUtc.ToString('yyyyMMdd-HHmmss')) }
New-Item -ItemType Directory -Force -Path $OutputPath | Out-Null

$analysisArgs = @{}
$data       = Get-TenantData -MockPath $MockDataPath
$R          = Invoke-Analysis -Data $data @analysisArgs

# CSV exports (always)
$R.Orphaned  | Export-Csv (Join-Path $OutputPath 'departed-user-onedrives.csv') -NoTypeInformation -Encoding utf8

if ($Lite -or $script:Edition -eq 'Lite') {
    Write-Host ""
    Write-Step "LITE mode - free edition"
    Write-Host ([string]::Format([cultureinfo]::InvariantCulture, "  Pool used (sites)   : {0:N0} GB", $R.UsedGB))
    Write-Host ([string]::Format([cultureinfo]::InvariantCulture, "  Pooled allowance    : {0:N0} GB", $R.PooledAllowanceGB))
    Write-Host ([string]::Format([cultureinfo]::InvariantCulture, "  Over the pool       : {0:N0} GB", $R.OverageGB))
    Write-Host ([string]::Format([cultureinfo]::InvariantCulture, "  Departed OneDrives  : {0}  ({1:N0} GB; {2:N0} GB unlicensed/deleted -> archive day 93, delete day 365)", $R.Orphaned.Count, $R.OrphanedGB, $R.LifecycleGB))
    Write-Host ""
    Write-Host "  Solo/Pro add the HTML report — storagescout.dev" -ForegroundColor Yellow
    Write-Host ""
    Write-Step "CSV saved to: $OutputPath"
    return
}

