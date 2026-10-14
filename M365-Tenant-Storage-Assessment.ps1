#requires -Version 7.2
<#
===============================================================================
                    ABACUS INFORMATION SYSTEMS, INC.
                  MICROSOFT 365 TENANT STORAGE ASSESSMENT
===============================================================================

Copyright (c) 2026 Abacus Information Systems, Inc. MIT licensed.

Licensed under the MIT License. See LICENSE in the project repository.
Copyright and license notices must be retained in copies or substantial
portions of the software.

The script is provided "as is," without warranty of any kind. Administrators
are responsible for reviewing, testing, and validating the script in their own
environment before production use. Microsoft 365 services, APIs, modules,
permissions, reports, pricing, and authentication behavior may change.

Owner:       Abacus Information Systems, Inc.
Author:      Brad Gerard
AI Assist:   Microsoft Copilot supplied some PowerShell and several strongly held opinions about HTML
Version:     2.4
Purpose:     Measure tenant storage with optional backup cost estimates
Environment: PowerShell 7.2 or later

===============================================================================

.SYNOPSIS
    Microsoft 365 Tenant Storage Assessment, Version 2.4.

.DESCRIPTION
    Prompts for an administrator UPN, discovers the actual SharePoint Online
    tenant hostname through Microsoft Graph, validates the SharePoint admin URL,
    and totals storage usage for Exchange Online, OneDrive, and SharePoint.

    Fast Exchange mode uses the organization-wide Microsoft Graph mailbox usage
    report for primary mailbox storage, then queries only active online archives.
    Exact mode retains the original mailbox-by-mailbox measurement method.

    Creates console, CSV, and standalone HTML summary reports. The HTML report opens automatically unless -NoOpenReport is specified. Online archives are included by default and can be omitted with -SkipArchives. Use -IncludeBackupEstimate to display optional backup cost projections.

    Teams files are included in SharePoint/OneDrive because Teams stores files
    in those workloads. Teams chat/message data is not separately calculated.

.NOTES
    Run in PowerShell 7 using a Microsoft 365 administrator with the Graph,
    Exchange Online, and SharePoint Online permissions required by the script.
#>

[CmdletBinding()]
param(
    [decimal]$BackupPricePerGB = 0.15,
    [switch]$IncludeBackupEstimate,
    [ValidateSet('Fast','Exact')]
    [string]$ExchangeMode = 'Fast',
    [switch]$SkipArchives,
    [switch]$NoOpenReport,
    [string]$OutputFolder = "$env:USERPROFILE\Documents\M365-Storage-Assessments"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '2.4'

function Write-Banner {
    Clear-Host
    Write-Host ''
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    Write-Host '  |        MICROSOFT 365 TENANT STORAGE ASSESSMENT v2.4         |' -ForegroundColor Cyan
    Write-Host '  |        Exchange | OneDrive | SharePoint | HTML             |' -ForegroundColor DarkCyan
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    Write-Host ''
}

function Write-Step { param([string]$Message) Write-Host ("  [>] {0}" -f $Message) -ForegroundColor Cyan }
function Write-OK { param([string]$Message) Write-Host ("  [OK] {0}" -f $Message) -ForegroundColor Green }
function Write-WarnLine { param([string]$Message) Write-Host ("  [!] {0}" -f $Message) -ForegroundColor Yellow }

function Confirm-Value {
    param([string]$Label,[string]$DetectedValue)
    Write-Host "  $Label" -ForegroundColor DarkGray
    Write-Host "  Detected: $DetectedValue" -ForegroundColor White
    $Override = Read-Host '  Press Enter to accept, or type an override'
    if ([string]::IsNullOrWhiteSpace($Override)) { return $DetectedValue }
    return $Override.Trim()
}

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name,[Parameter(Mandatory)][string]$MinimumVersion)
    $Installed = Get-Module -ListAvailable -Name $Name | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $Installed -or $Installed.Version -lt [version]$MinimumVersion) {
        Write-WarnLine "$Name $MinimumVersion or later is required."
        $Install = Read-Host "  Install/update $Name for the current user now? [Y/n]"
        if ($Install -match '^(n|no)$') { throw "$Name is required. Install it and run the script again." }
        Install-Module -Name $Name -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber
    }
    Import-Module $Name -MinimumVersion $MinimumVersion -Force
    Write-OK "$Name is ready."
}

function Convert-ExchangeSizeToBytes {
    param($SizeValue)
    if ($null -eq $SizeValue) { return [int64]0 }
    if ($SizeValue.PSObject.Properties.Name -contains 'Value' -and $null -ne $SizeValue.Value) {
        $Inner = $SizeValue.Value
        if ($Inner.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$Inner.ToBytes() }
    }
    if ($SizeValue.PSObject.Methods.Name -contains 'ToBytes') { return [int64]$SizeValue.ToBytes() }
    $Text = [string]$SizeValue
    if ($Text -match '\(([\d,]+) bytes\)') { return [int64](($Matches[1]) -replace ',', '') }
    $Normalized = $Text.Trim() -replace ',', ''
    if ($Normalized -match '^([\d\.]+)\s*(B|KB|MB|GB|TB)') {
        $Number = [decimal]$Matches[1]
        $Multiplier = switch ($Matches[2].ToUpperInvariant()) { 'B'{1};'KB'{1KB};'MB'{1MB};'GB'{1GB};'TB'{1TB} }
        return [int64]($Number * $Multiplier)
    }
    throw "Could not parse Exchange size value: $Text"
}

function Format-DataSize {
    param([decimal]$Bytes)
    if ($Bytes -ge 1TB) { return ('{0:N2} TB' -f ($Bytes / 1TB)) }
    return ('{0:N2} GB' -f ($Bytes / 1GB))
}

function Test-SPOAdminConnection {
    param([Parameter(Mandatory)][string]$AdminUrl)
    Connect-SPOService -Url $AdminUrl -ModernAuth $true -UseSystemBrowser $true
    $null = Get-SPOTenant
    return $true
}

function Get-ReportStorageColumn {
    param([Parameter(Mandatory)]$Row)
    $Candidates = @('Storage Used (Byte)','Storage Used (Bytes)','StorageUsedInBytes','Storage Used (MB)','StorageUsedInMB')
    foreach ($Name in $Candidates) {
        if ($Row.PSObject.Properties.Name -contains $Name) {
            $Raw = [string]$Row.$Name
            if ([string]::IsNullOrWhiteSpace($Raw)) { return [decimal]0 }
            $Number = [decimal]($Raw -replace ',','')
            if ($Name -match 'MB') { return $Number * 1MB }
            return $Number
        }
    }
    throw "Mailbox usage report does not contain a recognized storage column. Found: $($Row.PSObject.Properties.Name -join ', ')"
}

function Get-ExchangeBytesFast {
    param(
        [Parameter(Mandatory)][string]$WorkingFolder,
        [switch]$SkipArchives
    )
    Write-Step 'Downloading the Microsoft Graph mailbox usage report...'
    $ReportPath = Join-Path $WorkingFolder 'MailboxUsage-D7.csv'
    # Graph Reports can emit an invalid progress percentage (Int32.MaxValue),
    # causing a red PercentComplete error even when the CSV download succeeds.
    # Suppress progress only for this Graph request; restore the caller's setting.
    $PreviousProgressPreference = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        Get-MgReportMailboxUsageDetail -Period 'D7' -OutFile $ReportPath
    }
    finally {
        $ProgressPreference = $PreviousProgressPreference
    }
    if (-not (Test-Path -LiteralPath $ReportPath)) {
        throw 'The Microsoft Graph mailbox usage report was not created.'
    }
    if ((Get-Item -LiteralPath $ReportPath).Length -eq 0) {
        throw 'The Microsoft Graph mailbox usage report is empty.'
    }
    $Rows = @(Import-Csv -Path $ReportPath)
    if ($Rows.Count -eq 0) { throw 'The Microsoft Graph mailbox usage report returned no rows.' }

    [decimal]$PrimaryBytes = 0
    foreach ($Row in $Rows) { $PrimaryBytes += Get-ReportStorageColumn -Row $Row }
    Write-OK "Primary mailbox report total: $(Format-DataSize $PrimaryBytes) across $($Rows.Count) report rows."

    if ($SkipArchives) {
        Write-WarnLine 'Online archive collection was skipped by request.'
        return [pscustomobject]@{
            TotalBytes = $PrimaryBytes
            PrimaryBytes = $PrimaryBytes
            ArchiveBytes = [decimal]0
            Method = 'Fast: Graph mailbox usage report; online archives skipped'
            ReportRows = $Rows.Count
            ArchiveCount = 0
            ArchiveSucceeded = 0
            ArchiveFailed = 0
            ArchiveStatus = 'Skipped by administrator'
            ArchiveFailures = @()
        }
    }

    Write-Step 'Measuring active online archives...'
    $ArchiveMailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -Properties ArchiveStatus,ArchiveGuid |
        Where-Object { $_.ArchiveStatus -eq 'Active' -or ($_.ArchiveGuid -and $_.ArchiveGuid -ne [guid]::Empty) })
    [decimal]$ArchiveBytes = 0
    $ArchiveFailures = [System.Collections.Generic.List[object]]::new()
    $ArchiveSucceeded = 0
    $Index = 0
    foreach ($Mailbox in $ArchiveMailboxes) {
        $Index++
        Write-Progress -Activity 'Measuring Exchange Online archives' -Status "$Index of $($ArchiveMailboxes.Count): $($Mailbox.DisplayName)" -PercentComplete (($Index / [math]::Max($ArchiveMailboxes.Count,1)) * 100)
        try {
            $Stats = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -Archive -ErrorAction Stop
            $ArchiveBytes += Convert-ExchangeSizeToBytes $Stats.TotalItemSize
            $ArchiveSucceeded++
        }
        catch {
            $Message = $_.Exception.Message
            $ArchiveFailures.Add([pscustomobject]@{
                DisplayName = $Mailbox.DisplayName
                UserPrincipalName = $Mailbox.UserPrincipalName
                Error = $Message
            })
            Write-Warning "Archive size could not be read for $($Mailbox.UserPrincipalName): $Message"
        }
    }
    Write-Progress -Activity 'Measuring Exchange Online archives' -Completed
    $ArchiveFailed = $ArchiveFailures.Count
    $ArchiveStatus = if ($ArchiveFailed -eq 0) {
        "Included: $ArchiveSucceeded of $($ArchiveMailboxes.Count) archives measured"
    } else {
        "Partial: $ArchiveSucceeded of $($ArchiveMailboxes.Count) measured; $ArchiveFailed failed or timed out"
    }
    Write-OK "Online archive total: $(Format-DataSize $ArchiveBytes) across $ArchiveSucceeded successfully measured archives."
    if ($ArchiveFailed -gt 0) {
        Write-WarnLine "$ArchiveFailed archive mailbox query failed or timed out. The Exchange total may be understated."
    }
    return [pscustomobject]@{
        TotalBytes = $PrimaryBytes + $ArchiveBytes
        PrimaryBytes = $PrimaryBytes
        ArchiveBytes = $ArchiveBytes
        Method = 'Fast: Graph mailbox usage report plus live archive statistics'
        ReportRows = $Rows.Count
        ArchiveCount = $ArchiveMailboxes.Count
        ArchiveSucceeded = $ArchiveSucceeded
        ArchiveFailed = $ArchiveFailed
        ArchiveStatus = $ArchiveStatus
        ArchiveFailures = @($ArchiveFailures)
    }
}

function Get-ExchangeBytesExact {
    param([switch]$SkipArchives)
    Write-Step 'Measuring every Exchange Online mailbox and online archive...'
    [decimal]$TotalBytes = 0
    [decimal]$PrimaryBytes = 0
    [decimal]$ArchiveBytes = 0
    $Mailboxes = @(Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox,SharedMailbox -Properties ArchiveStatus,ArchiveGuid)
    $Index = 0
    foreach ($Mailbox in $Mailboxes) {
        $Index++
        Write-Progress -Activity 'Measuring Exchange Online' -Status "$Index of $($Mailboxes.Count): $($Mailbox.DisplayName)" -PercentComplete (($Index / [math]::Max($Mailboxes.Count,1)) * 100)
        try {
            $Stats = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
            $Size = Convert-ExchangeSizeToBytes $Stats.TotalItemSize
            $PrimaryBytes += $Size
            $TotalBytes += $Size
        }
        catch { Write-Warning "Primary mailbox size could not be read for $($Mailbox.UserPrincipalName): $($_.Exception.Message)" }
        if (-not $SkipArchives -and ($Mailbox.ArchiveStatus -eq 'Active' -or ($Mailbox.ArchiveGuid -and $Mailbox.ArchiveGuid -ne [guid]::Empty))) {
            try {
                $ArchiveStats = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -Archive -ErrorAction Stop
                $Size = Convert-ExchangeSizeToBytes $ArchiveStats.TotalItemSize
                $ArchiveBytes += $Size
                $TotalBytes += $Size
            }
            catch { Write-Warning "Archive size could not be read for $($Mailbox.UserPrincipalName): $($_.Exception.Message)" }
        }
    }
    Write-Progress -Activity 'Measuring Exchange Online' -Completed
    return [pscustomobject]@{ TotalBytes = $TotalBytes; PrimaryBytes = $PrimaryBytes; ArchiveBytes = $ArchiveBytes; Method = $(if ($SkipArchives) { 'Exact: live primary mailbox statistics; online archives skipped' } else { 'Exact: live statistics queried for every mailbox and archive' }); ReportRows = $Mailboxes.Count; ArchiveCount = $null; ArchiveSucceeded = $null; ArchiveFailed = $null; ArchiveStatus = $(if ($SkipArchives) { 'Skipped by administrator' } else { 'Included in exact mailbox scan' }); ArchiveFailures = @() }
}

function New-HtmlReport {
    param([switch]$IncludeBackupEstimate,[string]$Path,[string]$TenantName,[string]$TenantId,[string]$PrimaryDomain,[string]$AdminUrl,[datetime]$Date,[decimal]$ExchangeBytes,[decimal]$PrimaryMailboxBytes,[decimal]$ArchiveBytes,[bool]$ArchivesSkipped,[decimal]$OneDriveBytes,[decimal]$SharePointBytes,[decimal]$TotalBytes,[decimal]$Monthly,[decimal]$Annual,[decimal]$Rate,[string]$ExchangeMethod,[string]$ArchiveStatus,[int]$ArchiveFailed,[string[]]$ArchiveFailureNames)
    $Encode = { param($Value) [System.Net.WebUtility]::HtmlEncode([string]$Value) }
    $TenantHtml = & $Encode $TenantName; $TenantIdHtml = & $Encode $TenantId; $DomainHtml = & $Encode $PrimaryDomain; $AdminHtml = & $Encode $AdminUrl; $MethodHtml = & $Encode $ExchangeMethod; $ArchiveStatusHtml = & $Encode $ArchiveStatus
    $ExchangePct = if ($TotalBytes -gt 0) { [math]::Round(($ExchangeBytes / $TotalBytes) * 100,1) } else { 0 }
    $OneDrivePct = if ($TotalBytes -gt 0) { [math]::Round(($OneDriveBytes / $TotalBytes) * 100,1) } else { 0 }
    $SharePointPct = if ($TotalBytes -gt 0) { [math]::Round(($SharePointBytes / $TotalBytes) * 100,1) } else { 0 }
    $Html = @"
<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Microsoft 365 Tenant Storage Assessment - $TenantHtml</title>
<style>
:root{--navy:#12324a;--cyan:#00a6c8;--green:#24a148;--gold:#f4b400;--ink:#172b3a;--muted:#617684;--paper:#f4f7f9;--card:#fff;--line:#dce5ea}*{box-sizing:border-box}body{margin:0;background:var(--paper);font-family:Segoe UI,Arial,sans-serif;color:var(--ink)}.wrap{max-width:1080px;margin:36px auto;padding:0 22px}.hero{background:linear-gradient(135deg,var(--navy),#185d78);color:#fff;padding:34px;border-radius:16px;box-shadow:0 10px 30px #16384d30}.eyebrow{letter-spacing:.13em;font-size:12px;font-weight:700;color:#9eeafa}.hero h1{margin:8px 0 4px;font-size:32px}.hero p{margin:0;color:#d9eef4}.grid{display:grid;grid-template-columns:repeat(3,1fr);gap:16px;margin-top:18px}.card{background:var(--card);border:1px solid var(--line);border-radius:14px;padding:22px;box-shadow:0 5px 18px #17384a12}.label{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);font-weight:700}.value{font-size:28px;font-weight:750;margin-top:7px}.total{border-top:5px solid var(--green)}.cost{border-top:5px solid var(--gold)}.section{margin-top:18px}.bar{display:flex;height:14px;background:#e7eef2;border-radius:20px;overflow:hidden;margin-top:14px}.segment{height:100%}.exchange{background:#f28c28}.onedrive{background:#f4c542}.sharepoint{background:#0078d4}.legend{display:flex;gap:20px;flex-wrap:wrap;margin-top:12px;color:var(--muted);font-size:13px}.dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:6px}.warning{background:#fff5e6;border-left:4px solid #f28c28;padding:14px 16px;border-radius:7px;color:#61451d}table{width:100%;border-collapse:collapse;margin-top:10px}th,td{text-align:left;padding:12px;border-bottom:1px solid var(--line)}th{color:var(--muted);font-size:12px;text-transform:uppercase}.right{text-align:right}.meta{display:grid;grid-template-columns:180px 1fr;gap:8px 16px}.note{background:#edf7fa;border-left:4px solid var(--cyan);padding:14px 16px;border-radius:7px;color:#365360}.footer{text-align:center;color:var(--muted);font-size:12px;margin:24px 0}@media(max-width:760px){.grid{grid-template-columns:1fr}.meta{grid-template-columns:1fr}.right{text-align:left}}
</style></head><body><div class="wrap"><div class="hero"><div class="eyebrow">ABACUS INFORMATION SYSTEMS, INC.</div><h1>Microsoft 365 Tenant Storage Assessment</h1><p>$TenantHtml | Generated $($Date.ToString('MMMM d, yyyy h:mm tt'))</p></div>
<div class="grid"><div class="card total"><div class="label">Total tenant data</div><div class="value">$(Format-DataSize $TotalBytes)</div></div><div class="card"><div class="label">Exchange Online</div><div class="value">$(Format-DataSize $ExchangeBytes)</div></div><div class="card"><div class="label">OneDrive + SharePoint</div><div class="value">$(Format-DataSize ($OneDriveBytes + $SharePointBytes))</div></div></div>
$(if ($IncludeBackupEstimate) { '<div class="card section"><div class="label">Optional backup cost planning</div><div class="grid"><div><div class="label">Estimated monthly cost</div><div class="value">' + ('$' + ('{0:N2}' -f $Monthly)) + '</div></div><div><div class="label">Estimated annual cost</div><div class="value">' + ('$' + ('{0:N2}' -f $Annual)) + '</div></div><div><div class="label">Rate per GB/month</div><div class="value">' + ('$' + ('{0:N2}' -f $Rate)) + '</div></div></div></div>' } else { '' })
<div class="card section"><div class="label">Storage by workload</div><table><thead><tr><th>Service</th><th class="right">Used data</th><th class="right">Share of total</th></tr></thead><tbody>
<tr><td>Exchange Online</td><td class="right">$(Format-DataSize $ExchangeBytes)</td><td class="right">$ExchangePct%</td></tr><tr><td>OneDrive</td><td class="right">$(Format-DataSize $OneDriveBytes)</td><td class="right">$OneDrivePct%</td></tr><tr><td>SharePoint</td><td class="right">$(Format-DataSize $SharePointBytes)</td><td class="right">$SharePointPct%</td></tr></tbody></table>
<div class="bar" role="img" aria-label="Workload distribution: Exchange $ExchangePct percent, OneDrive $OneDrivePct percent, SharePoint $SharePointPct percent"><div class="segment exchange" style="width:$ExchangePct%"></div><div class="segment onedrive" style="width:$OneDrivePct%"></div><div class="segment sharepoint" style="width:$SharePointPct%"></div></div><div class="legend"><span><i class="dot exchange"></i>Exchange $ExchangePct%</span><span><i class="dot onedrive"></i>OneDrive $OneDrivePct%</span><span><i class="dot sharepoint"></i>SharePoint $SharePointPct%</span></div></div>
<div class="card section"><div class="label">Exchange Online storage breakdown</div><table><thead><tr><th>Mailbox storage</th><th class="right">Used data</th></tr></thead><tbody><tr><td>Primary mailboxes</td><td class="right">$(Format-DataSize $PrimaryMailboxBytes)</td></tr><tr><td>Online archives</td><td class="right">$(if ($ArchivesSkipped) { 'Not measured (skipped)' } else { Format-DataSize $ArchiveBytes })</td></tr><tr><td><strong>Exchange Online total</strong></td><td class="right"><strong>$(Format-DataSize $ExchangeBytes)</strong></td></tr></tbody></table></div>
<div class="card section"><div class="label">Assessment details</div><div class="meta"><strong>Tenant</strong><span>$TenantHtml</span><strong>Tenant ID</strong><span>$TenantIdHtml</span><strong>Primary domain</strong><span>$DomainHtml</span><strong>SharePoint admin</strong><span>$AdminHtml</span><strong>Exchange method</strong><span>$MethodHtml</span><strong>Archive collection</strong><span>$ArchiveStatusHtml</span>$(if ($IncludeBackupEstimate) { "<strong>Backup planning rate</strong><span>" + ('$' + ('{0:N2}/GB/month' -f $Rate)) + "</span>" } else { "" })</div></div>
$(if ($ArchiveFailed -gt 0) { '<div class="warning section"><strong>Archive warning:</strong> ' + $ArchiveFailed + ' archive mailbox query failed or timed out. The Exchange total may be understated. Review the PowerShell warnings for the affected archive mailbox account(s).</div>' } else { '' })
<div class="note section"><strong>Planning note:</strong> Teams files are included in SharePoint and OneDrive totals. Storage results are a point-in-time estimate, not a complete inventory of every Microsoft 365 data type. Backup providers may bill for retained/versioned data not reflected here. Fast Exchange mode uses Microsoft 365 usage-report data, which can lag live mailbox statistics. Some Exchange Online archive-statistics requests may time out; the script records the outcome and continues.</div>
<div class="footer">Microsoft 365 Tenant Storage Assessment v$ScriptVersion | Copyright &copy; 2026 Abacus Information Systems, Inc. MIT licensed.</div></div></body></html>
"@
    Set-Content -Path $Path -Value $Html -Encoding utf8
}

Write-Banner
$AdminUPN = Read-Host '  Enter the Microsoft 365 administrator UPN'
if ($AdminUPN -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') { throw 'Enter a valid administrator UPN, such as admin@contoso.com.' }
Write-Host ''
Write-WarnLine 'Expect several sign-in prompts when switching tenants or services.'
Write-Host '      Use the same tenant administrator account for Graph, Exchange, and SharePoint.' -ForegroundColor DarkGray
Write-Host ''

Ensure-Module -Name Microsoft.Graph.Authentication -MinimumVersion '2.0.0'
Ensure-Module -Name Microsoft.Graph.Sites -MinimumVersion '2.0.0'
Ensure-Module -Name Microsoft.Graph.Reports -MinimumVersion '2.0.0'
Ensure-Module -Name ExchangeOnlineManagement -MinimumVersion '3.4.0'
Ensure-Module -Name Microsoft.Online.SharePoint.PowerShell -MinimumVersion '16.0.0'

Write-Step 'Signing in to Microsoft Graph and discovering the tenant...'
Connect-MgGraph -Scopes 'Sites.Read.All','Organization.Read.All','Reports.Read.All' -ContextScope Process -NoWelcome
$GraphContext = Get-MgContext
if (-not $GraphContext -or -not $GraphContext.TenantId) { throw 'Microsoft Graph authentication completed without a tenant context.' }
$OrganizationResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,verifiedDomains'
$Organization = @($OrganizationResponse.value)[0]
if (-not $Organization) { throw 'Microsoft Graph did not return tenant organization information.' }
$RootSite = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/sites/root?$select=webUrl'
if (-not $RootSite.webUrl) { throw 'Microsoft Graph did not return the tenant root SharePoint site.' }
$RootUri = [uri]$RootSite.webUrl
$SharePointHost = $RootUri.Host
if ($SharePointHost -notmatch '^(?<Prefix>.+)\.sharepoint\.(?<Suffix>com|us|de|cn)$') { throw "Unexpected SharePoint hostname returned by Graph: $SharePointHost" }
$DetectedAdminUrl = "https://$($Matches.Prefix)-admin.sharepoint.$($Matches.Suffix)"

$TenantDisplayName = Confirm-Value -Label 'Tenant display name' -DetectedValue ([string]$Organization.displayName)
$TenantId = Confirm-Value -Label 'Tenant ID' -DetectedValue ([string]$GraphContext.TenantId)
$SharePointAdminUrl = Confirm-Value -Label 'SharePoint admin URL' -DetectedValue $DetectedAdminUrl
if (-not [uri]::IsWellFormedUriString($SharePointAdminUrl,[System.UriKind]::Absolute)) { throw "The SharePoint admin URL is not valid: $SharePointAdminUrl" }
$SharePointAdminUrl = $SharePointAdminUrl.TrimEnd('/')

Write-Host ''
Write-Step 'Connecting to Exchange Online...'
Connect-ExchangeOnline -UserPrincipalName $AdminUPN -ShowBanner:$false
$AcceptedDomain = Get-AcceptedDomain | Where-Object Default -eq $true | Select-Object -First 1
$PrimaryDomain = if ($AcceptedDomain) { [string]$AcceptedDomain.DomainName } else { ($AdminUPN -split '@',2)[1] }
Write-OK "Exchange Online connected for $PrimaryDomain."

Write-Step "Validating SharePoint admin URL: $SharePointAdminUrl"
try { $null = Test-SPOAdminConnection -AdminUrl $SharePointAdminUrl; Write-OK 'SharePoint admin URL validated successfully.' }
catch {
    Write-WarnLine "Validation failed for $SharePointAdminUrl"
    $ManualAdminUrl = Read-Host '  Enter the correct SharePoint admin URL'
    if ([string]::IsNullOrWhiteSpace($ManualAdminUrl)) { throw }
    $SharePointAdminUrl = $ManualAdminUrl.Trim().TrimEnd('/')
    $null = Test-SPOAdminConnection -AdminUrl $SharePointAdminUrl
    Write-OK 'Manually supplied SharePoint admin URL validated successfully.'
}

$AssessmentDate = Get-Date
$SafeTenantName = ($TenantDisplayName -replace '[^A-Za-z0-9._-]','_').Trim('_')
if ([string]::IsNullOrWhiteSpace($SafeTenantName)) { $SafeTenantName = $TenantId }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$Timestamp = $AssessmentDate.ToString('yyyyMMdd-HHmmss')
$WorkingFolder = Join-Path ([System.IO.Path]::GetTempPath()) "M365Sizing-$Timestamp"
New-Item -ItemType Directory -Path $WorkingFolder -Force | Out-Null

try {
    if ($ExchangeMode -eq 'Fast') {
        try { $ExchangeResult = Get-ExchangeBytesFast -WorkingFolder $WorkingFolder -SkipArchives:$SkipArchives }
        catch {
            Write-WarnLine "Fast Exchange mode failed: $($_.Exception.Message)"
            Write-WarnLine 'Falling back to exact mailbox-by-mailbox measurement.'
            $ExchangeResult = Get-ExchangeBytesExact -SkipArchives:$SkipArchives
        }
    } else { $ExchangeResult = Get-ExchangeBytesExact -SkipArchives:$SkipArchives }
    [decimal]$ExchangeBytes = $ExchangeResult.TotalBytes
    Write-OK "Exchange measurement complete: $(Format-DataSize $ExchangeBytes)"

    Write-Step 'Measuring SharePoint and OneDrive storage...'
    $Sites = @(Get-SPOSite -Limit All -IncludePersonalSite $true)
    [decimal]$SharePointMB = 0; [decimal]$OneDriveMB = 0
    foreach ($Site in $Sites) {
        $UsedMB = if ($null -eq $Site.StorageUsageCurrent) { 0 } else { [decimal]$Site.StorageUsageCurrent }
        if ($Site.Url -match '-my\.sharepoint\.(com|us|de|cn)/personal/') { $OneDriveMB += $UsedMB } else { $SharePointMB += $UsedMB }
    }
    [decimal]$SharePointBytes = $SharePointMB * 1MB
    [decimal]$OneDriveBytes = $OneDriveMB * 1MB
    [decimal]$TotalBytes = $ExchangeBytes + $SharePointBytes + $OneDriveBytes
    [decimal]$TotalGB = $TotalBytes / 1GB
    $MonthlyEstimate = [math]::Round($TotalGB * $BackupPricePerGB,2)
    $AnnualEstimate = [math]::Round($MonthlyEstimate * 12,2)

    $CsvPath = Join-Path $OutputFolder "$SafeTenantName-M365-Storage-Assessment-$Timestamp.csv"
    $HtmlPath = Join-Path $OutputFolder "$SafeTenantName-M365-Storage-Assessment-$Timestamp.html"
    $Summary = @(
        [pscustomobject]@{ Service='Exchange Online'; UsedGB=[math]::Round($ExchangeBytes/1GB,2); Notes=$ExchangeResult.Method }
        [pscustomobject]@{ Service='OneDrive'; UsedGB=[math]::Round($OneDriveBytes/1GB,2); Notes='Personal OneDrive sites' }
        [pscustomobject]@{ Service='SharePoint'; UsedGB=[math]::Round($SharePointBytes/1GB,2); Notes='SharePoint sites, including Teams-connected sites' }
        [pscustomobject]@{ Service='TOTAL'; UsedGB=[math]::Round($TotalGB,2); Notes='Current live-data planning estimate' }
    )
    $Summary | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding utf8
    New-HtmlReport -IncludeBackupEstimate:$IncludeBackupEstimate -Path $HtmlPath -TenantName $TenantDisplayName -TenantId $TenantId -PrimaryDomain $PrimaryDomain -AdminUrl $SharePointAdminUrl -Date $AssessmentDate -ExchangeBytes $ExchangeBytes -PrimaryMailboxBytes $ExchangeResult.PrimaryBytes -ArchiveBytes $ExchangeResult.ArchiveBytes -ArchivesSkipped:$SkipArchives -OneDriveBytes $OneDriveBytes -SharePointBytes $SharePointBytes -TotalBytes $TotalBytes -Monthly $MonthlyEstimate -Annual $AnnualEstimate -Rate $BackupPricePerGB -ExchangeMethod $ExchangeResult.Method -ArchiveStatus $ExchangeResult.ArchiveStatus -ArchiveFailed $ExchangeResult.ArchiveFailed -ArchiveFailureNames @($ExchangeResult.ArchiveFailures | ForEach-Object UserPrincipalName)

    Write-Host ''
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    Write-Host '  |                  ASSESSMENT COMPLETE                       |' -ForegroundColor Green
    Write-Host '  +------------------------------------------------------------+' -ForegroundColor DarkCyan
    Write-Host ("  Tenant             : {0}" -f $TenantDisplayName) -ForegroundColor White
    Write-Host ("  Tenant ID          : {0}" -f $TenantId) -ForegroundColor DarkGray
    Write-Host ("  Primary domain     : {0}" -f $PrimaryDomain) -ForegroundColor DarkGray
    Write-Host ("  SharePoint admin   : {0}" -f $SharePointAdminUrl) -ForegroundColor DarkGray
    Write-Host ("  Exchange method    : {0}" -f $ExchangeResult.Method) -ForegroundColor DarkGray
    Write-Host ("  Archive collection : {0}" -f $ExchangeResult.ArchiveStatus) -ForegroundColor DarkGray
    Write-Host ("  Assessment date    : {0}" -f $AssessmentDate.ToString('yyyy-MM-dd HH:mm')) -ForegroundColor DarkGray
    Write-Host ''
    Write-Host ('  {0,-25}{1,20}' -f 'SERVICE / PLATFORM','USED DATA') -ForegroundColor Cyan
    Write-Host ('  {0}' -f ('-' * 45)) -ForegroundColor DarkGray
    Write-Host ('  {0,-25}{1,20}' -f 'Exchange Online',(Format-DataSize $ExchangeBytes))
    Write-Host ('  {0,-25}{1,20}' -f '  Primary mailboxes',(Format-DataSize $ExchangeResult.PrimaryBytes)) -ForegroundColor DarkGray
    Write-Host ('  {0,-25}{1,20}' -f '  Online archives',$(if ($SkipArchives) { 'Not measured' } else { Format-DataSize $ExchangeResult.ArchiveBytes })) -ForegroundColor DarkGray
    Write-Host ('  {0,-25}{1,20}' -f 'OneDrive',(Format-DataSize $OneDriveBytes))
    Write-Host ('  {0,-25}{1,20}' -f 'SharePoint',(Format-DataSize $SharePointBytes))
    Write-Host ('  {0}' -f ('-' * 45)) -ForegroundColor DarkGray
    Write-Host ('  {0,-25}{1,20}' -f 'TOTAL TENANT DATA',(Format-DataSize $TotalBytes)) -ForegroundColor Green
    Write-Host ''
    if ($IncludeBackupEstimate) {
    Write-Host ('  {0,-25}{1,20}' -f 'Est. monthly cost',('$'+('{0:N2}' -f $MonthlyEstimate))) -ForegroundColor Yellow
    Write-Host ('  {0,-25}{1,20}' -f 'Est. annual cost',('$'+('{0:N2}' -f $AnnualEstimate))) -ForegroundColor Yellow
    Write-Host ('  {0,-25}{1,20}' -f 'Rate used',('$'+('{0:N2}/GB/month' -f $BackupPricePerGB))) -ForegroundColor DarkGray
    }
    Write-Host ''
    Write-Host '  Teams files are included in SharePoint/OneDrive totals.' -ForegroundColor DarkGray
    Write-Host '  Point-in-time estimate; not a complete inventory of all M365 data.' -ForegroundColor DarkGray
    Write-Host "  Summary CSV : $CsvPath" -ForegroundColor Cyan
    Write-Host "  HTML report : $HtmlPath" -ForegroundColor Cyan
    Write-Host ''
    if (-not $NoOpenReport) {
        Write-Step 'Opening the HTML assessment report...'
        Start-Process -FilePath $HtmlPath
    }
}
finally {
    if (Test-Path $WorkingFolder) { Remove-Item $WorkingFolder -Recurse -Force -ErrorAction SilentlyContinue }
    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
}
