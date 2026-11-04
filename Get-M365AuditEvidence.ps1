<#
.SYNOPSIS
  Collects Microsoft 365 / Entra ID / Intune audit evidence and builds one self-contained HTML
  report with a left navigation bar. Optionally also produces two Word documents:
    - MDM Configuration.docx              (Intune compliance policies for Android and iOS/iPadOS)
    - Conditional Access Policies.docx    (all Conditional Access policies)

.DESCRIPTION
  Sections produced:
    1. Executive Summary (narrative, KPI tiles, top priorities, progress since last report, tenant details)
    2. Review These Now: Possible Indications of Compromise (situation table: forwarding, hidden rules, abused apps, sharing, downloads, sign-ins)
    3. Security Controls Summary (one computed status per control)
    4. Remediation Plan (roadmap by phase, playbooks, stable AH- references, View ticket, CSV)
    5. Framework Alignment (every control mapped to NIST CSF 2.0 / SP 800-63B, CIS M365 Benchmark, CIS Controls v8, FFIEC / NCUA)
    6. MDM Compliance Policies        7. Intune Policies           8. Managed Devices
    9. Conditional Access            10. Authentication Methods   11. Enterprise Applications
   12. Application Consent Audit Events                           13. Users
   14. Privileged Users              15. Exchange Online          16. Defender and Email Protection
   17. Microsoft Teams               18. SharePoint and OneDrive  19. Sharing and Access
  README.md documents every section; docs/DEVELOPMENT.md explains the code layout for contributors.

  Notes on "who approved and when":
    Graph does not store an approver on a permission grant. The only record of who consented
    is the Entra directory audit log ("Consent to application" and "Add service principal").
    Retention is 30 days with Entra ID P1/P2 and 7 days on Free. For anything older the report
    shows the date the app was added to the tenant (servicePrincipal.createdDateTime) and
    marks the approver as "Not in audit log retention".

.REQUIREMENTS
  ============================================================================================
  REQUIREMENTS (the script installs the two PowerShell modules for you on first run)
  ============================================================================================
  1. Windows 10/11 or Windows Server with Windows PowerShell 5.1 or PowerShell 7.x.
  2. Internet access to graph.microsoft.com, login.microsoftonline.com, outlook.office365.com
     and the PowerShell Gallery (www.powershellgallery.com) for the one-time module install.
  3. PowerShell modules (auto-installed for the current user if missing, no admin rights needed):
        Microsoft.Graph.Authentication   - sign in to Microsoft Graph and call the REST API
        ExchangeOnlineManagement         - mail forwarding, POP/IMAP/SMTP AUTH checks
     To install them yourself instead:
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
        Install-Module ExchangeOnlineManagement -Scope CurrentUser
  4. Script execution allowed for this session:
        Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
  5. A Microsoft 365 account to sign in with. Sign-in is interactive in a browser window, so
     MFA, Conditional Access, and passwordless methods all work (use -DeviceCode in remote
     sessions where no browser can open on this PC). The account needs:
        - Global Reader (or Global Administrator) for every Graph section, and
        - View-Only Organization Management (or Exchange Administrator) for the Exchange section.
     The FIRST run in a tenant must be done by a Global Administrator (or Privileged Role /
     Cloud Application Administrator) so the delegated read-only Graph scopes can be consented.
     After that, a Global Reader can run it.
     Sign-in uses the "AuditHat Pulse Reports (PowerShell)" app: a multi-tenant public client (no
     secret) with delegated read-only permissions only. An admin consents once per tenant, either
     on the first run or ahead of time at:
        https://login.microsoftonline.com/<tenant>/adminconsent?client_id=ae173bf3-ba47-4ab8-ac54-3ea92fa6c337
     Prefer your own registration? See README "Register your own app" and run with -ClientId
     <app id> -TenantId <tenant id>, or save both in M365AuditEvidence.settings.json next to the script.
  6. Optional: Microsoft Word installed, only if you use -WordDocs to produce the two .docx files.

  The script is READ-ONLY. It requests only *.Read.* Graph scopes and runs only Get-* Exchange
  cmdlets. Nothing in the tenant is changed.

.NOTES
  Version:  2026.10.31  (date-based: year.month.day of the last change)
  Source:   https://github.com/ryanerb4/AuditHat-M365-Audit-Evidence
  See also: https://github.com/ryanerb4/audithat-ad-governance-report (companion report for on-premises AD)
  Author:   Ryan Erb
  Provided: Audit Hat  (https://AuditHat.com)
  Purpose:  Help banks and credit unions understand their Microsoft 365 security posture and risk,
            maintain it over time, and give IT committees, board members, examiners and auditors
            the information they need from one read-only run, without portal screenshots.
  License:  Apache License 2.0 (see LICENSE in the repository).
            Provided as-is, without warranty. Review the output before sharing it outside your
            organization; it contains user, device, and application inventory.
  Privacy:  THIS SCRIPT SENDS NO DATA TO AUDIT HAT OR ANY THIRD PARTY. It runs entirely on the
            computer where you start it, reads from your own Microsoft 365 tenant with the account
            you sign in with, and writes the report only to the output folder on this computer.
            Audit Hat has no server side for this tool and receives nothing. It is a free, internal
            tool provided to help you understand and reduce your Microsoft 365 risk.
  Connector: Microsoft requires every sign-in to identify an application. This script identifies
            itself as the "AuditHat Pulse Reports (PowerShell)" connector, an Entra enterprise
            application published by Audit Hat. Consenting to it in your tenant grants THIS LOCAL
            SCRIPT read-only rights for the signed-in user and nothing else: the connector holds no
            secret, has no application permissions, and cannot be used without a user signing in.
            Prefer your own registration? Use -ClientId (see README, "Register your own app").

.EXAMPLE
  .\Get-M365AuditEvidence.ps1
  Signs in, names the report after the tenant display name and default domain, writes the HTML report.

.EXAMPLE
  .\Get-M365AuditEvidence.ps1 -ClientName "Sample Bank" -WordDocs

.EXAMPLE
  .\Get-M365AuditEvidence.ps1 -TenantId contoso.onmicrosoft.com -OutputFolder C:\Audit\Contoso -AuditDays 30

.EXAMPLE
  .\Get-M365AuditEvidence.ps1 -ClientName "Sample Bank" -ClientId 11111111-2222-3333-4444-555555555555 -TenantId samplebank.onmicrosoft.com
#>
#Requires -Version 5.1
[CmdletBinding()]
param(
    [string]$ClientName,           # optional; defaults to the tenant display name read from the tenant (default domain is shown alongside)
    [string]$TenantId,             # tenant GUID or domain (contoso.onmicrosoft.com); pass it on every run when you audit more than one organization (it selects that organization's saved sign-in); required when -ClientId is a single-tenant app registration
    [string]$ClientId,             # override the default Audit Hat public-client app with your own registration (README "Register your own app")
    [string]$OutputFolder = ".\M365-Audit-Evidence",
    [int]$AuditDays       = 30,
    [int]$TargetScorePercent = 80,   # target Microsoft Secure Score as a percentage of the maximum achievable score
    [switch]$WordDocs,
    [switch]$SkipExchange,         # skip the Exchange Online checks (auto-forwarding, POP/IMAP/SMTP AUTH)
    [switch]$SkipInboxRules,       # skip per-mailbox inbox rule scan (slow on large tenants)
    [switch]$DeviceCode,           # sign in with a device code instead of the browser (for remote sessions or if the browser flow fails)
    [switch]$NoHistory,            # do not read or write the per-tenant history file used for the progress charts and change log
    [switch]$NoOpen,               # do not open the finished report in the default browser
    [switch]$SkipAuditLog,         # skip the unified audit log search (sharing, downloads, mailbox rule changes) used for indications of compromise
    [switch]$TestPermissions,      # sign in, probe every API and permission the report uses, print what is missing, and exit without collecting
    [switch]$SkipShareScan,        # skip the per-item sharing scan of SharePoint sites and OneDrive (section 19 then shows Teams, Exchange and guest access only)
    [switch]$SaveSignIn,           # after a successful browser sign-in, keep the refresh token encrypted (Windows DPAPI, this user, this PC) so later runs and scheduled tasks need no prompt
    [switch]$ForgetSignIn,         # delete the saved sign-in and sign in interactively
    [switch]$NonInteractive,       # for scheduled tasks: use the saved sign-in or stop with exit code 2; never open a browser or wait for a code
    [string]$SignInCachePath       # where the encrypted saved sign-in lives; default %LOCALAPPDATA%\AuditHat\M365AuditEvidence.signin, or M365AuditEvidence-<tenant>.signin with -TenantId
)

$ErrorActionPreference = 'Stop'
$ReportDate = Get-Date
$ScriptVersion = '2026.11.04'   # date-based version: year.month.day of the last change
$RepoUrl = 'https://github.com/ryanerb4/AuditHat-M365-Audit-Evidence'
$AdRepoUrl = 'https://github.com/ryanerb4/audithat-ad-governance-report'   # companion report for on-premises Active Directory

# Optional settings file next to the script so -ClientId / -TenantId do not have to be typed every run:
#   { "ClientId": "<app registration client id>", "TenantId": "<tenant id or domain>" }
#   Optional "PriorityOverrides": { "<regex matched against the finding name>": "Critical|High|Medium|Low" } moves findings
#   between roadmap phases for this institution (Critical = this week, High = 30 days, Medium = 90 days, Low = 12 months).
$settingsPath = Join-Path $PSScriptRoot 'M365AuditEvidence.settings.json'
$script:PriorityOverrides = [ordered]@{}
if (Test-Path $settingsPath) {
    try {
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
        if (-not $ClientId -and $settings.ClientId) { $ClientId = [string]$settings.ClientId }
        if (-not $TenantId -and $settings.TenantId) { $TenantId = [string]$settings.TenantId }
        if ($settings.PriorityOverrides) { foreach ($prop in $settings.PriorityOverrides.PSObject.Properties) { if ([string]$prop.Value -in @('Critical', 'High', 'Medium', 'Low')) { $script:PriorityOverrides[$prop.Name] = [string]$prop.Value } else { Write-Warning "PriorityOverrides: '$($prop.Name)' has an invalid priority '$($prop.Value)' (use Critical, High, Medium or Low)." } } }
        Write-Host "Using settings from $settingsPath" -ForegroundColor DarkGray
    } catch { Write-Warning "Could not read $settingsPath : $($_.Exception.Message)" }
}
# Default app: "AuditHat Pulse Reports (PowerShell)", a multi-tenant PUBLIC client registered in the Audit Hat tenant.
# Public client = no secret, no certificate, nothing to embed or rotate. Every permission is delegated and read-only,
# so the token can never do more than the signed-in person can. Redirect URI http://localhost. Any tenant can use it
# after an admin consents once. Override with -ClientId to use your own registration (README, "Register your own app").
$DefaultGraphClientId = 'ae173bf3-ba47-4ab8-ac54-3ea92fa6c337'
if (-not $ClientId) { $ClientId = $DefaultGraphClientId }
if ($ClientId -ne $DefaultGraphClientId -and -not $TenantId) {
    Write-Verbose "-ClientId was given without -TenantId. If that app registration is single-tenant, sign-in fails with AADSTS50194; add -TenantId <tenant GUID or domain>."
}

# ----------------------------------------------------------------------------------------------
# Connect
# ----------------------------------------------------------------------------------------------
$Scopes = @(
    'Organization.Read.All',
    'User.Read.All',
    'Group.Read.All',
    'Directory.Read.All',
    'Application.Read.All',
    'Policy.Read.All',
    'Policy.Read.ConditionalAccess',
    'DeviceManagementConfiguration.Read.All',
    'DeviceManagementManagedDevices.Read.All',
    'AuditLog.Read.All',
    'UserAuthenticationMethod.Read.All',
    'RoleManagement.Read.Directory',
    'Device.Read.All',
    'DeviceManagementApps.Read.All',
    'Policy.Read.AuthenticationMethod',
    'SecurityEvents.Read.All',         # Microsoft Secure Score (read-only)
    'IdentityRiskyUser.Read.All',      # Entra ID Protection risky users (needs Entra ID P2; indications of compromise)
    'IdentityRiskEvent.Read.All',      # Entra ID Protection risk detections (needs Entra ID P2)
    'Reports.Read.All',                # Teams, SharePoint and OneDrive usage reports (site and team inventories)
    'SharePointTenantSettings.Read.All', # SharePoint and OneDrive tenant sharing settings
    'TeamSettings.Read.All',           # per-team settings (archived, guest settings)
    'Sites.Read.All'                   # sharing links and permissions on SharePoint and OneDrive items the signed-in account can read
)

# --- Prerequisite check and auto-install (current user scope, no admin rights required) -------
function Install-RequiredModule {
    param([Parameter(Mandatory)][string]$Name, [string]$Why)
    if (Get-Module -ListAvailable -Name $Name) { return $true }
    Write-Host "Module '$Name' is not installed ($Why). Installing for the current user..." -ForegroundColor Yellow
    try {
        # TLS 1.2 is required by the PowerShell Gallery on Windows PowerShell 5.1
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }
        if ((Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue).InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        }
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "Installed $Name." -ForegroundColor Green
        return $true
    } catch {
        Write-Warning "Could not install $Name automatically: $($_.Exception.Message)"
        Write-Warning "Install it manually and re-run:  Install-Module $Name -Scope CurrentUser"
        return $false
    }
}

if (-not (Install-RequiredModule -Name 'Microsoft.Graph.Authentication' -Why 'required for all Microsoft Graph data')) {
    throw "Microsoft.Graph.Authentication is required. Stopping."
}
if (-not $SkipExchange) {
    if (-not (Install-RequiredModule -Name 'ExchangeOnlineManagement' -Why 'needed for the Exchange Online forwarding and legacy protocol checks')) {
        Write-Warning "Continuing without the Exchange Online section."
        $SkipExchange = $true
    }
}
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

# --- Sign in to Microsoft Graph (browser sign-in with MFA / Authenticator; device code as fallback) ---
Write-Host ""
Write-Host "Audit Hat - Microsoft 365 Audit Evidence (read-only)  version $ScriptVersion  $RepoUrl" -ForegroundColor Cyan
Write-Host "PRIVACY: this script runs only on this computer. It reads your tenant with the account you sign in with and writes" -ForegroundColor DarkCyan
Write-Host "the report to $OutputFolder. NO DATA IS SENT TO AUDIT HAT OR ANY THIRD PARTY. The 'AuditHat Pulse Reports (PowerShell)'" -ForegroundColor DarkCyan
Write-Host "connector you consent to only grants this local script read-only rights; it has no server side and receives nothing.
$(if ($SaveSignIn -or (Test-Path $SignInCachePath)) { "A saved sign-in (refresh token) is kept encrypted for this Windows user at $SignInCachePath; remove it with -ForgetSignIn." })" -ForegroundColor DarkCyan
Write-Host ""
Write-Host "Sign in to Microsoft 365 with a Global Reader or Global Administrator account. Complete MFA if prompted." -ForegroundColor Cyan
Write-Host "First run only: a Global Administrator must accept the read-only permission consent prompt." -ForegroundColor Cyan
Write-Host ""
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

# Why the script has its own browser sign-in:
# On Windows, Connect-MgGraph routes browser sign-in through the Web Account Manager (WAM) broker. WAM
# silently picks the account already signed in to Windows (often a personal Microsoft account) and fails
# with no visible error. The module's switch to disable WAM is not honored in current builds (verified on
# 2.35.1 and 2.38.1). So the script runs the standard OAuth 2.0 authorization-code flow with PKCE itself:
# it opens https://login.microsoftonline.com in the default browser (work accounts only), the user signs in
# with MFA / Microsoft Authenticator, and the browser returns to a loopback listener on this PC. The token is
# handed to Connect-MgGraph and refreshed automatically during long runs. Only the read-only scopes above
# are requested, through the Audit Hat public-client app (or the -ClientId you supply). Use -DeviceCode
# (or let a failed browser flow fall back) to sign in with a device code instead.

$script:GraphClientId = $ClientId
$script:AuthTenant    = if ($TenantId) { $TenantId } else { 'organizations' }
# One saved sign-in per tenant: with -TenantId the default file is named after the tenant, so auditing a second
# organization never silently reuses the first organization's sign-in.
if (-not $SignInCachePath) {
    $cacheLeaf = if ($TenantId) { "M365AuditEvidence-$(($TenantId -replace '[^A-Za-z0-9.-]', '_').ToLower()).signin" } else { 'M365AuditEvidence.signin' }
    $SignInCachePath = Join-Path $env:LOCALAPPDATA "AuditHat\$cacheLeaf"
}   # 'organizations' excludes personal Microsoft accounts
$script:GraphScopeStr = (@($Scopes | ForEach-Object { "https://graph.microsoft.com/$_" }) + @('openid', 'profile', 'offline_access')) -join ' '
$script:GraphToken    = $null

function ConvertTo-Base64Url { param([byte[]]$Bytes) return ([Convert]::ToBase64String($Bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')) }
function ConvertFrom-JwtPayload {
    param([string]$Jwt)
    $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/')
    switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
    return ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json)
}
function Request-GraphToken {
    # Exchanges an authorization code or a refresh token for an access token and hands it to the Graph module.
    param([hashtable]$Body)
    $Body.client_id = $script:GraphClientId
    $Body.scope     = $script:GraphScopeStr
    $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($script:AuthTenant)/oauth2/v2.0/token" -Body $Body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    $script:GraphToken = [pscustomobject]@{ AccessToken = $r.access_token; RefreshToken = $r.refresh_token; ExpiresAt = (Get-Date).AddSeconds([int]$r.expires_in) }
    Connect-MgGraph -AccessToken (ConvertTo-SecureString $r.access_token -AsPlainText -Force) -NoWelcome -ErrorAction Stop | Out-Null
}
$script:ConsentNeeded = [ordered]@{}   # service -> scopes the connector has not been approved for in this tenant; printed with approval links at the end of the run
function Get-TokenScope {
    # The delegated permissions (scp claim) inside an access token. Microsoft issues a token even when the tenant has not
    # approved the permission for that service; such a token carries no rights and the service answers "UnAuthorized".
    param([string]$Jwt)
    try {
        $p = $Jwt.Split('.')[1].Replace('-', '+').Replace('_', '/'); switch ($p.Length % 4) { 2 { $p += '==' } 3 { $p += '=' } }
        return [string](([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p)) | ConvertFrom-Json).scp)
    } catch { return $null }
}
function Get-ConsentLink {
    # One-time admin approval of exactly these delegated permissions for the connector in this tenant.
    param([string]$Scopes)
    "https://login.microsoftonline.com/$($ctx.TenantId)/v2.0/adminconsent?client_id=$($script:GraphClientId)&redirect_uri=http://localhost&scope=$($Scopes -replace ' ', '%20')"
}
function Request-ResourceToken {
    # Uses the refresh token from the browser sign-in to get a delegated token for another Microsoft API (same app, same user).
    param([string]$ScopeStr, [string]$RequireScope)
    if (-not $script:GraphToken -or -not $script:GraphToken.RefreshToken) { return $null }
    $body = @{ grant_type = 'refresh_token'; refresh_token = $script:GraphToken.RefreshToken; client_id = $script:GraphClientId; scope = $ScopeStr }
    $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($script:AuthTenant)/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    if ($r.refresh_token) { $script:GraphToken.RefreshToken = $r.refresh_token }
    if ($RequireScope) {
        $scp = Get-TokenScope $r.access_token
        if ($null -ne $scp -and (" $scp " -notlike "* $RequireScope *")) { throw "The connector's $RequireScope permission has not been approved (consent) in this tenant, so the token carries no rights for this service." }
    }
    return $r.access_token
}
function Save-SignInCache {
    # Refresh token encrypted with Windows DPAPI for the current Windows user on this computer. Nobody else, and no other PC, can decrypt it.
    if (-not $script:GraphToken -or -not $script:GraphToken.RefreshToken) { return }
    try {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $payload = @{ RefreshToken = $script:GraphToken.RefreshToken; Account = $script:SignInAccount; TenantId = $script:AuthTenant; ClientId = $script:GraphClientId; Saved = (Get-Date).ToString('yyyy-MM-dd HH:mm') } | ConvertTo-Json -Compress
        $bytes = [System.Security.Cryptography.ProtectedData]::Protect([Text.Encoding]::UTF8.GetBytes($payload), $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        $dir = Split-Path $SignInCachePath -Parent; if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        [IO.File]::WriteAllBytes($SignInCachePath, $bytes)
        $script:SignInCacheSaved = $true
    } catch { Write-Warning "Could not save the sign-in: $($_.Exception.Message)" }
}
function Read-SignInCache {
    if (-not (Test-Path $SignInCachePath)) { return $null }
    try {
        Add-Type -AssemblyName System.Security -ErrorAction SilentlyContinue
        $bytes = [System.Security.Cryptography.ProtectedData]::Unprotect([IO.File]::ReadAllBytes($SignInCachePath), $null, [System.Security.Cryptography.DataProtectionScope]::CurrentUser)
        return ([Text.Encoding]::UTF8.GetString($bytes) | ConvertFrom-Json)
    } catch { Write-Warning "The saved sign-in at $SignInCachePath could not be read (it belongs to another Windows user or PC, or is damaged): $($_.Exception.Message)"; return $null }
}
function Remove-SignInCache { if (Test-Path $SignInCachePath) { Remove-Item $SignInCachePath -Force -ErrorAction SilentlyContinue; Write-Host "Saved sign-in removed from $SignInCachePath" -ForegroundColor DarkGray } }
function Confirm-GraphToken {
    # Refreshes the Graph token when it is within 5 minutes of expiry so long collections do not fail mid-run.
    if ($script:GraphToken -and $script:GraphToken.RefreshToken -and (Get-Date) -gt $script:GraphToken.ExpiresAt.AddMinutes(-5)) {
        try { Request-GraphToken -Body @{ grant_type = 'refresh_token'; refresh_token = $script:GraphToken.RefreshToken } }
        catch { Write-Warning "Could not refresh the Graph token: $($_.Exception.Message)" }
    }
}
function Connect-GraphBrowser {
    # Authorization-code flow with PKCE through the default browser and a loopback listener. Returns $true on success.
    $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
    $vb = New-Object byte[] 32; $rng.GetBytes($vb); $verifier = ConvertTo-Base64Url $vb
    $sb = New-Object byte[] 16; $rng.GetBytes($sb); $state    = ConvertTo-Base64Url $sb
    $challenge = ConvertTo-Base64Url ([Security.Cryptography.SHA256]::Create().ComputeHash([Text.Encoding]::ASCII.GetBytes($verifier)))

    $listener = $null; $port = $null
    foreach ($attempt in 1..10) {
        $p = Get-Random -Minimum 49152 -Maximum 65000
        $l = New-Object System.Net.HttpListener
        $l.Prefixes.Add("http://localhost:$p/")
        try { $l.Start(); $listener = $l; $port = $p; break } catch { $l.Close() }
    }
    if (-not $listener) { Write-Warning "Could not open a local port for the browser sign-in."; return $false }

    $redirect = "http://localhost:$port/"
    $authUrl  = "https://login.microsoftonline.com/$($script:AuthTenant)/oauth2/v2.0/authorize" +
                "?client_id=$($script:GraphClientId)&response_type=code&response_mode=query" +
                "&redirect_uri=$([Uri]::EscapeDataString($redirect))&scope=$([Uri]::EscapeDataString($script:GraphScopeStr))" +
                "&state=$state&code_challenge=$challenge&code_challenge_method=S256&prompt=select_account"

    Write-Host "BROWSER SIGN-IN" -ForegroundColor Yellow
    Write-Host "  A Microsoft sign-in page is opening in your default browser. Pick or type your work admin account," -ForegroundColor Yellow
    Write-Host "  approve the Microsoft Authenticator / MFA prompt, accept the permissions page if it appears, then return here." -ForegroundColor Yellow
    Write-Host "  If no page opens within a few seconds, copy this address into a browser on this PC:" -ForegroundColor Yellow
    Write-Host "  $authUrl" -ForegroundColor DarkGray
    Write-Host ""
    try { Start-Process $authUrl -ErrorAction Stop } catch { Write-Verbose "Could not start the browser automatically." }

    try {
        $task = $listener.GetContextAsync()
        $deadline = (Get-Date).AddMinutes(5)
        while (-not $task.Wait(1000)) {
            if ((Get-Date) -gt $deadline) { Write-Warning "No sign-in response was received within 5 minutes."; return $false }
        }
        $http = $task.Result
        $q = @{}
        foreach ($kv in $http.Request.Url.Query.TrimStart('?').Split('&')) {
            $i = $kv.IndexOf('=')
            if ($i -gt 0) { $q[[Uri]::UnescapeDataString($kv.Substring(0, $i))] = [Uri]::UnescapeDataString($kv.Substring($i + 1).Replace('+', ' ')) }
        }
        $ok  = ($q['code'] -and $q['state'] -eq $state)
        $msg = if ($ok) { 'Sign-in complete. You can close this tab and return to the PowerShell window.' } else { "Sign-in failed: $($q['error']) $($q['error_description'])" }
        $bytes = [Text.Encoding]::UTF8.GetBytes("<html><head><title>M365 Audit Evidence</title></head><body style='font-family:Segoe UI,Arial;margin:40px'><h2>Audit Hat - M365 Audit Evidence</h2><p>$([System.Net.WebUtility]::HtmlEncode($msg))</p></body></html>")
        $http.Response.ContentType = 'text/html; charset=utf-8'
        $http.Response.ContentLength64 = $bytes.Length
        $http.Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $http.Response.OutputStream.Close()
        if (-not $ok) { Write-Warning $msg; return $false }
        Request-GraphToken -Body @{ grant_type = 'authorization_code'; code = $q['code']; redirect_uri = $redirect; code_verifier = $verifier }
        return $true
    } catch {
        Write-Warning "Browser sign-in failed: $($_.Exception.Message)"
        return $false
    } finally { $listener.Stop(); $listener.Close() }
}

$signedIn = $false; $script:SignInFromCache = $false; $script:SignInCacheSaved = $false; $script:SignInAccount = $null
if ($ForgetSignIn) { Remove-SignInCache }
if (-not $DeviceCode) {
    $cached = Read-SignInCache
    if ($cached -and $cached.RefreshToken) {
        if ([string]$cached.ClientId -ne [string]$script:GraphClientId) { Write-Warning 'The saved sign-in was made with a different app registration; signing in again.' }
        elseif ($TenantId -and $cached.TenantId -and [string]$cached.TenantId -ne 'organizations' -and [string]$cached.TenantId -ne [string]$TenantId) { Write-Warning "The saved sign-in at $SignInCachePath is for tenant $($cached.TenantId), not $TenantId; signing in again." }
        else {
            try {
                Request-GraphToken -Body @{ grant_type = 'refresh_token'; refresh_token = [string]$cached.RefreshToken }
                $signedIn = $true; $script:SignInFromCache = $true; $script:SignInAccount = [string]$cached.Account
                Write-Host "Signed in silently as $($cached.Account) with the saved sign-in from $($cached.Saved) (no prompt). Run with -ForgetSignIn to remove it." -ForegroundColor Green
            } catch {
                $m = $_.Exception.Message; try { $b = $_.ErrorDetails.Message; if ($b) { $j = $b | ConvertFrom-Json; if ($j.error_description) { $m = ([string]$j.error_description -split "`r?`n")[0] } } } catch { Write-Verbose 'Token error body was not JSON.' }
                Write-Warning "The saved sign-in could not be used ($m). It expires after 90 days without use, and a password change, revoked sessions or a Conditional Access sign-in frequency policy also invalidate it."
                Remove-SignInCache
            }
        }
    }
}
if (-not $signedIn -and $NonInteractive) {
    Write-Error "No usable saved sign-in and -NonInteractive is set, so no browser or device-code prompt was shown. Run the script once interactively with -SaveSignIn as the same Windows user, then schedule it again."
    exit 2
}
if (-not $signedIn -and -not $DeviceCode) { $signedIn = Connect-GraphBrowser }
if (-not $signedIn) {
    Write-Host ""
    Write-Host "DEVICE CODE SIGN-IN" -ForegroundColor Yellow
    Write-Host "  1. A browser tab is opening at https://microsoft.com/devicelogin (open it yourself if it does not appear)." -ForegroundColor Yellow
    Write-Host "  2. Type the code shown on the next line into that page." -ForegroundColor Yellow
    Write-Host "  3. Sign in with your Global Reader / Global Administrator account and complete MFA." -ForegroundColor Yellow
    Write-Host "  4. Come back to this window; the script continues automatically." -ForegroundColor Yellow
    Write-Host ""
    try { Start-Process 'https://microsoft.com/devicelogin' -ErrorAction Stop } catch { Write-Verbose "Could not open the browser automatically." }
    $connectParams = @{ Scopes = $Scopes; NoWelcome = $true; UseDeviceCode = $true; ClientId = $ClientId }
    if ($TenantId) { $connectParams.TenantId = $TenantId }
    Connect-MgGraph @connectParams -ErrorAction Stop | Out-Null
}

# Build one context object that works for both sign-in methods (Get-MgContext has no account details for a user-provided token).
$mg = Get-MgContext
if ($script:GraphToken) {
    $claims  = ConvertFrom-JwtPayload $script:GraphToken.AccessToken
    $account = @($claims.upn, $claims.preferred_username, $claims.unique_name, $claims.email) | Where-Object { $_ } | Select-Object -First 1
    $ctx = [pscustomobject]@{ Account = $account; TenantId = $claims.tid; Scopes = @([string]$claims.scp -split ' ' | Where-Object { $_ }) }
} else {
    $ctx = [pscustomobject]@{ Account = $mg.Account; TenantId = $mg.TenantId; Scopes = @($mg.Scopes) }
}
if (-not $mg -or -not $ctx.Account) { throw "Sign-in to Microsoft Graph did not complete. Stopping." }
Write-Host "Connected to tenant $($ctx.TenantId) as $($ctx.Account)" -ForegroundColor Green
if (-not $script:SignInAccount) { $script:SignInAccount = [string]$ctx.Account }
if ($SaveSignIn -or $script:SignInFromCache) {
    if ($script:GraphToken -and $script:GraphToken.RefreshToken) { Save-SignInCache; if ($SaveSignIn -and -not $script:SignInFromCache) { Write-Host "Sign-in saved for $($ctx.Account), encrypted for the Windows user $env:USERNAME on this PC at $SignInCachePath. Later runs (and a scheduled task running as this Windows user) will not prompt." -ForegroundColor Green } }
    elseif ($SaveSignIn) { Write-Warning 'The sign-in cannot be saved after a device-code sign-in (no refresh token is available to the script). Use the browser sign-in.' }
}
$missingScopes = @($Scopes | Where-Object { $ctx.Scopes -notcontains $_ })
if ($missingScopes.Count) {
    Write-Warning "The following Graph permissions were not granted; related sections will show as not available: $($missingScopes -join ', ')$(if ($script:SignInFromCache) { '. The saved sign-in predates these permissions: run once with -ForgetSignIn -SaveSignIn to consent and save again.' })"
    Write-Warning "A Global Administrator must run the script once and accept the consent prompt to grant them."
}

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$OutputFolder = (Resolve-Path $OutputFolder).Path

# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Get-GraphAll {
    param([Parameter(Mandatory)][string]$Uri, [switch]$NoPaging, [int]$MaxItems = 0)
    $items = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        Confirm-GraphToken
        $r = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -Headers @{ ConsistencyLevel = 'eventual' }
        if ($null -ne $r.PSObject.Properties['value']) {
            foreach ($v in $r.value) { $items.Add($v) }
        } else {
            $items.Add($r)
        }
        $next = if ($NoPaging -or ($MaxItems -gt 0 -and $items.Count -ge $MaxItems)) { $null } else { $r.'@odata.nextLink' }
    }
    return $items
}

$script:GraphFailures = @{}   # What -> reason, so sections can say "could not be read" instead of "none exist"
function Get-GraphErrorText {
    param($ErrorRecord)
    $msg = $ErrorRecord.Exception.Message
    $code = $null
    try {
        $raw = $ErrorRecord.ErrorDetails.Message
        if ($raw) { $j = $raw | ConvertFrom-Json; if ($j.error.message) { $msg = $j.error.message; $code = $j.error.code } }
    } catch { Write-Verbose "Graph error body was not JSON." }
    $hint = ''
    if ($msg -match '(?i)premium|license|licence|not a (B2C|Premium)|AadPremiumLicenseRequired' -or $code -match '(?i)license') { $hint = ' This feature needs an Entra ID P1/P2 license.' }
    elseif ($ErrorRecord.Exception.Message -match 'BadRequest' -and $msg -match '(?i)intune|device ?management|tenant.*not.*(enabled|found)|No subscription') { $hint = ' Intune does not appear to be licensed or enabled in this tenant.' }
    elseif ($ErrorRecord.Exception.Message -match 'Forbidden') { $hint = ' The signed-in account or the app lacks the permission or the license for this data.' }
    $status = if ($ErrorRecord.Exception.Message -match '(BadRequest|Forbidden|NotFound|Unauthorized|TooManyRequests|InternalServerError|ServiceUnavailable)') { "$($Matches[1]): " } else { '' }
    return "$status$msg$(if ($code) { " [$code]" })$hint"
}
function Get-GraphAllSafe {
    param([string]$Uri, [string]$What, [switch]$NoPaging, [int]$MaxItems = 0)
    try { return (Get-GraphAll -Uri $Uri -NoPaging:$NoPaging -MaxItems $MaxItems) }
    catch {
        $reason = Get-GraphErrorText $_
        # Graph returns a bare 400/403 for unlicensed features; explain from the tenant's own subscriptions when we can.
        if ($Uri -match 'deviceManagement|deviceAppManagement' -and $script:HasIntune -eq $false) { $reason += ' Intune is not licensed in this tenant (no Intune service plan in the subscribed SKUs).' }
        elseif ($Uri -match 'userRegistrationDetails|signInActivity' -and $script:HasEntraP1 -eq $false) { $reason += ' Requires Entra ID P1 or P2, which is not licensed in this tenant.' }
        elseif ($Uri -match 'roleEligibilitySchedule|roleAssignmentSchedule' -and $script:HasEntraP2 -eq $false) { $reason += ' Privileged Identity Management requires Entra ID P2, which is not licensed in this tenant.' }
        elseif ($Uri -match 'auditLogs/signIns' -and $script:HasEntraP1 -eq $false) { $reason += ' The sign-in log requires Entra ID P1 or P2, which is not licensed in this tenant.' }
        elseif ($Uri -match 'identityProtection' -and $script:HasEntraP2 -eq $false) { $reason += ' Entra ID Protection requires Entra ID P2, which is not licensed in this tenant.' }
        elseif ($Uri -match 'admin/sharepoint/settings' -and $_.Exception.Message -match 'Forbidden') { $reason += ' SharePointTenantSettings.Read.All has not been consented, or the account is not a SharePoint or Global administrator / Global Reader.' }
        elseif ($Uri -match '/reports/' -and $_.Exception.Message -match 'Forbidden') { $reason += ' Reports.Read.All has not been consented, or the account lacks the Reports Reader / Global Reader role.' }
        $script:GraphFailures[$What] = $reason
        Write-Warning "Could not retrieve $What : $reason"
        return @()
    }
}
function Get-GraphFailure { param([string]$What) if ($script:GraphFailures.ContainsKey($What)) { return $script:GraphFailures[$What] } ; return $null }
function Get-GraphReportCsv {
    # Microsoft 365 usage reports (Teams, SharePoint, OneDrive) come back as CSV behind a redirect; download and parse.
    param([string]$Report, [string]$What)
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("m365audit-" + [guid]::NewGuid().ToString('N') + ".csv")
    try {
        Confirm-GraphToken
        Invoke-MgGraphRequest -Method GET -Uri "v1.0/reports/$Report" -OutputFilePath $tmp -ErrorAction Stop
        $rows = @(Import-Csv -Path $tmp -ErrorAction Stop)
        return $rows
    } catch {
        $reason = Get-GraphErrorText $_
        if ($_.Exception.Message -match 'Forbidden') { $reason += ' Reports.Read.All has not been consented, or the account lacks the Reports Reader / Global Reader role.' }
        $script:GraphFailures[$What] = $reason
        Write-Warning "Could not retrieve $What : $reason"
        return @()
    } finally { if (Test-Path $tmp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue } }
}
function Test-ConcealedReportNames {
    # Microsoft 365 usage reports hide user, group and site names (hashes instead) when the tenant's report privacy setting is on.
    param($Rows, [string]$Column)
    $v = @($Rows | Select-Object -First 5 | ForEach-Object { [string]$_.$Column } | Where-Object { $_ })
    if (-not $v.Count) { return $false }
    return (@($v | Where-Object { $_ -match '^[0-9A-Fa-f]{20,}$' }).Count -eq $v.Count)
}

function HtmlEnc { param($s) if ($null -eq $s) { return '' } ; return [System.Net.WebUtility]::HtmlEncode([string]$s) }

function Format-DateText { param($d) if (-not $d) { return '' } ; try { return ([datetime]$d).ToLocalTime().ToString('yyyy-MM-dd HH:mm') } catch { return [string]$d } }

function ConvertFrom-HtmlText { param([string]$s) if (-not $s) { return '' } ; $x = $s -replace '<br\s*/?>', ' ' -replace '</(p|li|ul|ol|div|h\d)>', ' ' -replace '<[^>]+>', ' '; $x = [System.Net.WebUtility]::HtmlDecode($x); return (($x -replace '[\u00A0\s]+', ' ').Trim()) }

function Join-List { param($arr, $sep = ', ') if ($null -eq $arr) { return '' } ; return (@($arr) -join $sep) }

$script:NameCache = @{}
function Resolve-ObjectNames {
    param([string[]]$Ids)
    $need = @($Ids | Where-Object { $_ -and -not $script:NameCache.ContainsKey($_) } | Select-Object -Unique)
    for ($i = 0; $i -lt $need.Count; $i += 1000) {
        $chunk = $need[$i..([Math]::Min($i + 999, $need.Count - 1))]
        try {
            $body = @{ ids = @($chunk); types = @('user', 'group', 'servicePrincipal', 'device', 'directoryRole', 'application') } | ConvertTo-Json
            Confirm-GraphToken
            $r = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/directoryObjects/getByIds' -Body $body -ContentType 'application/json' -OutputType PSObject
            foreach ($o in $r.value) {
                $label = $o.displayName
                if ($o.userPrincipalName) { $label = "$($o.displayName) ($($o.userPrincipalName))" }
                $script:NameCache[$o.id] = $label
            }
        } catch { Write-Warning "getByIds failed: $($_.Exception.Message)" }
        foreach ($id in $chunk) { if (-not $script:NameCache.ContainsKey($id)) { $script:NameCache[$id] = "$id (deleted or not found)" } }
    }
    return ($Ids | ForEach-Object { if ($script:NameCache.ContainsKey($_)) { $script:NameCache[$_] } else { $_ } })
}

$script:RoleTemplates = @{}
function Resolve-RoleTemplateNames {
    param([string[]]$Ids)
    if ($script:RoleTemplates.Count -eq 0) {
        foreach ($t in (Get-GraphAllSafe 'v1.0/directoryRoleTemplates' 'directory role templates')) { $script:RoleTemplates[$t.id] = $t.displayName }
    }
    return ($Ids | ForEach-Object { if ($script:RoleTemplates.ContainsKey($_)) { $script:RoleTemplates[$_] } else { $_ } })
}

function Resolve-CaTargets {
    param($list, [switch]$Roles)
    if ($null -eq $list -or @($list).Count -eq 0) { return 'None' }
    $out = foreach ($x in @($list)) {
        switch ($x) {
            'All'                    { 'All'; continue }
            'None'                   { 'None'; continue }
            'GuestsOrExternalUsers'  { 'Guests or external users'; continue }
            'Office365'              { 'Office 365'; continue }
            'MicrosoftAdminPortals'  { 'Microsoft Admin Portals'; continue }
            'AllTrusted'             { 'All trusted locations'; continue }
            default {
                if ($x -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$') {
                    if ($Roles) { Resolve-RoleTemplateNames -Ids @($x) } else { Resolve-ObjectNames -Ids @($x) }
                } else { $x }
            }
        }
    }
    return (Join-List $out '; ')
}

# ----------------------------------------------------------------------------------------------
# Configuration issue catalog. Every flagged value in the report carries an (i) icon; clicking it shows the
# title, why it matters to a bank or credit union, and the recommended fix. Keyed by a short issue id.
# ----------------------------------------------------------------------------------------------
if ($TestPermissions) {
    Write-Host "`nPERMISSION TEST: each line is one capability the report uses. Nothing is collected or written." -ForegroundColor Cyan
    $script:permResults = @()
    function Add-PermResult { param([string]$Capability, [string]$Result, [string]$Needs) $script:permResults += [pscustomobject]@{ Capability = $Capability; Result = $Result; 'Needs / fix' = $Needs } }
    $missing = @($Scopes | Where-Object { $ctx.Scopes -notcontains $_ })
    Add-PermResult 'Graph delegated scopes granted' $(if ($missing.Count) { "MISSING $($missing.Count)" } else { 'OK' }) $(if ($missing.Count) { "Add to the connector and consent: $($missing -join ', ')" } else { "All $($Scopes.Count) requested scopes are granted" })
    function Test-GraphGet { param([string]$Cap, [string]$Uri, [string]$Needs)
        try { Confirm-GraphToken; Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop | Out-Null; Add-PermResult $Cap 'OK' '' }
        catch { $t = Get-GraphErrorText $_; Add-PermResult $Cap $(if ($t -match 'Forbidden|Unauthorized') { 'NO PERMISSION' } elseif ($t -match '(?i)licen[cs]e|premium|BadRequest') { 'NOT LICENSED' } else { 'FAILED' }) "$Needs. $t" }
    }
    Test-GraphGet 'Tenant, users, groups, roles'            'v1.0/organization?$select=id'                                       'Organization.Read.All, Directory.Read.All'
    Test-GraphGet 'Conditional Access policies'              'v1.0/identity/conditionalAccess/policies?$top=1'                    'Policy.Read.ConditionalAccess (Entra ID P1 for policies to exist)'
    Test-GraphGet 'Authentication methods policy'            'v1.0/policies/authenticationMethodsPolicy'                           'Policy.Read.AuthenticationMethod'
    Test-GraphGet 'Per-user authentication methods'          'v1.0/me/authentication/methods'                                     'UserAuthenticationMethod.Read.All (used for MFA and passkey facts when the P1 report is unavailable)'
    Test-GraphGet 'MFA registration report (Entra ID P1)'    'v1.0/reports/authenticationMethods/userRegistrationDetails?$top=1'  'Entra ID P1; without it the report falls back to per-user methods'
    Test-GraphGet 'Directory audit log'                      'v1.0/auditLogs/directoryAudits?$top=1'                              'AuditLog.Read.All'
    Test-GraphGet 'Sign-in log (Entra ID P1)'                'v1.0/auditLogs/signIns?$top=1'                                      'AuditLog.Read.All plus Entra ID P1'
    Test-GraphGet 'Entra ID Protection (Entra ID P2)'        'v1.0/identityProtection/riskyUsers?$top=1'                          'IdentityRiskyUser.Read.All plus Entra ID P2'
    Test-GraphGet 'Intune compliance policies'               'v1.0/deviceManagement/deviceCompliancePolicies?$top=1'              'DeviceManagementConfiguration.Read.All plus an Intune licence'
    Test-GraphGet 'Microsoft Secure Score'                   'v1.0/security/secureScores?$top=1'                                  'SecurityEvents.Read.All'
    Test-GraphGet 'Enterprise applications and consents'     'v1.0/servicePrincipals?$top=1'                                      'Application.Read.All'
    Test-GraphGet 'Teams (groups) and group settings'        'v1.0/groupSettings'                                                 'Group.Read.All, Directory.Read.All'
    Test-GraphGet 'SharePoint tenant sharing settings'       'v1.0/admin/sharepoint/settings'                                     'SharePointTenantSettings.Read.All plus SharePoint Administrator or Global Reader'
    $null = Get-GraphReportCsv "getOffice365ActiveUserCounts(period='D7')" 'permission test report'
    Add-PermResult 'Usage reports (Teams, SharePoint, OneDrive)' $(if (Get-GraphFailure 'permission test report') { 'NO PERMISSION' } else { 'OK' }) $(if (Get-GraphFailure 'permission test report') { "Reports.Read.All plus Reports Reader or Global Reader. $(Get-GraphFailure 'permission test report')" } else { '' })
    function Test-ResourceToken { param([string]$Cap, [string]$ScopeStr, [string]$Needs, [string]$RequireScope)
        if (-not $script:GraphToken -or -not $script:GraphToken.RefreshToken) { Add-PermResult $Cap 'SKIPPED' 'Device-code sign-in has no refresh token; run the browser sign-in to test this'; return $null }
        try { $tok = Request-ResourceToken $ScopeStr -RequireScope $RequireScope; Add-PermResult $Cap 'OK' ''; return $tok }
        catch { $m = $_.Exception.Message; try { $b = $_.ErrorDetails.Message; if ($b) { $j = $b | ConvertFrom-Json; if ($j.error_description) { $m = ([string]$j.error_description -split "`r?`n")[0] } } } catch { Write-Verbose 'Token error body was not JSON.' }; Add-PermResult $Cap $(if ($m -match 'AADSTS65001|AADSTS650|consent') { 'NOT CONSENTED' } else { 'FAILED' }) "$Needs. $m"; return $null }
    }
    $exoTok   = Test-ResourceToken 'Exchange Online single sign-on'          'https://outlook.office365.com/.default' 'Office 365 Exchange Online > Exchange.Manage (delegated); without it a second sign-in prompt appears' -RequireScope 'Exchange.Manage'
    $null     = Test-ResourceToken -RequireScope 'user_impersonation' 'Teams admin API (tenant Teams policies)' '48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default' 'Skype and Teams Tenant Admin API > user_impersonation (delegated); optional, also needs the MicrosoftTeams module'
    $mdeTok   = Test-ResourceToken 'Defender for Endpoint API consent'      'https://api.securitycenter.microsoft.com/Machine.Read https://api.securitycenter.microsoft.com/Software.Read https://api.securitycenter.microsoft.com/Vulnerability.Read' 'WindowsDefenderATP > Machine.Read, Software.Read, Vulnerability.Read (delegated)'
    if ($mdeTok) {
        try { Invoke-RestMethod -Method Get -Uri 'https://api.securitycenter.microsoft.com/api/machines?$top=1' -Headers @{ Authorization = "Bearer $mdeTok" } -ErrorAction Stop | Out-Null; Add-PermResult 'Defender for Endpoint licence' 'OK' '' }
        catch { $m = $_.Exception.Message; try { $b = $_.ErrorDetails.Message; if ($b) { $j = $b | ConvertFrom-Json; if ($j.error.message) { $m = $j.error.message } } } catch { Write-Verbose 'Defender error body was not JSON.' }; Add-PermResult 'Defender for Endpoint licence' $(if ($m -match '(?i)licen[cs]e') { 'NOT LICENSED' } else { 'FAILED' }) "Defender for Business or Defender for Endpoint P1/P2. $m" }
    }
    Add-PermResult 'ExchangeOnlineManagement module' $(if (Get-Module -ListAvailable ExchangeOnlineManagement) { 'OK' } else { 'MISSING' }) $(if (Get-Module -ListAvailable ExchangeOnlineManagement) { '' } else { 'Install-Module ExchangeOnlineManagement -Scope CurrentUser (a normal run installs it)' })
    Add-PermResult 'MicrosoftTeams module (optional)'   $(if (Get-Module -ListAvailable MicrosoftTeams) { 'OK' } else { 'MISSING' }) $(if (Get-Module -ListAvailable MicrosoftTeams) { '' } else { 'Install-Module MicrosoftTeams -Scope CurrentUser to include tenant-wide Teams policies' })
    if ($exoTok -and (Get-Module -ListAvailable ExchangeOnlineManagement)) {
        try {
            Import-Module ExchangeOnlineManagement -ErrorAction Stop
            Connect-ExchangeOnline -AccessToken $exoTok -UserPrincipalName $ctx.Account -ShowBanner:$false -ErrorAction Stop
            Add-PermResult 'Exchange Online connection (token)' 'OK' ''
            $cfg = Get-AdminAuditLogConfig
            Add-PermResult 'Unified audit log turned on' $(if ($cfg.UnifiedAuditLogIngestionEnabled) { 'OK' } else { 'TURNED OFF' }) $(if ($cfg.UnifiedAuditLogIngestionEnabled) { '' } else { 'Purview portal > Audit > Start recording user and admin activity (or Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true)' })
            try { Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -ResultSize 1 -ErrorAction Stop | Out-Null; Add-PermResult 'Audit log search role' 'OK' '' }
            catch { Add-PermResult 'Audit log search role' 'NO PERMISSION' "View-Only Audit Logs or Audit Logs role in Exchange Online. $($_.Exception.Message)" }
            if ($script:HasDefenderP1 -or (Get-Command Get-SafeLinksDetailReport -ErrorAction SilentlyContinue)) { try { Get-SafeLinksDetailReport -StartDate (Get-Date).AddDays(-1) -EndDate (Get-Date) -PageSize 1 -ErrorAction Stop | Out-Null; Add-PermResult 'Safe Links click report' 'OK' '' } catch { Add-PermResult 'Safe Links click report' 'NOT LICENSED' "Defender for Office 365 Plan 1 or 2. $($_.Exception.Message)" } }
        } catch { Add-PermResult 'Exchange Online connection (token)' 'FAILED' $_.Exception.Message }
        finally { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'Exchange Online session was already closed.' } }
    }
    $script:permResults | Format-Table -AutoSize -Wrap | Out-String -Width 230 | Write-Host
    $bad = @($script:permResults | Where-Object { $_.Result -ne 'OK' })
    if ($bad.Count) { Write-Host "$($bad.Count) item(s) need attention; see 'Needs / fix'. Items marked NOT LICENSED are licence gaps, not permission problems. Admin consent for the connector: https://login.microsoftonline.com/$($ctx.TenantId)/adminconsent?client_id=$($script:GraphClientId)" -ForegroundColor Yellow }
    else { Write-Host 'Everything the report uses is available.' -ForegroundColor Green }
    return
}

$IssueCatalog = @{
    'mfa-not-enforced'     = @{ Title = 'MFA is not enforced for every user';                 Why = 'Passwords alone are the most common way attackers get into Microsoft 365. Without enforced MFA a single phished or reused password gives full mailbox and file access, and examiners expect MFA for all users.'; Fix = 'Create an enabled Conditional Access policy that requires MFA (or a phishing-resistant authentication strength) for all users and all cloud apps with a monitored break-glass exclusion, or turn on security defaults if Entra ID P1 is not licensed.' }
    'mfa-not-registered'   = @{ Title = 'Some users have no MFA method registered';           Why = 'A user who has never registered a method cannot be protected by an MFA policy and is the account an attacker will target first.'; Fix = 'Run a registration campaign, require Microsoft Authenticator or a FIDO2 key, and block sign-in for accounts that have not registered after the deadline.' }
    'legacy-auth'          = @{ Title = 'Legacy authentication is not blocked';               Why = 'Legacy protocols (POP, IMAP, SMTP AUTH, Exchange ActiveSync basic) cannot do MFA, so they let attackers bypass it with just a password. Most password-spray attacks use them.'; Fix = 'Add a Conditional Access policy that blocks legacy clients for all users, disable SMTP AUTH tenant-wide, and turn off POP and IMAP on mailboxes that do not need them.' }
    'device-ca'            = @{ Title = 'Access is not limited to managed devices';            Why = 'With valid credentials, any personal or unknown computer can open company mail and files. Requiring a compliant or domain-joined device stops stolen credentials being used from an attacker''s machine and keeps data on devices you control.'; Fix = 'Enroll company devices in Intune, then add a Conditional Access policy requiring a compliant or hybrid-joined device for Microsoft 365 apps. Start in report-only mode.' }
    'intune-coverage'      = @{ Title = 'Windows computers are not all managed by Intune';    Why = 'Unmanaged computers have no enforced encryption, patching, antivirus or screen-lock settings, and cannot be wiped if lost.'; Fix = 'Enroll the listed devices in Intune through Entra join or group-policy auto-enrollment.' }
    'compliance-none'      = @{ Title = 'No Intune compliance policies exist';                Why = 'Without compliance policies Intune cannot say whether a device is healthy, so no Conditional Access policy can require a compliant device and no device can be found noncompliant.'; Fix = 'Create a compliance policy per platform requiring encryption, PIN/password, minimum OS version and real-time protection.' }
    'config-none'          = @{ Title = 'No Intune configuration or security policies exist'; Why = 'Device security settings (BitLocker, Defender, firewall, screen lock, updates) are left to each user and each device''s defaults.'; Fix = 'Deploy a security baseline, BitLocker, Defender Antivirus and firewall, screen lock, and update rings to all devices.' }
    'policy-unassigned'    = @{ Title = 'Policy is not assigned to any device or user';       Why = 'A policy that exists but is not assigned does nothing; auditors will see a control that looks configured but is not applied.'; Fix = 'Assign the policy to the intended device or user groups, or remove it so the inventory reflects what is actually enforced.' }
    'consent-allowed'      = @{ Title = 'Users can consent to applications themselves';       Why = 'Users can grant third-party apps access to mail, files and contacts without IT review. Consent phishing uses exactly this path to steal data without a password.'; Fix = 'Set user consent to "Do not allow" or limit it to verified publishers and low-impact permissions, and enable the admin consent request workflow with named reviewers.' }
    'consent-workflow-off' = @{ Title = 'Admin consent request workflow is off';              Why = 'When users cannot consent and there is no request workflow, they ask for admin credentials or find workarounds; with the workflow, requests are logged and reviewed.'; Fix = 'Enable the admin consent request workflow and name reviewers who understand application permissions.' }
    'app-registration-open'= @{ Title = 'Any user can register applications';                Why = 'App registrations can be given credentials and permissions; letting all users create them widens the attack surface and clutters the tenant with unreviewed apps.'; Fix = 'Set "Users can register applications" to No and grant the Application Developer role to the people who need it.' }
    'app-unverified'       = @{ Title = 'Application from an unverified or unknown publisher'; Why = 'Microsoft has not confirmed who owns this application. Malicious consent-phishing apps are almost always unverified.'; Fix = 'Confirm the business owner and purpose of the app, remove it if unknown, and restrict user consent to verified publishers.' }
    'app-user-consent'     = @{ Title = 'Permissions granted by individual users';            Why = 'These permissions were approved by end users rather than an administrator, so nobody with security responsibility reviewed what the app can read.'; Fix = 'Review each user-consented app, revoke what is not needed, and restrict user consent going forward.' }
    'app-high-privilege'   = @{ Title = 'Application holds high-privilege permissions';       Why = 'Permissions that can read all mail or files, or write to the directory, let the app (and anyone who compromises it) act across the whole organization without a user signing in.'; Fix = 'Confirm the permission is required, prefer least-privilege alternatives, protect the app''s credentials, and review it at least annually.' }
    'ca-not-enabled'       = @{ Title = 'Conditional Access policy is not enforced';          Why = 'A policy in report-only or disabled state protects nothing; the control appears configured but is not in effect.'; Fix = 'Review the report-only results, then turn the policy on, or delete it if it is no longer needed.' }
    'security-defaults-off'= @{ Title = 'Security defaults are off without equivalent Conditional Access'; Why = 'Security defaults are Microsoft''s baseline (MFA for everyone, legacy authentication blocked). With them off and no Conditional Access policies providing the same protection, the tenant has no enforced MFA.'; Fix = 'Either turn security defaults on or implement Conditional Access policies for MFA and legacy-auth blocking.' }
    'tap-disabled'         = @{ Title = 'Temporary Access Pass is disabled';                  Why = 'Without TAP, help desks fall back to insecure onboarding and MFA-reset methods such as sharing passwords or temporarily removing MFA.'; Fix = 'Enable Temporary Access Pass with one-time use and a short lifetime, scoped to a help-desk-managed group.' }
    'auth-weak-method'     = @{ Title = 'Weak authentication method is enabled';              Why = 'SMS, voice calls and email codes can be intercepted or SIM-swapped and are not phishing-resistant. Regulators increasingly expect app-based or hardware MFA.'; Fix = 'Move users to Microsoft Authenticator (number matching) or FIDO2 keys, then disable SMS and voice for all but documented exceptions.' }
    'device-join-open'     = @{ Title = 'Any user can join devices to Entra ID';              Why = 'Users can join personal or unknown devices to the directory, which can then receive single sign-on and look like company devices.'; Fix = 'Restrict device join to selected users or groups and require MFA to register or join a device.' }
    'ga-count-high'        = @{ Title = 'More than four Global Administrators';               Why = 'Every Global Administrator can read all data and change every setting; each extra account is another way the whole tenant can be lost. Microsoft recommends fewer than five.'; Fix = 'Reduce to two to four, move day-to-day admins to least-privilege roles, and use PIM eligible assignments where licensed.' }
    'ga-single'            = @{ Title = 'Only one Global Administrator';                     Why = 'If this account is locked, lost or its owner leaves, nobody can administer the tenant.'; Fix = 'Create a second cloud-only break-glass Global Administrator with a long random password stored offline, excluded from Conditional Access and alerted on sign-in.' }
    'admin-no-mfa'         = @{ Title = 'Administrator without MFA registered';               Why = 'An administrator protected only by a password is the highest-value target in the tenant; compromise means total control.'; Fix = 'Register Microsoft Authenticator or a FIDO2 key for this account today and require MFA for all administrator roles.' }
    'user-no-mfa'          = @{ Title = 'User has no MFA method registered';                  Why = 'This account cannot be protected by MFA and will be the first to fall to a phished or reused password.'; Fix = 'Have the user register Microsoft Authenticator; block sign-in if they have left.' }
    'user-stale'           = @{ Title = 'Enabled account with no recent sign-in';             Why = 'Accounts nobody uses are rarely watched, keep a licence, and are a quiet way in for an attacker who guesses or finds the password. Examiners ask for a leaver and dormant-account process.'; Fix = 'Confirm whether the person still works here; disable the account, remove licences and convert the mailbox to shared if not.' }
    'user-disabled-licensed'= @{ Title = 'Disabled account still holds licences';            Why = 'Licences cost money and keep mail and files reachable through other channels; a disabled leaver should have none.'; Fix = 'Remove the licences, convert the mailbox to a shared mailbox if the content must be kept, and delete the account after the retention period.' }
    'device-stale'         = @{ Title = 'Device has not been seen recently';                  Why = 'Stale device objects may be lost, retired or re-imaged; they can still hold tokens and clutter compliance figures.'; Fix = 'Confirm the device''s status, wipe or retire it in Intune, and delete the Entra object if it is gone.' }
    'device-not-intune'    = @{ Title = 'Device is not managed by Intune';                    Why = 'Nothing enforces encryption, patching, antivirus or screen lock on this device, and it cannot be wiped if lost or stolen.'; Fix = 'Enroll the device in Intune, or remove it from the directory if it is personal and should not have access.' }
    'device-noncompliant'  = @{ Title = 'Device is noncompliant';                            Why = 'The device fails at least one required setting (encryption, OS version, protection). If Conditional Access requires compliance it is or should be blocked.'; Fix = 'Open the device in Intune to see the failing setting and remediate, or retire the device.' }
    'device-not-encrypted' = @{ Title = 'Device storage is not encrypted';                    Why = 'A lost or stolen unencrypted laptop exposes every file and cached mailbox on it and is a reportable incident for most institutions.'; Fix = 'Enable BitLocker (Windows) or FileVault (macOS) through an Intune policy and require encryption in the compliance policy.' }
    'forwarding'           = @{ Title = 'Mail is forwarded outside the organization';        Why = 'External forwarding is the classic sign of a compromised mailbox: attackers silently copy every message to themselves. It also moves regulated data outside your control.'; Fix = 'Remove the forward or rule unless there is a documented business reason, then block external auto-forwarding in the outbound spam policy.' }
    'remote-domain-forward'= @{ Title = 'Default remote domain allows automatic forwarding'; Why = 'This is one of two switches that control auto-forwarding; left on, it relies entirely on the outbound spam policy to block forwards.'; Fix = 'Set AutoForwardEnabled to false on the default remote domain (*) as defence in depth.' }
    'spam-forward-on'      = @{ Title = 'Outbound spam policy allows automatic forwarding';  Why = 'With automatic forwarding On, any user or attacker can set a rule that sends all mail outside the organization.'; Fix = 'Set automatic forwarding to Off (Automatic is also acceptable) in every outbound spam filter policy.' }
    'pop-imap'             = @{ Title = 'POP or IMAP is enabled';                            Why = 'POP and IMAP are legacy protocols that cannot use MFA and download whole mailboxes to any client.'; Fix = 'Disable POP and IMAP on every mailbox that does not have a documented need, and block legacy authentication with Conditional Access.' }
    'smtp-auth'            = @{ Title = 'SMTP AUTH is enabled';                              Why = 'SMTP AUTH lets any device send mail with just a username and password and is heavily used in password-spray and business email compromise attacks.'; Fix = 'Disable SMTP AUTH tenant-wide and enable it per mailbox only for devices (printers, scanners) that truly need it.' }
    'modern-auth-off'      = @{ Title = 'Modern authentication is disabled in Exchange';     Why = 'Without modern authentication, Outlook falls back to basic authentication, which cannot do MFA.'; Fix = 'Enable OAuth2ClientProfileEnabled on the organization configuration.' }
    'secure-score'         = @{ Title = 'Microsoft Secure Score is below target';            Why = 'Secure Score measures how many of Microsoft''s recommended security actions are in place compared with similar organizations; a low score means well-known protections are missing.'; Fix = 'Work through the Remediation Plan from the top; the Secure Score actions list the points each one recovers.' }
    'license-p1'           = @{ Title = 'Entra ID P1 is not licensed';                      Why = 'Conditional Access, the MFA registration report, risk-based sign-in and sign-in logs all require P1. Without it, protection relies on security defaults and these controls cannot be measured.'; Fix = 'License Entra ID P1 (included in Microsoft 365 Business Premium and E3).' }
    'license-intune'       = @{ Title = 'Intune is not licensed';                           Why = 'Device compliance, configuration policies and device-based Conditional Access are impossible without Intune, so endpoint security cannot be enforced or evidenced.'; Fix = 'License Microsoft Intune (included in Microsoft 365 Business Premium, E3 and E5).' }
    'sync-stale'           = @{ Title = 'Directory synchronization has not run recently';   Why = 'Entra Connect normally syncs every 30 minutes. A stale sync means disabled or removed AD accounts are still active in Microsoft 365 and password changes do not flow.'; Fix = 'Check the Entra Connect server and the Microsoft Entra Connect Sync service; review Entra Connect Health or the server event log.' }
    'spf-missing'          = @{ Title = 'No SPF record';                                    Why = 'Without SPF, any server on the internet can send mail claiming to be from this domain and receiving systems have no way to tell. Spoofed invoices and CEO-fraud emails to your customers become trivial.'; Fix = 'Publish a TXT record "v=spf1 include:spf.protection.outlook.com -all" (add other legitimate senders before -all).' }
    'spf-weak'             = @{ Title = 'SPF record ends in a soft fail or neutral';        Why = '~all or ?all tells receivers to accept (and merely flag) mail from unauthorized servers, so spoofing still gets delivered.'; Fix = 'Once every legitimate sender is listed, change the record to end in -all.' }
    'dkim-missing'         = @{ Title = 'DKIM is not enabled for this domain';              Why = 'DKIM signs outgoing mail so receivers can prove it was not altered and really came from your tenant. Without it DMARC cannot be enforced reliably and your mail is more likely to be marked as spam.'; Fix = 'Publish the two selector CNAME records Microsoft provides (selector1/selector2._domainkey) and enable DKIM signing for the domain in the Defender portal or with Set-DkimSigningConfig.' }
    'dmarc-missing'        = @{ Title = 'No DMARC record';                                  Why = 'DMARC tells receivers what to do with mail that fails SPF and DKIM and sends you reports of who is spoofing your domain. Examiners and cyber-insurers increasingly require it.'; Fix = 'Publish a TXT record at _dmarc.<domain>: start with "v=DMARC1; p=none; rua=mailto:<reports>" to collect reports, then move to p=quarantine and p=reject.' }
    'dmarc-none'           = @{ Title = 'DMARC policy is p=none (monitor only)';            Why = 'p=none collects reports but does not stop spoofed mail from being delivered; the domain is still fully spoofable.'; Fix = 'Review the DMARC reports, fix any legitimate senders failing alignment, then move to p=quarantine and finally p=reject.' }
    'spam-action-weak'     = @{ Title = 'High-confidence spam or phish is not quarantined'; Why = 'Messages Microsoft is highly confident are spam or phishing are being delivered to the Junk folder (or inbox) where a user can still open them.'; Fix = 'Set the high-confidence spam, phishing and high-confidence phishing actions to Quarantine in every anti-spam policy.' }
    'bulk-threshold'       = @{ Title = 'Bulk email threshold is high';                     Why = 'A bulk complaint level above 6 lets most marketing and grey-mail through, which carries a large share of credential-phishing links.'; Fix = 'Set the bulk threshold to 6 or lower (5 for strict).' }
    'zap-off'              = @{ Title = 'Zero-hour auto purge is disabled';                 Why = 'ZAP removes messages from inboxes when Microsoft learns they were malicious after delivery. Without it, a phish that was identified minutes later stays in every inbox it reached.'; Fix = 'Enable zero-hour auto purge for spam and phishing in the anti-spam policy.' }
    'malware-filefilter-off'= @{ Title = 'Common attachments filter is off';               Why = 'The filter blocks executable and script attachments (.exe, .js, .vbs ...) that are the most common malware carriers, regardless of signature detection.'; Fix = 'Enable the common attachments filter in the anti-malware policy and keep Microsoft''s default file-type list.' }
    'antiphish-impersonation-off' = @{ Title = 'Impersonation protection is not enabled'; Why = 'Impersonation protection catches lookalike display names and domains (the CEO-fraud and vendor-invoice attacks that cost banks the most). It is licensed but not switched on.'; Fix = 'In the anti-phishing policy enable user and domain impersonation protection, mailbox intelligence and its protection action, and add your executives as protected users.' }
    'phish-threshold-low'  = @{ Title = 'Phishing threshold is at the lowest level';        Why = 'Threshold 1 (Standard) applies the least aggressive machine-learning detection; Microsoft recommends 2 (Aggressive) or higher for organizations handling money.'; Fix = 'Set the phishing email threshold to 2 or 3 in the anti-phishing policy.' }
    'safelinks-none'       = @{ Title = 'No Safe Links policy in effect';                   Why = 'Safe Links rewrites and checks URLs at the moment a user clicks, blocking links that were safe at delivery and weaponized later. It is licensed but not applied.'; Fix = 'Create a Safe Links policy (or turn on the Standard preset) covering all users, with URL scanning for email, Teams and Office apps and click tracking enabled.' }
    'safeattach-none'      = @{ Title = 'No Safe Attachments policy in effect';             Why = 'Safe Attachments detonates attachments in a sandbox before delivery, catching new malware signatures miss. It is licensed but not applied.'; Fix = 'Create a Safe Attachments policy (or turn on the Standard preset) with the Block action for all users.' }
    'safeattach-spo-off'   = @{ Title = 'Defender for SharePoint, OneDrive and Teams is off'; Why = 'Malicious files uploaded to SharePoint, OneDrive or Teams are not scanned or blocked for other users.'; Fix = 'Turn on Defender for Office 365 for SharePoint, OneDrive and Microsoft Teams and enable Safe Documents in the global Safe Attachments settings.' }
    'preset-off'           = @{ Title = 'Preset security policies are not enabled';         Why = 'Microsoft''s Standard and Strict presets apply its recommended anti-spam, anti-phishing, Safe Links and Safe Attachments settings in one step and are kept current automatically. With neither enabled, protection depends on hand-maintained custom policies.'; Fix = 'Enable the Standard preset security policy for all users (Strict for executives and finance), or document why custom policies are equivalent.' }
    'defender-not-licensed'= @{ Title = 'Defender for Office 365 is not licensed';          Why = 'Only basic Exchange Online Protection is active: no Safe Links, Safe Attachments, impersonation protection or attack simulation. These are the controls that stop targeted phishing against bank staff.'; Fix = 'License Defender for Office 365 Plan 1 (included in Microsoft 365 Business Premium) or Plan 2 (E5).' }
    'mde-not-licensed'     = @{ Title = 'Defender for Endpoint is not licensed';            Why = 'Without endpoint detection and response, there is no central view of device threats, vulnerabilities or attacker activity on workstations.'; Fix = 'License Defender for Business (included in Microsoft 365 Business Premium) or Defender for Endpoint P1/P2 and onboard all devices.' }
    'mde-high-risk'        = @{ Title = 'Device has a high risk score in Defender';        Why = 'Defender has active alerts on this device indicating malware or attacker activity that has not been resolved.'; Fix = 'Investigate the alerts in the Defender portal, isolate the device if needed, and resolve the incidents.' }
    'mde-inactive'         = @{ Title = 'Defender sensor is inactive';                      Why = 'The device has not reported to Defender for Endpoint recently, so it is unprotected by EDR and may be offline, re-imaged or tampered with.'; Fix = 'Confirm the device status; re-onboard or retire it in Defender.' }
    'mde-not-onboarded'    = @{ Title = 'Managed device is not onboarded to Defender for Endpoint'; Why = 'The device is licensed for EDR but not sending telemetry, so threats on it are invisible.'; Fix = 'Deploy the Defender for Endpoint onboarding profile through Intune to all managed devices.' }
    # ---- Possible indications of compromise (section 2)
    'ioc-urgent'           = @{ Title = 'Urgent indicator: investigate today';                 Why = 'This pattern is one attackers create almost every time they get into a Microsoft 365 account or tenant. It can be legitimate, but it must be confirmed with the people involved before it is dismissed.'; Fix = 'Treat it as a security incident until proven otherwise: confirm with the user or administrator named, reset the password and revoke sessions for the affected account, remove the rule, credential or setting if it is not expected, and record the outcome. Involve your IT provider or incident response contact immediately if you cannot explain it.' }
    'ioc-review'           = @{ Title = 'Review indicator: confirm it was expected';           Why = 'Legitimate administration and attacker activity look the same in the audit log. A quick check with the person who made the change, or the user affected, turns an unknown into evidence.'; Fix = 'Compare each entry with change tickets, help-desk records or the administrator who made it. Anything nobody can account for should be escalated as a possible incident.' }
    'ioc-role-add'         = @{ Title = 'Privileged role assigned recently';                   Why = 'Adding an account to Global Administrator or another privileged role is a standard persistence step after a compromise and also the single most consequential change a legitimate administrator can make.'; Fix = 'Confirm every assignment against a change record. Remove any that are not justified, and use Privileged Identity Management (Entra ID P2) so elevated roles are time-limited and approved.' }
    'ioc-mfa-change'       = @{ Title = 'MFA methods registered or changed';                   Why = 'Attackers who obtain a password register their own Authenticator app or phone number so they keep access after the user changes the password.'; Fix = 'Ask each user whether they registered the method themselves. Remove unknown methods from the user''s authentication methods, reset the password, revoke sessions and review the user''s sign-in history.' }
    'ioc-pw-reset'         = @{ Title = 'Password reset by an administrator';                  Why = 'Help-desk password resets are normal, but a reset nobody requested is how attackers with admin access take over other accounts, and social-engineered resets are a common entry point.'; Fix = 'Match each reset to a help-desk ticket. Require identity verification before resets, and alert on resets of privileged accounts.' }
    'ioc-ca-change'        = @{ Title = 'Conditional Access or security policy changed';       Why = 'Weakening or excluding accounts from Conditional Access is how attackers switch off MFA for themselves. Policy changes should be rare and planned.'; Fix = 'Confirm each change with the administrator named and the change record. Restore any policy that was weakened and review exclusions.' }
    'ioc-domain-change'    = @{ Title = 'Domain or federation settings changed';               Why = 'Changing a domain to federated authentication lets an attacker sign in as any user without a password or MFA. This is a hallmark of tenant-level compromise and is almost never a routine change.'; Fix = 'Verify immediately with whoever manages identity. If unexpected, revert the domain to managed authentication, rotate Global Administrator credentials and engage incident response.' }
    'ioc-app-cred'         = @{ Title = 'Credential added to an application';                  Why = 'Adding a secret or certificate to an existing app registration gives the holder the app''s permissions, often mailbox or directory-wide, without any user sign-in or MFA. It is a favoured persistence technique.'; Fix = 'Confirm who added each credential and why. Remove unknown credentials, review the app''s permissions, and restrict who can manage applications.' }
    'ioc-new-app'          = @{ Title = 'New application added to the tenant';                 Why = 'Malicious OAuth apps, consented by a tricked user or an attacker, read mail and files through the API without touching the user''s password.'; Fix = 'Check each new app''s publisher and permissions in the Enterprise Applications section. Remove apps nobody recognises and restrict user consent to verified publishers or admin approval.' }
    'ioc-new-user'         = @{ Title = 'User accounts created recently';                      Why = 'Attackers create their own accounts, often with innocuous names, to keep access after the original victim is cleaned up.'; Fix = 'Match each new account to an HR or onboarding record. Disable any that cannot be matched.' }
    'ioc-guest'            = @{ Title = 'Guests invited recently';                             Why = 'A guest invitation gives an outside identity access to Teams, SharePoint and files. Invitations attackers send to themselves survive password resets.'; Fix = 'Confirm who invited each guest and why. Remove guests nobody can justify and restrict who may invite guests.' }
    'ioc-inbox-rule-hide'  = @{ Title = 'Inbox rule that deletes, hides or forwards mail';     Why = 'After taking over a mailbox, attackers create rules that delete or move incoming mail (often to RSS Feeds or Archive) or forward it out, so the victim does not see replies, security alerts or the fraudulent payment instructions being exchanged in their name. This is the most common indicator of business email compromise.'; Fix = 'Contact the mailbox owner. If the rule is unexpected, disable it, reset the password, revoke sessions, review sent items and forwarding, and search for mail the rule touched.' }
    'ioc-delegate'         = @{ Title = 'Full Access delegate on a mailbox';                   Why = 'Full Access lets another account read every message in the mailbox. Assistants and shared mailboxes legitimately use it; an unexpected delegate is a persistence technique.'; Fix = 'Review each delegate with the mailbox owner and remove any that are not required. Audit delegate changes.' }
    'ioc-sendas'           = @{ Title = 'Send As permission granted to another account';       Why = 'Send As lets one account send mail that appears to come from another, the exact capability needed for payment fraud in an executive''s name.'; Fix = 'Confirm each grant is required for a shared mailbox or assistant arrangement, and remove the rest.' }
    'ioc-transport-change' = @{ Title = 'Mail flow (transport) rule changed recently';          Why = 'Tenant-wide transport rules can silently copy, redirect or allow-list mail for every user. Attackers with admin access use them to exfiltrate mail or whitelist phishing.'; Fix = 'Confirm each change with the Exchange administrator and the change record. Review redirect, Bcc and bypass-filtering actions.' }
    'ioc-audit-off'        = @{ Title = 'Mailbox auditing is disabled';                        Why = 'Without mailbox audit logging there is no record of who read, moved or sent mail, so a compromise cannot be investigated or proven.'; Fix = 'Enable organization-wide mailbox auditing (Set-OrganizationConfig -AuditDisabled $false). It is on by default in Exchange Online.' }
    'ioc-risky-user'       = @{ Title = 'Entra ID Protection flags this user as at risk';      Why = 'Microsoft has correlated signals (leaked credentials, anomalous travel, malware-linked IP addresses) that indicate the account is probably compromised.'; Fix = 'Investigate the user''s recent sign-ins, reset the password, revoke sessions and review mailbox rules and MFA methods. Enable a user-risk Conditional Access policy so this happens automatically.' }
    'ioc-risk-detection'   = @{ Title = 'Risk detections recorded';                            Why = 'Each detection is a sign-in or user event Microsoft considers anomalous (unfamiliar location, impossible travel, anonymous IP, password spray).'; Fix = 'Review the detections for the users involved and confirm or dismiss each one in Entra ID Protection.' }
    'ioc-multi-country'    = @{ Title = 'User signed in from more than one country';           Why = 'Sign-ins from several countries in a short window usually mean travel, a VPN, or someone else using the credentials. For a community bank with no overseas staff it deserves a direct question.'; Fix = 'Confirm travel or VPN use with the user. If unexplained, reset the password, revoke sessions and review mailbox rules. Consider a Conditional Access policy that blocks sign-ins from countries where you have no staff.' }
    'ioc-legacy-signin'    = @{ Title = 'Successful legacy authentication sign-ins';            Why = 'Legacy protocols (IMAP, POP, SMTP AUTH, ActiveSync basic) bypass MFA. A successful legacy sign-in is a password being used without a second factor, possibly by an attacker.'; Fix = 'Identify the users and clients involved, migrate them to modern authentication, then block legacy authentication with Conditional Access.' }
    'ioc-mfa-denied'       = @{ Title = 'MFA prompts denied or failed';                        Why = 'A user denying MFA prompts they did not initiate means someone has their password. Repeated prompts are the "MFA fatigue" attack that preceded several major breaches.'; Fix = 'Ask the user whether they initiated the sign-ins. If not, reset the password immediately. Enable number matching and additional context in Authenticator.' }
    'ioc-ca-blocked'       = @{ Title = 'Sign-ins blocked by Conditional Access';               Why = 'Conditional Access only evaluates after the password was accepted. A block from an unexpected location or device means the credentials are known to someone outside your controls.'; Fix = 'Review the users and locations involved and reset passwords where the attempt was not the user.' }
    'ioc-spray'            = @{ Title = 'Possible password spray from one address';            Why = 'Many failed passwords against many different accounts from one IP address is the signature of a password spray. Any account using a weak or reused password will fall to it.'; Fix = 'Confirm MFA and legacy-authentication blocking cover every account, enable Smart Lockout and password protection, and review which accounts the address targeted.' }
    'ioc-lockout'          = @{ Title = 'Account lockouts';                                    Why = 'Lockouts show repeated wrong passwords against specific accounts, either a user with a stale saved password or an attacker guessing.'; Fix = 'Check the sign-in log for the accounts involved; if the attempts come from unknown addresses, treat them as attack traffic and confirm MFA on the targeted accounts.' }
    # ---- Teams (section 17)
    'teams-guests'         = @{ Title = 'Team contains guests';                               Why = 'Guests see every channel, file and conversation in the team. Guests who are no longer working with the institution are a common data-leak path.'; Fix = 'Have each team owner confirm every guest is still needed. Run access reviews (Entra ID P2) or review guests quarterly, and remove stale ones.' }
    'teams-public'         = @{ Title = 'Public team';                                        Why = 'Anyone in the organization can join a public team and read everything in it, including files. Customer, HR or board material in a public team is exposed to every employee and guest with directory access.'; Fix = 'Set teams that hold confidential material to Private and review membership.' }
    'teams-no-owner'       = @{ Title = 'Team has no owner';                                  Why = 'Without an owner nobody reviews membership, guests or sharing, and nobody can be asked whether the content is still needed.'; Fix = 'Assign at least two owners to every team; the Teams admin center lists ownerless teams.' }
    'teams-single-owner'   = @{ Title = 'Team has a single owner';                            Why = 'If the one owner leaves, the team becomes ownerless and its data and guests are unmanaged.'; Fix = 'Add a second owner.' }
    'teams-inactive'       = @{ Title = 'Inactive team';                                      Why = 'Teams nobody uses still hold data and guests. Stale content accumulates risk without any benefit.'; Fix = 'Archive or delete inactive teams after confirming with the owner; apply a Microsoft 365 group expiration policy.' }
    'groups-guests-open'   = @{ Title = 'Any member may add guests to Teams and Microsoft 365 Groups'; Why = 'With the default setting every team owner can invite any outside email address into the institution''s Teams and SharePoint content with no approval.'; Fix = 'Restrict guest invitations to administrators or to approved domains (Entra external collaboration settings and the Group.Unified AllowToAddGuests setting), or require an approval workflow.' }
    'teams-federation-open' = @{ Title = 'Teams external access allows all external domains'; Why = 'Open federation lets anyone with a Teams or Skype account chat and call your staff, which is now a major phishing and malware delivery channel ("Teams phishing").'; Fix = 'Restrict external access to an allow list of partner domains and block Teams consumer (personal) accounts.' }
    'teams-anon-join'      = @{ Title = 'Anonymous users can join meetings';                  Why = 'Anonymous join lets unauthenticated participants enter meetings with customers or the board by link alone.'; Fix = 'Set the global meeting policy so anonymous users wait in the lobby or cannot join, and restrict who bypasses the lobby.' }
    'teams-policies-na'    = @{ Title = 'Teams tenant policies not collected';                 Why = 'External access, guest access and meeting policies are tenant-wide controls that decide who outside the institution can reach staff through Teams.'; Fix = 'Install the MicrosoftTeams PowerShell module (Install-Module MicrosoftTeams -Scope CurrentUser) and grant the connector the Skype and Teams Tenant Admin API permission so the next report includes them.' }
    # ---- SharePoint and OneDrive (section 18)
    'spo-anyone-links'     = @{ Title = 'Anyone links are allowed';                           Why = '"Anyone" links open a file or folder to any person on the internet who has the link, with no sign-in and no record of who opened it. For a financial institution this is the highest-risk sharing setting.'; Fix = 'Set external sharing to "New and existing guests" or stricter in the SharePoint admin center, for both SharePoint and OneDrive, and expire existing anyone links.' }
    'spo-external-sharing' = @{ Title = 'External sharing is enabled';                        Why = 'Staff can share documents with any outside email address. That is often necessary, but it needs domain restrictions, expiry and monitoring.'; Fix = 'Limit sharing to allowed partner domains, set default link type to "People you choose", require guests to sign in with the invited account, and review externally shared files monthly.' }
    'spo-resharing'        = @{ Title = 'Guests can re-share content';                        Why = 'A guest who received a file can pass it on to anyone else, so the institution loses control after the first share.'; Fix = 'Turn off "Allow guests to share items they don''t own" in the SharePoint admin center.' }
    'spo-legacy-auth'      = @{ Title = 'Legacy authentication is allowed for SharePoint';    Why = 'Legacy (basic) authentication to SharePoint bypasses MFA and Conditional Access.'; Fix = 'Disable legacy authentication protocols in SharePoint admin center > Policies > Access control > Apps that don''t use modern authentication.' }
    'spo-idle-timeout-off' = @{ Title = 'Idle session sign-out is off';                       Why = 'On shared or unmanaged computers a browser left signed in to SharePoint or OneDrive stays open indefinitely.'; Fix = 'Enable idle session sign-out (for example warn at 55 minutes, sign out at 60) in SharePoint admin center > Policies > Access control.' }
    'spo-unmanaged-sync'   = @{ Title = 'OneDrive sync is allowed from unmanaged devices';    Why = 'Any PC, including a home computer, can sync the institution''s files locally where they are outside Intune, Defender and backup.'; Fix = 'Restrict the sync client to domain-joined or compliant devices and block sync of sensitive file types.' }
    'spo-site-inactive'    = @{ Title = 'Inactive site';                                      Why = 'Sites with no activity for months still hold documents, permissions and sharing links that nobody reviews.'; Fix = 'Confirm with the owner and archive or delete; apply retention and site lifecycle policies.' }
    'spo-site-no-owner'    = @{ Title = 'Site has no owner';                                  Why = 'Without an owner nobody is accountable for who can read the site''s documents or whether they are still needed.'; Fix = 'Assign an owner in the SharePoint admin center or delete the site.' }
    'od-owner-disabled'    = @{ Title = 'OneDrive belongs to a disabled or missing account';  Why = 'The files of departed staff remain accessible to whoever has been granted access, often indefinitely, and may include customer data.'; Fix = 'Transfer needed files to the manager and let the OneDrive retention period delete the rest; set the retention to the shortest period policy allows.' }
    'spo-external-shares'  = @{ Title = 'User shared files externally';                       Why = 'Each externally shared file is institutional data outside the tenant. Volume by user shows where review is needed.'; Fix = 'Review what was shared (Purview or the SharePoint sharing reports), confirm business need, and remove sharing that is no longer required.' }
    'passkey-not-started'  = @{ Title = 'Phishing-resistant MFA (passkeys) not rolled out';  Why = 'Push notifications and one-time codes can be relayed by modern phishing kits in real time; passkeys, security keys and Windows Hello cannot. NIST SP 800-63B calls these AAL3 authenticators and CISA recommends them for every organization, starting with administrators and finance.'; Fix = 'Enable the passkey (FIDO2) method, register passkeys for administrators and payment staff first, then all users, and track the percentage in this report each month.' }
    'passkey-admins'       = @{ Title = 'Administrators without phishing-resistant MFA';      Why = 'An administrator phished through a push or code prompt hands the attacker the whole tenant. Administrators are the smallest group and the highest value, so they go first.'; Fix = 'Register a passkey or FIDO2 key for every administrator and require the Phishing-resistant MFA authentication strength through Conditional Access for administrator roles.' }
    'reports-concealed'    = @{ Title = 'Usage report names are concealed';                  Why = 'The tenant''s report privacy setting replaces user, group and site names with hashes in the Microsoft 365 usage reports, so the OneDrive, SharePoint and Teams inventories in this report cannot show who owns what. Auditors and examiners need the names.'; Fix = 'Microsoft 365 admin center > Settings > Org settings > Services > Reports: clear "Display concealed user, group, and site names in all reports" and save. It is a tenant setting that this read-only script never changes; the next run shows real names.' }
    'ioc-rogue-app'        = @{ Title = 'Known-abused OAuth application is present';         Why = 'Products such as PerfectData Software, eM Client, rclone and Newsletter Software Supermailer are legitimate, but in a bank tenant they almost always arrive because an attacker with a stolen session consented to them to export the mailbox or OneDrive, send phishing, or keep access after a password reset.'; Fix = 'Confirm with the user and the IT provider whether anyone installed it deliberately. If not: Entra admin center > Enterprise applications > the app > Properties > Enabled for users to sign in = No, then Delete it; revoke the user''s sessions, reset the password, and review the mailbox for rules and sent mail. Block future consent to unverified apps.' }
    'ioc-unverified-app'   = @{ Title = 'Unverified publisher app with mail, file or contact access'; Why = 'Consent phishing tricks one user into approving an app that reads mail and files through the API, bypassing MFA. An app from an unverified publisher holding those permissions is the pattern, even when the name looks harmless.'; Fix = 'Identify the business owner of the app. If nobody claims it, disable and delete it and revoke its permission grants. Restrict user consent to verified publishers or admin approval.' }
    'ioc-anon-link'        = @{ Title = 'Anyone link created';                                Why = 'An Anyone link opens the file to any person on the internet who has the link, with no sign-in and no record of who opened it.'; Fix = 'Ask the user what was shared and why; remove the link (Manage access on the file) if it is not required, and set the tenant sharing level so Anyone links are not available.' }
    'ioc-ext-share'        = @{ Title = 'File shared with someone outside the organization';  Why = 'Often legitimate, but a compromised account sharing folders to an outside address is a common second route for data to leave.'; Fix = 'Confirm the recipient is a known partner, auditor or customer; otherwise remove the sharing and review the account for other indicators.' }
    'ioc-mass-download'    = @{ Title = 'Unusually high number of files downloaded';          Why = 'Hundreds of downloads or a full OneDrive sync in a few days from one account is how data walks out, whether through an intruder with a tool such as rclone or an employee about to leave.'; Fix = 'Ask the user and their manager what the downloads were for; check the IP addresses and client against their normal device; if unexplained, revoke sessions, reset the password and preserve the audit records. Restrict sync to managed devices.' }
    'ioc-audit-rule-change' = @{ Title = 'Mailbox rule, forwarding or permission change recorded'; Why = 'The audit log shows who created a rule that forwards, deletes or hides mail, or who granted mailbox access, from which address and client, including rules that were created and later removed.'; Fix = 'Compare the actor, IP address and client with the mailbox owner''s normal pattern; if it was not the owner from a known device, treat the mailbox as compromised.' }
    'ioc-url-click'        = @{ Title = 'User clicked a malicious link';                      Why = 'Safe Links blocked or warned about the destination; a user who clicked through may have typed their password into a phishing page.'; Fix = 'Contact the user, reset the password if they entered it, revoke sessions and check for new MFA methods and inbox rules. Use the event for awareness training.' }
    'ual-off'              = @{ Title = 'Cannot be checked: the unified audit log is turned off'; Why = 'With audit logging off, Microsoft 365 keeps no record of who shared, downloaded, forwarded or changed anything. Anyone links, external shares, mass downloads and mailbox-rule changes cannot be ruled out, a compromise cannot be investigated or scoped, and examiners have no activity evidence. CIS 3.1.1 and every regulator expect it on.'; Fix = 'Purview portal > Audit > Start recording user and admin activity (or Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true). Recording starts within hours; set retention to at least 180 days, ideally one year.' }
    'share-anonymous'      = @{ Title = 'Open to anyone on the internet';                      Why = 'An Anyone link or anonymous calendar permission needs no sign-in and leaves no record of who opened it. Whoever holds the link, including anyone it was forwarded to, has the data.'; Fix = 'Open the item (link in the row), Manage access, remove the Anyone link; then set the tenant and site sharing level so Anyone links cannot be created. For calendars set the Anonymous permission to None.' }
    'share-org'            = @{ Title = 'Open to the whole organization';                     Why = 'A "People in your organization" link or a public team lets every account in the tenant, including guests who have signed in, read the content. Board packs, HR and lending files should not be reachable by every teller and vendor account.'; Fix = 'Replace organization-wide links with specific people, make teams Private, and set the default link type to Specific people.' }
    'share-external'       = @{ Title = 'Shared with people outside the organization';         Why = 'Guests and external recipients keep access until someone removes it; most institutions have no process for that. Each one is a route for data to leave.'; Fix = 'Confirm each external person still needs access; remove the rest. Set guest access to expire and restrict sharing to approved domains.' }
    'share-outside-team'   = @{ Title = 'Shared with an insider outside the owning team';      Why = 'Someone who is not a member of the team or site has been given the file directly. That bypasses the team membership the owners manage, and it is how data crosses departments unnoticed.'; Fix = 'Ask the owner whether the person should be a team member instead, or remove the direct permission. Review cross-department grants in section 19 with department heads.' }
    'calendar-default-open' = @{ Title = 'Calendar readable by everyone in the organization';  Why = 'With the Default permission above Free/Busy, every account can read meeting subjects, locations and attendees, which reveals customer meetings, exams and deals.'; Fix = 'Set the Default calendar permission to AvailabilityOnly (free/busy) or LimitedDetails for the mailboxes listed; share full details only with named people.' }
    'calendar-anonymous'   = @{ Title = 'Calendar readable by anonymous users';               Why = 'Anyone with the published calendar link can read it without signing in.'; Fix = 'Set the Anonymous calendar permission to None and turn off anonymous calendar publishing in the sharing policy.' }
    'sharing-policy-external' = @{ Title = 'Calendar sharing policy allows external or anonymous sharing'; Why = 'The Exchange sharing policy decides whether users can publish or share calendar details outside the organization.'; Fix = 'In the Exchange admin center sharing policy, limit external calendar sharing to free/busy for approved domains and remove Anonymous publishing.' }
    'repeat-finding'       = @{ Title = 'Repeat finding';                                     Why = 'This finding has been open in three or more consecutive reports. Examiners read a repeat finding as a sign that the program is not being worked, not merely as a technical gap, and it often becomes a Matter Requiring Attention.'; Fix = 'Assign a named owner and a target date in the committee minutes, remove whatever is blocking it (budget, licence, vendor), and report the closure at the next review.' }
    'ga-count'             = @{ Title = 'Global Administrator count outside the recommended two to four'; Why = 'One administrator is a single point of failure; more than four widens the blast radius of a phished account.'; Fix = 'Keep two to four Global Administrators including a break-glass account; give everyone else a least-privilege role.' }
    'guest'                = @{ Title = 'Guest account';                                     Why = 'Guests are external people with access to shared content; stale or unreviewed guests are a common data-leak path.'; Fix = 'Review guests periodically, remove those no longer needed, and consider access reviews or restricting guest invitations.' }
}
# One rule for every status word in the report: good (pass), warn (action needed), mid (partial / review), na (not measured), info (neutral count)
function Get-StatusKind { param($Status)
    $v = [string]$Status
    if ($v -match '^(Enabled|Enforced|Deployed|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded|None found|Clear|On track|Done|Pass|True|Compliant|Required|Improved|Resolved|Low|On$|Active|Succeeded|success|Verified|\d+ polic)') { return 'good' }
    if ($v -match '^(Not available|Not collected|Not measured|Not checked|Not scanned|Not applicable|Unknown|n/a|Baseline|Skipped)') { return 'na' }
    if ($v -match '^(Disabled|Not enforced|Not blocked|No$|Not deployed|None$|None \(|Not started|Active external|Allowed by policy|Below target|Missing|Licensed, not configured|Licensed, not deployed|Licensed, not assigned|Licensed, setup not finished|Licensed, no onboarding policy|Attention needed|\d+ urgent|Anyone links|Off track|Urgent|False|Off$|Regressed|New$|High|Critical|Not configured|Not published|Not licensed|Slipping|Failed|Inactive|Impaired|Noncompliant|Not registered|Action needed|enabledForReportingButNotEnforced)') { return 'warn' }
    if ($v -match '^(Partial|Partly deployed|Onboarding policy in place|Licensed, deployment not verified|Allowed|All users allowed|No Windows|EOP only|No email domains|No devices onboarded|Guests present|In progress|\d+ to review|Review|Changed|Medium|Holding|Soft fail|Neutral|One selector|none \(monitor|quarantine|reportOnly|Pass all)') { return 'mid' }
    if ($v -match '^(\d+ assignments|\d+ \/ \d+|Info)') { return 'info' }
    return 'info'
}
function Get-ResultWord { param($Status) switch (Get-StatusKind $Status) { 'good' { 'Pass' } 'warn' { 'Action needed' } 'mid' { 'Partial: review' } 'na' { 'Not measured' } default { 'Info' } } }
function New-StatusBadge {
    # Icon + coloured text, identical everywhere: Pass, Partial, Action needed, Not measured.
    param($Status, [switch]$WordOnly)
    $kind = Get-StatusKind $Status
    $icon = switch ($kind) { 'good' { '&#10004;' } 'warn' { '&#10006;' } 'mid' { '&#9684;' } 'na' { '&mdash;' } default { '&bull;' } }
    $text = if ($WordOnly) { Get-ResultWord $Status } else { [string]$Status }
    return "<span class='st $kind' title='$(Get-ResultWord $Status)'>$icon</span><span class='stt $kind'>$(HtmlEnc $text)</span>"
}
function New-InfoIcon {
    param([string]$Issue, [string]$Detail)
    $i = $IssueCatalog[$Issue]
    if (-not $i) { return '' }
    return "<details class='info'><summary title='Configuration issue: click for details' aria-label='Configuration issue details'>i</summary><div class='tip'><b>$(HtmlEnc $i.Title)</b>$(if ($Detail) { "<p>$(HtmlEnc $Detail)</p>" })<p><b>Why it matters:</b> $(HtmlEnc $i.Why)</p><p><b>Recommended:</b> $(HtmlEnc $i.Fix)</p></div></details>"
}
function Get-RowIssueIcon {
    # $Issues: hashtable column -> issue id, or column -> @(issue id, detail text)
    param($Issues, [string]$Column)
    if (-not $Issues -or -not $Issues.Contains($Column)) { return '' }
    $v = $Issues[$Column]
    if ($v -is [array]) { return (New-InfoIcon -Issue $v[0] -Detail $v[1]) } ; return (New-InfoIcon -Issue ([string]$v))
}

function New-HtmlTable {
    <# $Rows: array of ordered hashtables or PSObjects. $Columns: property names in display order. #>
    <# $Id + $Filters add a filter bar (quick-filter buttons and a search box). Rows carry their filter tags in a
       hidden '_tags' key (space-separated); each filter is @{ Label = 'Stale'; Tag = 'stale' }. #>
    param($Rows, [string[]]$Columns, [string]$EmptyText = 'No records returned.', [string]$Class = 'grid', [string]$Id, $Filters)
    $rows = @($Rows | Where-Object { $null -ne $_ })
    if ($rows.Count -eq 0) { return "<p class='empty'>$(HtmlEnc $EmptyText)</p>" }
    if (-not $Columns) {
        $first = $rows[0]
        if ($first -is [System.Collections.IDictionary]) { $Columns = @($first.Keys) } else { $Columns = @($first.PSObject.Properties.Name) }
    }
    $Columns = @($Columns | Where-Object { $_ -notin @('_tags', '_issues', '_links', '_html') })
    $sb = New-Object System.Text.StringBuilder
    if ($Id -and $Filters) {
        [void]$sb.Append("<div class='filters' data-table='$Id'><button class='on' data-tag='all'>All ($($rows.Count))</button>")
        foreach ($flt in @($Filters)) {
            $cnt = @($rows | Where-Object { $tg = if ($_ -is [System.Collections.IDictionary]) { $_['_tags'] } else { $_._tags }; (" $tg " -like "* $($flt.Tag) *") }).Count
            [void]$sb.Append("<button data-tag='$(HtmlEnc $flt.Tag)'>$(HtmlEnc $flt.Label) ($cnt)</button>")
        }
        [void]$sb.Append("<input type='search' placeholder='Search this table...' aria-label='Search'><span class='n'></span></div>")
    }
    [void]$sb.Append("<table class='$Class'$(if ($Id) { " id='$Id'" })><thead><tr>")
    foreach ($c in $Columns) { [void]$sb.Append("<th>$(HtmlEnc $c)</th>") }
    [void]$sb.Append('</tr></thead><tbody>')
    foreach ($r in $rows) {
        $tags = if ($r -is [System.Collections.IDictionary]) { $r['_tags'] } else { $r.PSObject.Properties['_tags'].Value }
        $rowIssues = if ($r -is [System.Collections.IDictionary]) { $r['_issues'] } else { $r.PSObject.Properties['_issues'].Value }
        $rowLinks = if ($r -is [System.Collections.IDictionary]) { $r['_links'] } else { $r.PSObject.Properties['_links'].Value }
        $rowHtml = if ($r -is [System.Collections.IDictionary]) { $r['_html'] } else { $r.PSObject.Properties['_html'].Value }
        [void]$sb.Append($(if ($tags) { "<tr data-tags='$(HtmlEnc $tags)'>" } else { '<tr>' }))
        foreach ($c in $Columns) {
            $v = if ($r -is [System.Collections.IDictionary]) { $r[$c] } else { $r.$c }
            if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = Join-List $v '; ' }
            $cell = HtmlEnc $v
            if ($rowLinks -and $rowLinks[$c]) { $cell = "<a href='$(HtmlEnc $rowLinks[$c])'>$cell</a>" }
            if ($rowHtml -and $rowHtml.Contains($c)) { $cell = [string]$rowHtml[$c] }   # pre-built, already encoded HTML
            $cls = ''
            if ($c -in @('Status', 'Result', 'Level', 'Rollout stage') -and ([string]$v) -ne '' -and -not ($rowHtml -and $rowHtml.Contains($c))) { $cell = New-StatusBadge $v }
            switch -Regex ([string]$v) {
                '^(True|Enabled|enabled|Compliant|Required|Yes|Block|block|On|Improved|Resolved|Clear|Low|On track|Done)$' { $cls = " class='good'" }
                '^(False|Disabled|disabled|No|Not registered|Off|enabledForReportingButNotEnforced|Regressed|New|Urgent|High|Critical|Off track)$' { $cls = " class='warn'" }
                '^(Review|Partial|Changed|Medium|Not started|In progress.*)$' { $cls = " class='mid'" }
            }
            [void]$sb.Append("<td$cls>$cell$(Get-RowIssueIcon $rowIssues $c)</td>")
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</tbody></table>')
    return $sb.ToString()
}

function New-KvTable {
    # $Issues: hashtable key -> issue id (or @(issue id, detail)) to flag that row with an (i) icon
    param([System.Collections.Specialized.OrderedDictionary]$Pairs, $Issues)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<table class='kv'><tbody>")
    foreach ($k in $Pairs.Keys) {
        $v = $Pairs[$k]
        if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) {
            # array values render as a bulleted list
            $cell = '<ul class="kvlist">' + ((@($v) | Where-Object { $null -ne $_ } | ForEach-Object { "<li>$(HtmlEnc $_)</li>" }) -join '') + '</ul>'
        } else { $cell = HtmlEnc $v }
        [void]$sb.Append("<tr><th>$(HtmlEnc $k)</th><td>$cell$(Get-RowIssueIcon $Issues $k)</td></tr>")
    }
    [void]$sb.Append('</tbody></table>')
    return $sb.ToString()
}

function Get-Platform {
    param([string]$odataType)
    switch -Regex ($odataType) {
        'androidWorkProfile|androidDeviceOwner|aosp' { 'Android Enterprise'; break }
        'android'                                    { 'Android device administrator'; break }
        '^#microsoft\.graph\.ios'                    { 'iOS/iPadOS'; break }
        'macOS'                                      { 'macOS'; break }
        'windows10|windows81|windowsPhone'           { 'Windows'; break }
        default                                      { ($odataType -replace '#microsoft.graph.', '') }
    }
}

function Format-Assignments {
    param($assignments)
    if ($null -eq $assignments -or @($assignments).Count -eq 0) { return @{ Included = 'Not assigned'; Excluded = 'None' } }
    $inc = @(); $exc = @()
    foreach ($a in $assignments) {
        $t = $a.target
        $label = switch -Regex ($t.'@odata.type') {
            'allLicensedUsersAssignmentTarget' { 'All Users'; break }
            'allDevicesAssignmentTarget'       { 'All Devices'; break }
            'GroupAssignmentTarget'            { [string](Resolve-ObjectNames -Ids @($t.groupId)); break }
            default                            { ($t.'@odata.type' -replace '#microsoft.graph.', '') }
        }
        if ($t.deviceAndAppManagementAssignmentFilterId) {
            $label = "$label [filter: $($t.deviceAndAppManagementAssignmentFilterType)]"
        }
        if ($t.'@odata.type' -match 'exclusionGroupAssignmentTarget') { $exc += $label } else { $inc += $label }
    }
    return @{ Included = (Join-List $inc '; '); Excluded = $(if ($exc.Count) { Join-List $exc '; ' } else { 'None' }) }
}

$SettingsSkip = @('id', '@odata.type', '@odata.context', 'createdDateTime', 'lastModifiedDateTime', 'version', 'roleScopeTagIds',
    'assignments', 'scheduledActionsForRule', 'displayName', 'description', 'assignments@odata.context',
    'scheduledActionsForRule@odata.context', 'supportsScopeTags', 'deviceManagementApplicabilityRuleOsEdition',
    'deviceManagementApplicabilityRuleOsVersion', 'deviceManagementApplicabilityRuleDeviceMode')

function Format-SettingName { param([string]$n)
    $s = ($n -creplace '([a-z])([A-Z])', '$1 $2') -replace '_', ' '
    if ($s.Length -gt 1) { return $s.Substring(0, 1).ToUpper() + $s.Substring(1) } else { return $s.ToUpper() }
}

function Format-SettingValue { param($v)
    if ($null -eq $v) { return 'Not configured' }
    if ($v -is [bool]) { return $(if ($v) { 'Required' } else { 'Not configured' }) }
    if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $arr = @($v); if ($arr.Count -eq 0) { return 'None' } ; return (Join-List ($arr | ForEach-Object { if ($_ -is [PSObject] -and $_.PSObject.Properties.Count -gt 1) { ($_ | ConvertTo-Json -Compress) } else { $_ } }) '; ') }
    if ($v -is [PSObject] -and $v.PSObject.Properties.Count -gt 1) { return ($v | ConvertTo-Json -Compress) }
    return [string]$v
}

# ----------------------------------------------------------------------------------------------
# 1. Tenant overview
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting tenant information..." -ForegroundColor Cyan
$org = (Get-GraphAllSafe 'v1.0/organization' 'organization') | Select-Object -First 1
$skus = Get-GraphAllSafe 'v1.0/subscribedSkus' 'subscribed SKUs'
$SkuMap = @{}
foreach ($s in $skus) { $SkuMap[$s.skuId] = $s.skuPartNumber }

# Identify the client from the tenant itself unless -ClientName was given: tenant display name plus default domain.
$DefaultDomain = ($org.verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1).name
if (-not $DefaultDomain) { $DefaultDomain = ($org.verifiedDomains | Where-Object { $_.isInitial } | Select-Object -First 1).name }
if (-not $ClientName) { $ClientName = $(if ($org.displayName) { $org.displayName } elseif ($DefaultDomain) { $DefaultDomain } else { 'Client' }) }
$ClientLabel = $(if ($DefaultDomain -and $ClientName -ne $DefaultDomain) { "$ClientName ($DefaultDomain)" } else { $ClientName })
Write-Host "Client: $ClientLabel" -ForegroundColor Cyan

$tenantKv = [ordered]@{
    'Client'                  = $ClientName
    'Default domain'          = $(if ($DefaultDomain) { $DefaultDomain } else { 'Not found' })
    'Tenant display name'     = $org.displayName
    'Tenant ID'               = $org.id
    'Report generated'        = $ReportDate.ToString('yyyy-MM-dd HH:mm zzz')
    'Generated by'            = $ctx.Account
    'Audit log window'        = "$AuditDays days"
}
$skuRows = foreach ($s in $skus) {
    [ordered]@{ 'License SKU' = $s.skuPartNumber; 'Enabled' = $s.prepaidUnits.enabled; 'Consumed' = $s.consumedUnits; 'Status' = $s.capabilityStatus }
}

# Licensed capabilities, derived from the service plans inside the subscribed SKUs. Used to explain gaps
# (Graph answers a bare 400/403 for unlicensed features) and listed on the overview for the auditor.
$servicePlanNames = @($skus | Where-Object { $_.capabilityStatus -eq 'Enabled' } | ForEach-Object { $_.servicePlans } | Where-Object { $_.provisioningStatus -ne 'Disabled' } | ForEach-Object { $_.servicePlanName })
$script:HasIntune  = [bool]@($servicePlanNames | Where-Object { $_ -match '^INTUNE_(A|EDU|P2|SMBIZ)(_|$)' }).Count   # INTUNE_O365 (basic mobility) is not Intune
$script:HasEntraP1 = [bool]@($servicePlanNames | Where-Object { $_ -match '^AAD_PREMIUM' }).Count
$script:HasEntraP2 = [bool]@($servicePlanNames | Where-Object { $_ -eq 'AAD_PREMIUM_P2' }).Count
$script:HasDefenderP1 = [bool]@($servicePlanNames | Where-Object { $_ -match '^ATP_ENTERPRISE|^THREAT_INTELLIGENCE' }).Count
$script:HasDefenderP2 = [bool]@($servicePlanNames | Where-Object { $_ -match '^THREAT_INTELLIGENCE' }).Count
$script:HasMDE        = [bool]@($servicePlanNames | Where-Object { $_ -match '^WINDEFATP|^MDE_SMB|^DEFENDER_ENDPOINT|^MDE_LITE|Microsoft_Defender_For_Endpoint' }).Count
$script:MdePlan       = if (@($servicePlanNames | Where-Object { $_ -match '^WINDEFATP' }).Count) { 'Defender for Endpoint Plan 2' } elseif (@($servicePlanNames | Where-Object { $_ -match '^MDE_SMB' }).Count) { 'Defender for Business' } elseif (@($servicePlanNames | Where-Object { $_ -match '^MDE_LITE' }).Count) { 'Defender for Endpoint Plan 1' } elseif ($script:HasMDE) { 'Defender for Endpoint' } else { '' }
$tenantKv['Licensed capabilities'] = @(
    "Microsoft Intune (device management): $(if ($script:HasIntune) { 'Licensed' } else { 'Not licensed' })",
    "Entra ID P1 (Conditional Access, MFA registration report, sign-in logs): $(if ($script:HasEntraP1) { 'Licensed' } else { 'Not licensed' })",
    "Entra ID P2 (Privileged Identity Management, Identity Protection): $(if ($script:HasEntraP2) { 'Licensed' } else { 'Not licensed' })",
    "Defender for Office 365 (Safe Links, Safe Attachments, impersonation protection): $(if ($script:HasDefenderP2) { 'Licensed (Plan 2)' } elseif ($script:HasDefenderP1) { 'Licensed (Plan 1)' } else { 'Not licensed (Exchange Online Protection only)' })",
    "Defender for Endpoint / Defender for Business (endpoint detection and response): $(if ($script:HasMDE) { "Licensed ($script:MdePlan)" } else { 'Not licensed' })"
)
$domainRows = foreach ($d in (@($org.verifiedDomains) | Sort-Object @{ Expression = { -not $_.isDefault } }, name)) {
    [ordered]@{ 'Domain' = $d.name; 'Default' = $(if ($d.isDefault) { 'Yes' } else { '' }); 'Initial (.onmicrosoft.com)' = $(if ($d.isInitial) { 'Yes' } else { '' }); 'Type' = $d.type; 'Capabilities' = $d.capabilities }
}

# Microsoft Secure Score: current score, target, comparison averages, and the improvement actions (for the remediation plan).
Write-Host "Collecting Microsoft Secure Score..." -ForegroundColor Cyan
$secureScore   = (Get-GraphAllSafe 'v1.0/security/secureScores?$top=1' 'Microsoft Secure Score' -NoPaging) | Select-Object -First 1
$scoreProfiles = if ($secureScore) { Get-GraphAllSafe 'v1.0/security/secureScoreControlProfiles' 'Secure Score control profiles' } else { @() }
$scoreCurrent = $null; $scoreMax = $null; $scorePct = $null; $scoreTarget = $null; $scoreGap = $null; $scoreCompareText = ''
if ($secureScore -and $secureScore.maxScore) {
    $scoreCurrent = [math]::Round([double]$secureScore.currentScore, 1)
    $scoreMax     = [math]::Round([double]$secureScore.maxScore, 1)
    $scorePct     = [math]::Round($scoreCurrent / $scoreMax * 100)
    $scoreTarget  = [math]::Round($scoreMax * $TargetScorePercent / 100, 1)
    $scoreGap     = [math]::Round([math]::Max(0, $scoreTarget - $scoreCurrent), 1)
    $scoreCompareText = Join-List (@($secureScore.averageComparativeScores) | ForEach-Object {
        $label = switch ($_.basis) { 'AllTenants' { 'all organizations' } 'TotalSeats' { 'organizations of similar size' } 'IndustryTypes' { 'same industry' } default { $_.basis } }
        "$label average $([math]::Round([double]$_.averageScore, 1))"
    }) '; '
    $tenantKv['Microsoft Secure Score'] = "$scoreCurrent of $scoreMax points ($scorePct%) as of $(Format-DateText $secureScore.createdDateTime)"
    $tenantKv['Target Secure Score']    = "$scoreTarget points ($TargetScorePercent% of maximum). $(if ($scoreGap -gt 0) { "Gap: $scoreGap points. See the Remediation Plan." } else { 'Target met.' })"
    if ($scoreCompareText) { $tenantKv['Secure Score comparison'] = $scoreCompareText }
} else {
    $tenantKv['Microsoft Secure Score'] = "Not available. $(Get-GraphFailure 'Microsoft Secure Score')"
}

# ----------------------------------------------------------------------------------------------
# 2. Intune compliance policies
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting Intune compliance policies..." -ForegroundColor Cyan
$compliancePolicies = Get-GraphAllSafe 'v1.0/deviceManagement/deviceCompliancePolicies?$expand=assignments,scheduledActionsForRule($expand=scheduledActionConfigurations)' 'compliance policies'
$complianceSummary  = (Get-GraphAllSafe 'v1.0/deviceManagement/deviceCompliancePolicyDeviceStateSummary' 'compliance state summary') | Select-Object -First 1
$complianceSettings = (Get-GraphAllSafe 'v1.0/deviceManagement?$select=settings' 'compliance policy settings') | Select-Object -First 1

$complianceHtml = New-Object System.Text.StringBuilder
if ($complianceSettings.settings) {
    $cs = $complianceSettings.settings
    [void]$complianceHtml.Append('<h3>Compliance policy settings (tenant-wide)</h3>')
    [void]$complianceHtml.Append((New-KvTable ([ordered]@{
        'Mark devices with no compliance policy assigned as' = $(if ($cs.secureByDefault) { 'Not compliant' } else { 'Compliant' })
        'Compliance status validity period (days)'           = $cs.deviceComplianceCheckinThresholdDays
    })))
}
if ($complianceSummary) {
    [void]$complianceHtml.Append('<h3>Device compliance state (all policies)</h3>')
    [void]$complianceHtml.Append((New-KvTable ([ordered]@{
        'Compliant'        = $complianceSummary.compliantDeviceCount
        'Not compliant'    = $complianceSummary.nonCompliantDeviceCount
        'In grace period'  = $complianceSummary.inGracePeriodCount
        'Not applicable'   = $complianceSummary.notApplicableDeviceCount
        'Error'            = $complianceSummary.errorDeviceCount
        'Conflict'         = $complianceSummary.conflictDeviceCount
        'Unknown'          = $complianceSummary.unknownDeviceCount
    })))
}

$policyGroups = @($compliancePolicies | Group-Object { Get-Platform $_.'@odata.type' } | Sort-Object Name)
if (@($compliancePolicies).Count -eq 0) { [void]$complianceHtml.Append("<p class='empty'>$(HtmlEnc $(if (Get-GraphFailure 'compliance policies') { "Compliance policies could not be read: $(Get-GraphFailure 'compliance policies')" } else { 'No device compliance policies exist in this tenant.' }))</p>") }

foreach ($grp in $policyGroups) {
    [void]$complianceHtml.Append("<h3 class='platform'>$(HtmlEnc $grp.Name)</h3>")
    foreach ($p in ($grp.Group | Sort-Object displayName)) {
        $asg = Format-Assignments $p.assignments
        $basics = [ordered]@{
            'Name'          = $p.displayName
            'Description'   = $(if ($p.description) { $p.description } else { '--' })
            'Platform'      = $grp.Name
            'Profile type'  = ($p.'@odata.type' -replace '#microsoft.graph.', '')
            'Created'       = Format-DateText $p.createdDateTime
            'Last modified' = Format-DateText $p.lastModifiedDateTime
            'Version'       = $p.version
        }
        $settings = [ordered]@{}
        foreach ($prop in $p.PSObject.Properties) {
            if ($SettingsSkip -contains $prop.Name) { continue }
            $settings[(Format-SettingName $prop.Name)] = Format-SettingValue $prop.Value
        }
        $actionRows = @()
        foreach ($rule in @($p.scheduledActionsForRule)) {
            foreach ($a in @($rule.scheduledActionConfigurations)) {
                $when = switch ($a.gracePeriodHours) { 0 { 'Immediately' } default { "$([int]($a.gracePeriodHours / 24)) days ($($a.gracePeriodHours) hours)" } }
                $actionName = switch ($a.actionType) {
                    'block'              { 'Mark device noncompliant' }
                    'notification'       { 'Send email to end user' }
                    'pushNotification'   { 'Send push notification to end user' }
                    'retire'             { 'Retire the noncompliant device' }
                    'remoteLock'         { 'Remotely lock the noncompliant device' }
                    'wipe'               { 'Wipe device' }
                    default              { $a.actionType }
                }
                $actionRows += [ordered]@{ 'Action' = $actionName; 'Schedule' = $when; 'Additional recipients' = $(if ($a.notificationMessageCCList) { Join-List (Resolve-ObjectNames $a.notificationMessageCCList) } else { 'None selected' }) }
            }
        }
        $assignKv = [ordered]@{ 'Included groups' = $asg.Included; 'Excluded groups' = $asg.Excluded }

        [void]$complianceHtml.Append("<div class='card'><h4>$(HtmlEnc $p.displayName)</h4>")
        [void]$complianceHtml.Append('<h5>Basics</h5>'); [void]$complianceHtml.Append((New-KvTable $basics))
        [void]$complianceHtml.Append('<h5>Compliance settings</h5>'); [void]$complianceHtml.Append((New-KvTable $settings))
        [void]$complianceHtml.Append('<h5>Actions for noncompliance</h5>'); [void]$complianceHtml.Append((New-HtmlTable $actionRows -Columns 'Action', 'Schedule', 'Additional recipients' -EmptyText 'No actions configured.'))
        [void]$complianceHtml.Append('<h5>Assignments</h5>'); [void]$complianceHtml.Append((New-KvTable $assignKv))
        [void]$complianceHtml.Append('</div>')
    }
}

# ----------------------------------------------------------------------------------------------
# 3. Intune device configuration profiles
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting Intune configuration profiles and policies..." -ForegroundColor Cyan
$configProfiles  = Get-GraphAllSafe 'v1.0/deviceManagement/deviceConfigurations?$expand=assignments' 'device configuration profiles'
$settingsCatalog = Get-GraphAllSafe 'beta/deviceManagement/configurationPolicies?$expand=assignments' 'settings catalog policies'
$secIntents      = Get-GraphAllSafe 'beta/deviceManagement/intents?$expand=assignments' 'endpoint security policies'
$enrollConfigs   = Get-GraphAllSafe 'v1.0/deviceManagement/deviceEnrollmentConfigurations?$expand=assignments' 'enrollment configurations'
$appProtection   = Get-GraphAllSafe 'v1.0/deviceAppManagement/managedAppPolicies' 'app protection policies'

$profileRows = @()
$profileDetailHtml = New-Object System.Text.StringBuilder

# Classic device configuration profiles: every typed setting
foreach ($c in ($configProfiles | Sort-Object displayName)) {
    $asg = Format-Assignments $c.assignments
    $platform = Get-Platform $c.'@odata.type'
    $profileRows += [ordered]@{ 'Policy' = $c.displayName; 'Kind' = 'Configuration profile'; 'Platform' = $platform; 'Type' = ($c.'@odata.type' -replace '#microsoft.graph.', ''); 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included' = $asg.Included; 'Excluded' = $asg.Excluded }
    $settings = [ordered]@{}
    foreach ($prop in $c.PSObject.Properties) {
        if ($SettingsSkip -contains $prop.Name) { continue }
        $val = Format-SettingValue $prop.Value
        if ($val -eq 'Not configured' -or $val -eq 'None') { continue }   # keep the table to what is actually set
        $settings[(Format-SettingName $prop.Name)] = $val
    }
    [void]$profileDetailHtml.Append("<div class='card'><h4>$(HtmlEnc $c.displayName)</h4>")
    [void]$profileDetailHtml.Append((New-KvTable ([ordered]@{ 'Kind' = 'Configuration profile'; 'Platform' = $platform; 'Type' = ($c.'@odata.type' -replace '#microsoft.graph.', ''); 'Description' = $(if ($c.description) { $c.description } else { '--' }); 'Created' = (Format-DateText $c.createdDateTime); 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included groups' = $asg.Included; 'Excluded groups' = $asg.Excluded })))
    [void]$profileDetailHtml.Append('<h5>Configured settings</h5>')
    if ($settings.Count) { [void]$profileDetailHtml.Append((New-KvTable $settings)) } else { [void]$profileDetailHtml.Append("<p class='empty'>No settings are configured in this profile.</p>") }
    [void]$profileDetailHtml.Append('</div>')
}

# Settings catalog policies: each setting definition and its value
function Get-CatalogValue {
    param($inst)
    switch -Regex ($inst.'@odata.type') {
        'choiceSettingInstance$'           { $v = ($inst.choiceSettingValue.value -replace "^$([regex]::Escape($inst.settingDefinitionId))_", ''); return $v }
        'simpleSettingInstance$'           { return [string]$inst.simpleSettingValue.value }
        'simpleSettingCollectionInstance$' { return (Join-List ($inst.simpleSettingCollectionValue | ForEach-Object { $_.value }) '; ') }
        'choiceSettingCollectionInstance$' { return (Join-List ($inst.choiceSettingCollectionValue | ForEach-Object { $_.value -replace "^$([regex]::Escape($inst.settingDefinitionId))_", '' }) '; ') }
        'groupSettingCollectionInstance$'  { return "$(@($inst.groupSettingCollectionValue).Count) group(s) configured" }
        default                            { return ($inst.'@odata.type' -replace '#microsoft.graph.deviceManagementConfiguration', '') }
    }
}
foreach ($c in ($settingsCatalog | Sort-Object name)) {
    $asg = Format-Assignments $c.assignments
    $profileRows += [ordered]@{ 'Policy' = $c.name; 'Kind' = 'Settings catalog'; 'Platform' = $c.platforms; 'Type' = $c.technologies; 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included' = $asg.Included; 'Excluded' = $asg.Excluded }
    $catSettings = Get-GraphAllSafe "beta/deviceManagement/configurationPolicies/$($c.id)/settings" "settings for $($c.name)"
    $settings = [ordered]@{}
    foreach ($s in $catSettings) {
        $inst = $s.settingInstance
        $name = ($inst.settingDefinitionId -split '_')[-1]
        $settings[(Format-SettingName $name)] = Get-CatalogValue $inst
        if ($inst.choiceSettingValue.children) {
            foreach ($child in $inst.choiceSettingValue.children) { $settings["  $(Format-SettingName (($child.settingDefinitionId -split '_')[-1]))"] = Get-CatalogValue $child }
        }
    }
    [void]$profileDetailHtml.Append("<div class='card'><h4>$(HtmlEnc $c.name)</h4>")
    [void]$profileDetailHtml.Append((New-KvTable ([ordered]@{ 'Kind' = 'Settings catalog policy'; 'Platform' = $c.platforms; 'Technology' = $c.technologies; 'Description' = $(if ($c.description) { $c.description } else { '--' }); 'Created' = (Format-DateText $c.createdDateTime); 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included groups' = $asg.Included; 'Excluded groups' = $asg.Excluded })))
    [void]$profileDetailHtml.Append('<h5>Configured settings</h5>')
    if ($settings.Count) { [void]$profileDetailHtml.Append((New-KvTable $settings)) } else { [void]$profileDetailHtml.Append("<p class='empty'>No settings are configured in this policy.</p>") }
    [void]$profileDetailHtml.Append('</div>')
}

# Endpoint security policies (intents): settings with values
foreach ($c in ($secIntents | Sort-Object displayName)) {
    $asg = Format-Assignments $c.assignments
    $profileRows += [ordered]@{ 'Policy' = $c.displayName; 'Kind' = 'Endpoint security'; 'Platform' = ''; 'Type' = $c.templateId; 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included' = $asg.Included; 'Excluded' = $asg.Excluded }
    $intentSettings = Get-GraphAllSafe "beta/deviceManagement/intents/$($c.id)/settings" "settings for $($c.displayName)"
    $settings = [ordered]@{}
    foreach ($s in $intentSettings) {
        $v = if ($null -ne $s.value) { $s.value } elseif ($s.valueJson) { $s.valueJson } else { $null }
        if ($null -eq $v -or "$v" -in @('null', '', 'notConfigured')) { continue }
        $settings[(Format-SettingName (($s.definitionId -split '_')[-1]))] = Format-SettingValue $v
    }
    [void]$profileDetailHtml.Append("<div class='card'><h4>$(HtmlEnc $c.displayName)</h4>")
    [void]$profileDetailHtml.Append((New-KvTable ([ordered]@{ 'Kind' = 'Endpoint security policy'; 'Enabled' = $c.isAssigned; 'Description' = $(if ($c.description) { $c.description } else { '--' }); 'Last modified' = (Format-DateText $c.lastModifiedDateTime); 'Included groups' = $asg.Included; 'Excluded groups' = $asg.Excluded })))
    [void]$profileDetailHtml.Append('<h5>Configured settings</h5>')
    if ($settings.Count) { [void]$profileDetailHtml.Append((New-KvTable $settings)) } else { [void]$profileDetailHtml.Append("<p class='empty'>No settings are configured in this policy.</p>") }
    [void]$profileDetailHtml.Append('</div>')
}

# Enrollment restrictions and app protection: list with key values
$enrollRows = foreach ($e in ($enrollConfigs | Sort-Object priority)) {
    $asg = Format-Assignments $e.assignments
    $detail = @()
    foreach ($prop in $e.PSObject.Properties) {
        if ($SettingsSkip -contains $prop.Name -or $prop.Name -in @('priority', 'deviceEnrollmentConfigurationType')) { continue }
        $val = Format-SettingValue $prop.Value
        if ($val -notin @('Not configured', 'None')) { $detail += "$(Format-SettingName $prop.Name)=$val" }
    }
    [ordered]@{ 'Configuration' = $e.displayName; 'Type' = ($e.'@odata.type' -replace '#microsoft.graph.', ''); 'Priority' = $e.priority; 'Settings' = (Join-List $detail '; '); 'Assigned to' = $asg.Included }
}
$mamRows = foreach ($m in ($appProtection | Sort-Object displayName)) {
    $detail = @()
    foreach ($prop in $m.PSObject.Properties) {
        if ($SettingsSkip -contains $prop.Name -or $prop.Name -in @('isAssigned', 'deployedAppCount', 'apps', 'deploymentSummary')) { continue }
        $val = Format-SettingValue $prop.Value
        if ($val -notin @('Not configured', 'None', 'False', '0')) { $detail += "$(Format-SettingName $prop.Name)=$val" }
    }
    [ordered]@{ 'App protection policy' = $m.displayName; 'Platform' = ($m.'@odata.type' -replace '#microsoft.graph.', '' -replace 'ManagedAppProtection', ''); 'Assigned' = $m.isAssigned; 'Last modified' = (Format-DateText $m.lastModifiedDateTime); 'Key settings' = (Join-List $detail '; ') }
}

# ----------------------------------------------------------------------------------------------
# 3b. Intune managed devices vs Entra devices
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting Intune managed devices and Entra devices..." -ForegroundColor Cyan
$managedDevices = Get-GraphAllSafe 'v1.0/deviceManagement/managedDevices?$top=999' 'Intune managed devices'
$entraDeviceSelect = 'id,deviceId,displayName,operatingSystem,operatingSystemVersion,trustType,isManaged,isCompliant,accountEnabled,approximateLastSignInDateTime,registrationDateTime,managementType,profileType,manufacturer,model'
$entraDevices = @()
try { $entraDevices = Get-GraphAll "v1.0/devices?`$select=$entraDeviceSelect&`$expand=registeredOwners(`$select=id,userPrincipalName)&`$top=999" }
catch {
    Write-Verbose "Device owners could not be expanded; retrying without owners: $($_.Exception.Message)"
    $entraDevices = Get-GraphAllSafe "v1.0/devices?`$select=$entraDeviceSelect&`$top=999" 'Entra devices'
}

$ManagedByAadId = @{}; foreach ($d in $managedDevices) { if ($d.azureADDeviceId) { $ManagedByAadId[$d.azureADDeviceId] = $d } }
$EntraByDeviceId = @{}; foreach ($d in $entraDevices) { if ($d.deviceId) { $EntraByDeviceId[$d.deviceId] = $d } }
function Get-TrustLabel { param($t) switch ($t) { 'AzureAd' { 'Entra joined' } 'ServerAd' { 'Hybrid Entra joined' } 'Workplace' { 'Entra registered' } default { $(if ($t) { $t } else { 'Not registered' }) } } }

$winEntra   = @($entraDevices | Where-Object { $_.accountEnabled -and $_.operatingSystem -match 'Windows' })
$winManaged = @($winEntra | Where-Object { $ManagedByAadId.ContainsKey($_.deviceId) })
$deviceStats = [ordered]@{
    'Devices enrolled in Intune'                      = @($managedDevices).Count
    'Intune devices by OS'                            = (Join-List (($managedDevices | Group-Object operatingSystem | Sort-Object Name) | ForEach-Object { "$($_.Name): $($_.Count)" }) '; ')
    'Intune devices compliant / noncompliant'         = "$(@($managedDevices | Where-Object { $_.complianceState -eq 'compliant' }).Count) / $(@($managedDevices | Where-Object { $_.complianceState -eq 'noncompliant' }).Count)"
    'Intune devices not synced in 30 days'            = @($managedDevices | Where-Object { $_.lastSyncDateTime -and ([datetime]$_.lastSyncDateTime) -lt (Get-Date).AddDays(-30) }).Count
    'Entra Windows devices (enabled)'                 = $winEntra.Count
    'Entra Windows devices enrolled in Intune'        = $winManaged.Count
    'Entra Windows devices NOT enrolled in Intune'    = ($winEntra.Count - $winManaged.Count)
    'Entra devices by join type'                      = (Join-List (($entraDevices | Where-Object { $_.accountEnabled } | Group-Object { Get-TrustLabel $_.trustType } | Sort-Object Name) | ForEach-Object { "$($_.Name): $($_.Count)" }) '; ')
}

# ----------------------------------------------------------------------------------------------
# 4. Conditional Access
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting Conditional Access policies..." -ForegroundColor Cyan
$caPolicies = Get-GraphAllSafe 'v1.0/identity/conditionalAccess/policies' 'Conditional Access policies'
$namedLocations = Get-GraphAllSafe 'v1.0/identity/conditionalAccess/namedLocations' 'named locations'
$LocMap = @{}; foreach ($l in $namedLocations) { $LocMap[$l.id] = $l.displayName }
$authStrengths = Get-GraphAllSafe 'v1.0/identity/conditionalAccess/authenticationStrength/policies' 'authentication strengths'
$AuthStrengthMap = @{}; foreach ($a in $authStrengths) { $AuthStrengthMap[$a.id] = $a.displayName }

function Resolve-Locations { param($list) if (-not $list) { return 'None' } ; return (Join-List (@($list) | ForEach-Object { if ($LocMap.ContainsKey($_)) { $LocMap[$_] } elseif ($_ -eq 'All') { 'Any location' } elseif ($_ -eq 'AllTrusted') { 'All trusted locations' } else { $_ } }) '; ') }

$caSummaryRows = @()
$caHtml = New-Object System.Text.StringBuilder
foreach ($p in ($caPolicies | Sort-Object displayName)) {
    $stateLabel = switch ($p.state) { 'enabled' { 'On' } 'disabled' { 'Off' } 'enabledForReportingButNotEnforced' { 'Report-only' } default { $p.state } }
    $caSummaryRows += [ordered]@{ 'Policy' = $p.displayName; 'State' = $stateLabel; 'Created' = (Format-DateText $p.createdDateTime); 'Modified' = (Format-DateText $p.modifiedDateTime) }

    $u = $p.conditions.users; $apps = $p.conditions.applications
    $kv = [ordered]@{
        'Name'                            = $p.displayName
        'State'                           = $stateLabel
        'Created'                         = Format-DateText $p.createdDateTime
        'Last modified'                   = Format-DateText $p.modifiedDateTime
        'Users: include'                  = Resolve-CaTargets $u.includeUsers
        'Users: exclude'                  = Resolve-CaTargets $u.excludeUsers
        'Groups: include'                 = Resolve-CaTargets $u.includeGroups
        'Groups: exclude'                 = Resolve-CaTargets $u.excludeGroups
        'Directory roles: include'        = Resolve-CaTargets $u.includeRoles -Roles
        'Directory roles: exclude'        = Resolve-CaTargets $u.excludeRoles -Roles
        'Guest/external users: include'   = $(if ($u.includeGuestsOrExternalUsers) { Join-List $u.includeGuestsOrExternalUsers.guestOrExternalUserTypes } else { 'None' })
        'Cloud apps: include'             = Resolve-CaTargets $apps.includeApplications
        'Cloud apps: exclude'             = Resolve-CaTargets $apps.excludeApplications
        'User actions'                    = $(if ($apps.includeUserActions) { Join-List $apps.includeUserActions } else { 'None' })
        'Authentication context'          = $(if ($apps.includeAuthenticationContextClassReferences) { Join-List $apps.includeAuthenticationContextClassReferences } else { 'None' })
        'Device platforms: include'       = $(if ($p.conditions.platforms) { Join-List $p.conditions.platforms.includePlatforms } else { 'Not configured' })
        'Device platforms: exclude'       = $(if ($p.conditions.platforms) { Join-List $p.conditions.platforms.excludePlatforms } else { 'None' })
        'Locations: include'              = $(if ($p.conditions.locations) { Resolve-Locations $p.conditions.locations.includeLocations } else { 'Not configured' })
        'Locations: exclude'              = $(if ($p.conditions.locations) { Resolve-Locations $p.conditions.locations.excludeLocations } else { 'None' })
        'Client app types'                = Join-List $p.conditions.clientAppTypes
        'Sign-in risk levels'             = $(if ($p.conditions.signInRiskLevels) { Join-List $p.conditions.signInRiskLevels } else { 'Not configured' })
        'User risk levels'                = $(if ($p.conditions.userRiskLevels) { Join-List $p.conditions.userRiskLevels } else { 'Not configured' })
        'Device filter'                   = $(if ($p.conditions.devices.deviceFilter) { "$($p.conditions.devices.deviceFilter.mode): $($p.conditions.devices.deviceFilter.rule)" } else { 'Not configured' })
        'Grant: operator'                 = $(if ($p.grantControls) { $p.grantControls.operator } else { 'Not configured' })
        'Grant: controls'                 = $(if ($p.grantControls) { Join-List $p.grantControls.builtInControls } else { 'None' })
        'Grant: authentication strength'  = $(if ($p.grantControls.authenticationStrength) { $AuthStrengthMap[$p.grantControls.authenticationStrength.id] } else { 'None' })
        'Grant: terms of use'             = $(if ($p.grantControls.termsOfUse) { Join-List $p.grantControls.termsOfUse } else { 'None' })
        'Session: sign-in frequency'      = $(if ($p.sessionControls.signInFrequency.isEnabled) { "$($p.sessionControls.signInFrequency.value) $($p.sessionControls.signInFrequency.type) ($($p.sessionControls.signInFrequency.frequencyInterval))" } else { 'Not configured' })
        'Session: persistent browser'     = $(if ($p.sessionControls.persistentBrowser.isEnabled) { $p.sessionControls.persistentBrowser.mode } else { 'Not configured' })
        'Session: app enforced restrictions' = $(if ($p.sessionControls.applicationEnforcedRestrictions.isEnabled) { 'Enabled' } else { 'Not configured' })
        'Session: Defender for Cloud Apps' = $(if ($p.sessionControls.cloudAppSecurity.isEnabled) { $p.sessionControls.cloudAppSecurity.cloudAppSecurityType } else { 'Not configured' })
        'Session: disable resilience defaults' = $(if ($p.sessionControls.disableResilienceDefaults) { 'Yes' } else { 'No' })
    }
    [void]$caHtml.Append("<div class='card'><h4>$(HtmlEnc $p.displayName)</h4>")
    [void]$caHtml.Append((New-KvTable $kv))
    [void]$caHtml.Append('</div>')
}
if (@($caPolicies).Count -eq 0) { [void]$caHtml.Append("<p class='empty'>No Conditional Access policies exist in this tenant.</p>") }

$namedLocRows = foreach ($l in $namedLocations) {
    [ordered]@{ 'Location' = $l.displayName; 'Type' = ($l.'@odata.type' -replace '#microsoft.graph.', ''); 'Trusted' = $(if ($null -ne $l.isTrusted) { $l.isTrusted } else { '' }); 'Ranges / Countries' = $(if ($l.ipRanges) { Join-List ($l.ipRanges | ForEach-Object { $_.cidrAddress }) } elseif ($l.countriesAndRegions) { Join-List $l.countriesAndRegions } else { '' }); 'Modified' = (Format-DateText $l.modifiedDateTime) }
}

# ----------------------------------------------------------------------------------------------
# 5/6. Enterprise applications, permissions, and consent audit events
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting directory audit events (last $AuditDays days)..." -ForegroundColor Cyan
$since = (Get-Date).AddDays(-$AuditDays).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$auditFilter = "activityDateTime ge $since and (activityDisplayName eq 'Consent to application' or activityDisplayName eq 'Add service principal' or activityDisplayName eq 'Add app role assignment to service principal' or activityDisplayName eq 'Add delegated permission grant' or activityDisplayName eq 'Add app role assignment grant to user')"
$auditEvents = Get-GraphAllSafe "v1.0/auditLogs/directoryAudits?`$filter=$([uri]::EscapeDataString($auditFilter))&`$top=999" 'directory audit log'

$ConsentBySp = @{}      # servicePrincipal id -> list of "date : actor : activity"
$auditRows = @()
foreach ($e in ($auditEvents | Sort-Object activityDateTime -Descending)) {
    $actor = if ($e.initiatedBy.user) { $(if ($e.initiatedBy.user.displayName) { "$($e.initiatedBy.user.displayName) ($($e.initiatedBy.user.userPrincipalName))" } else { [string]$e.initiatedBy.user.userPrincipalName }) } elseif ($e.initiatedBy.app) { "App: $($e.initiatedBy.app.displayName)" } else { 'Unknown' }
    $targetSp = $e.targetResources | Where-Object { $_.type -eq 'ServicePrincipal' } | Select-Object -First 1
    $targetName = if ($targetSp) { $targetSp.displayName } else { ($e.targetResources | Select-Object -First 1).displayName }
    $consentType = ($targetSp.modifiedProperties | Where-Object { $_.displayName -eq 'ConsentContext.IsAdminConsent' } | Select-Object -First 1).newValue
    $auditRows += [ordered]@{ 'Date (local)' = (Format-DateText $e.activityDateTime); 'Activity' = $e.activityDisplayName; 'Performed by' = $actor; 'Application' = $targetName; 'Admin consent' = $(if ($consentType) { $consentType -replace '"', '' } else { '' }); 'Result' = $e.result }
    if ($targetSp -and $targetSp.id) {
        if (-not $ConsentBySp.ContainsKey($targetSp.id)) { $ConsentBySp[$targetSp.id] = New-Object System.Collections.Generic.List[string] }
        $ConsentBySp[$targetSp.id].Add("$(Format-DateText $e.activityDateTime) : $($e.activityDisplayName) by $actor")
    }
}

Write-Host "Collecting enterprise applications and permissions..." -ForegroundColor Cyan
$servicePrincipals = Get-GraphAllSafe 'v1.0/servicePrincipals?$filter=tags/any(t:t eq ''WindowsAzureActiveDirectoryIntegratedApp'')&$select=id,appId,displayName,accountEnabled,servicePrincipalType,appOwnerOrganizationId,publisherName,verifiedPublisher,createdDateTime,appRoleAssignmentRequired,signInAudience,homepage,tags&$top=999' 'enterprise applications'
$MicrosoftTenants = @('f8cdef31-a31e-4b4a-93e4-5f571e91255a', '72f988bf-86f1-41af-91ab-2d7cd011db47')

$script:ResourceRoleCache = @{}
function Get-AppRoleValue {
    param([string]$ResourceId, [string]$AppRoleId)
    if (-not $script:ResourceRoleCache.ContainsKey($ResourceId)) {
        $map = @{}
        try { Confirm-GraphToken; $rsp = Invoke-MgGraphRequest -Method GET -Uri "v1.0/servicePrincipals/$ResourceId`?`$select=appRoles" -OutputType PSObject; foreach ($r in $rsp.appRoles) { $map[$r.id] = $r.value } } catch { Write-Verbose "Could not read appRoles for resource $ResourceId; raw role IDs will be shown." }
        $script:ResourceRoleCache[$ResourceId] = $map
    }
    $m = $script:ResourceRoleCache[$ResourceId]
    if ($m.ContainsKey($AppRoleId)) { return $m[$AppRoleId] } else { return $AppRoleId }
}

$appRows = @()
$i = 0
foreach ($sp in ($servicePrincipals | Sort-Object displayName)) {
    $i++
    Write-Progress -Activity 'Enterprise applications' -Status $sp.displayName -PercentComplete ([int](100 * $i / [Math]::Max(1, @($servicePrincipals).Count)))

    $grants = Get-GraphAllSafe "v1.0/servicePrincipals/$($sp.id)/oauth2PermissionGrants" "delegated grants for $($sp.displayName)"
    $adminScopes = @(); $userScopes = @()
    foreach ($g in $grants) {
        $resName = Resolve-ObjectNames -Ids @($g.resourceId)
        $scopes = ($g.scope -split ' ' | Where-Object { $_ }) | ForEach-Object { "$resName/$_" }
        if ($g.consentType -eq 'AllPrincipals') { $adminScopes += $scopes } else { $userScopes += $scopes }
    }
    $roleAsg = Get-GraphAllSafe "v1.0/servicePrincipals/$($sp.id)/appRoleAssignments" "application permissions for $($sp.displayName)"
    $appPerms = foreach ($r in $roleAsg) { "$($r.resourceDisplayName)/$(Get-AppRoleValue -ResourceId $r.resourceId -AppRoleId $r.appRoleId)" }

    $assignedTo = Get-GraphAllSafe "v1.0/servicePrincipals/$($sp.id)/appRoleAssignedTo?`$top=999" "user assignments for $($sp.displayName)"
    $assignedCount = @($assignedTo | Where-Object { $_.principalType -in @('User', 'Group') }).Count

    $publisher = if ($MicrosoftTenants -contains $sp.appOwnerOrganizationId) { 'Microsoft' }
                 elseif ($sp.appOwnerOrganizationId -eq $org.id) { $(if ($sp.publisherName) { "This tenant ($($sp.publisherName))" } else { 'This tenant' }) }
                 elseif ($sp.verifiedPublisher.displayName) { "$($sp.verifiedPublisher.displayName) (verified)" }
                 elseif ($sp.publisherName) { "$($sp.publisherName) (unverified)" }
                 else { 'Unknown' }

    $approval = if ($ConsentBySp.ContainsKey($sp.id)) { Join-List $ConsentBySp[$sp.id] ' | ' }
                elseif ($adminScopes.Count -or @($appPerms).Count) { 'Not in audit log retention' }
                else { 'No admin consent recorded' }

    $appRows += [ordered]@{
        'Application'              = $sp.displayName
        'Publisher'                = $publisher
        'Type'                     = $sp.servicePrincipalType
        'Enabled for sign-in'      = $sp.accountEnabled
        'Assignment required'      = $sp.appRoleAssignmentRequired
        'Users/groups assigned'    = $assignedCount
        'Added to tenant'          = (Format-DateText $sp.createdDateTime)
        'Approved by / when'       = $approval
        'Admin-consented delegated permissions' = $(if ($adminScopes.Count) { Join-List ($adminScopes | Sort-Object -Unique) '; ' } else { 'None' })
        'User-consented delegated permissions'  = $(if ($userScopes.Count) { "$(@($userScopes | Sort-Object -Unique).Count) grant(s): " + (Join-List ($userScopes | Sort-Object -Unique) '; ') } else { 'None' })
        'Application permissions'  = $(if (@($appPerms).Count) { Join-List ($appPerms | Sort-Object -Unique) '; ' } else { 'None' })
        'App ID'                   = $sp.appId
    }
}
Write-Progress -Activity 'Enterprise applications' -Completed

$consentPolicy = (Get-GraphAllSafe 'v1.0/policies/authorizationPolicy' 'authorization policy') | Select-Object -First 1
$userConsentSetting = if ($consentPolicy) {
    $perms = @($consentPolicy.defaultUserRolePermissions.permissionGrantPoliciesAssigned)
    if ($perms.Count -eq 0) { 'Do not allow user consent' }
    elseif ($perms -match 'microsoft-user-default-low') { 'Allow user consent for apps from verified publishers, for selected (low impact) permissions' }
    elseif ($perms -match 'microsoft-user-default-legacy') { 'Allow user consent for all apps' }
    elseif ($perms -match 'microsoft-user-default-recommended') { 'Microsoft recommended: allow user consent for apps from verified publishers or this tenant, for selected (low impact) permissions' }
    elseif ($perms -match 'ManagePermissionGrantsForSelf') { 'Custom user consent policy: ' + (Join-List ($perms | Where-Object { $_ -match 'ManagePermissionGrantsForSelf' } | ForEach-Object { $_ -replace '^ManagePermissionGrantsForSelf\.', '' })) }
    else { 'Do not allow user consent (only resource-owner / team policies assigned)' }
} else { 'Could not read' }
$adminConsentWorkflow = (Get-GraphAllSafe 'v1.0/policies/adminConsentRequestPolicy' 'admin consent request policy') | Select-Object -First 1
$appGovKv = [ordered]@{
    'User consent setting'                          = $userConsentSetting
    'Users can register applications'               = $(if ($null -ne $consentPolicy.defaultUserRolePermissions.allowedToCreateApps) { $consentPolicy.defaultUserRolePermissions.allowedToCreateApps } else { '' })
    'Admin consent request workflow'                = $(if ($adminConsentWorkflow.isEnabled) { 'Enabled' } else { 'Disabled' })
    'Admin consent reviewers'                       = $(if ($adminConsentWorkflow.reviewers) { Join-List (Resolve-ObjectNames ($adminConsentWorkflow.reviewers | ForEach-Object { ($_.query -split '/')[-1] })) '; ' } else { 'None' })
    'Enterprise applications in tenant'             = @($servicePrincipals).Count
    'Applications with admin-consented permissions' = @($appRows | Where-Object { $_.'Admin-consented delegated permissions' -ne 'None' -or $_.'Application permissions' -ne 'None' }).Count
    'Applications from unverified or unknown publishers' = @($appRows | Where-Object { $_.Publisher -match 'unverified|Unknown' }).Count
}

# ----------------------------------------------------------------------------------------------
# 7. Users
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting users..." -ForegroundColor Cyan
$userSelect = 'id,displayName,userPrincipalName,accountEnabled,userType,createdDateTime,jobTitle,department,assignedLicenses,onPremisesSyncEnabled,onPremisesDomainName,onPremisesLastSyncDateTime,signInActivity'
$users = @()
try { $users = Get-GraphAll "v1.0/users?`$select=$userSelect&`$top=999" }
catch {
    Write-Warning "signInActivity not available (requires Entra ID P1). Retrying without it."
    $users = Get-GraphAllSafe "v1.0/users?`$select=$($userSelect -replace ',signInActivity','')&`$top=999" 'users'
}
$roles = Get-GraphAllSafe 'v1.0/directoryRoles?$expand=members' 'directory roles'
$AdminIds = @{}; foreach ($r in $roles) { foreach ($m in @($r.members)) { if ($m.id) { $AdminIds[$m.id] = $true } } }
$regDetails = Get-GraphAllSafe 'v1.0/reports/authenticationMethods/userRegistrationDetails?$top=999' 'MFA registration details'
$RegMap = @{}; foreach ($r in $regDetails) { $RegMap[$r.id] = $r }
$script:RegFromMethods = $false
if (-not @($regDetails).Count -and @($users).Count) {
    # Without Entra ID P1 the registration report is refused; build the same facts from each user's authentication methods
    # (UserAuthenticationMethod.Read.All) with Graph $batch requests of 20.
    $regTargets = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' } | Select-Object -First 3000)
    Write-Host "Registration report unavailable (Entra ID P1); reading authentication methods for $($regTargets.Count) accounts instead..." -ForegroundColor Cyan
    $MethodTypeMap = @{ 'fido2AuthenticationMethod' = 'fido2SecurityKey'; 'microsoftAuthenticatorAuthenticationMethod' = 'microsoftAuthenticatorPush'; 'windowsHelloForBusinessAuthenticationMethod' = 'windowsHelloForBusiness'; 'softwareOathAuthenticationMethod' = 'softwareOneTimePasscode'; 'hardwareOathAuthenticationMethod' = 'hardwareOneTimePasscode'; 'emailAuthenticationMethod' = 'email'; 'temporaryAccessPassAuthenticationMethod' = 'temporaryAccessPass'; 'platformCredentialAuthenticationMethod' = 'passKeyDeviceBound'; 'x509CertificateAuthenticationMethod' = 'x509Certificate'; 'externalAuthenticationMethod' = 'externalAuthMethod' }
    $regFallback = New-Object System.Collections.Generic.List[object]; $regBatchFailed = $false
    for ($bi = 0; $bi -lt $regTargets.Count -and -not $regBatchFailed; $bi += 20) {
        $chunk = @($regTargets[$bi..([Math]::Min($bi + 19, $regTargets.Count - 1))])
        $reqs = @(); $rn = 0; foreach ($u in $chunk) { $rn++; $reqs += @{ id = "$rn"; method = 'GET'; url = "/users/$($u.id)/authentication/methods" } }
        try {
            Confirm-GraphToken
            $resp = Invoke-MgGraphRequest -Method POST -Uri 'v1.0/$batch' -Body (@{ requests = $reqs } | ConvertTo-Json -Depth 5) -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            $rn = 0
            foreach ($u in $chunk) {
                $rn++
                $r = @($resp.responses | Where-Object { [string]$_.id -eq "$rn" }) | Select-Object -First 1
                if (-not $r -or [int]$r.status -ge 400) {
                    if ($r -and [int]$r.status -in @(401, 403)) { $regBatchFailed = $true; $script:GraphFailures['MFA registration details'] = "authentication methods per user returned HTTP $($r.status): UserAuthenticationMethod.Read.All not consented, or the account lacks an authentication administrator / Global Reader role." }
                    continue
                }
                $methods = @()
                foreach ($m in @($r.body.value)) {
                    $mt = ([string]$m.'@odata.type') -replace '^#microsoft\.graph\.', ''
                    if ($mt -eq 'passwordAuthenticationMethod') { continue }
                    if ($mt -eq 'phoneAuthenticationMethod') { $methods += $(switch ([string]$m.phoneType) { 'office' { 'officePhone' } 'alternateMobile' { 'alternateMobilePhone' } default { 'mobilePhone' } }) }
                    elseif ($mt -eq 'fido2AuthenticationMethod' -and ([string]$m.model) -match 'Authenticator|Microsoft') { $methods += 'passKeyDeviceBoundAuthenticator' }
                    elseif ($MethodTypeMap.ContainsKey($mt)) { $methods += $MethodTypeMap[$mt] }
                    elseif ($mt) { $methods += $mt }
                }
                $methods = @($methods | Select-Object -Unique)
                $mfaReg = @($methods | Where-Object { $_ -notin @('email', 'securityQuestion', 'temporaryAccessPass') }).Count -gt 0
                $pwdless = @($methods | Where-Object { $_ -in @('fido2SecurityKey', 'passKeyDeviceBoundAuthenticator', 'passKeyDeviceBound', 'windowsHelloForBusiness', 'x509Certificate') }).Count -gt 0
                $regFallback.Add([pscustomobject]@{ id = $u.id; userPrincipalName = $u.userPrincipalName; userDisplayName = $u.displayName; isMfaRegistered = $mfaReg; isMfaCapable = $mfaReg; isPasswordlessCapable = $pwdless; methodsRegistered = $methods; defaultMfaMethod = 'none'; isAdmin = $AdminIds.ContainsKey($u.id) })
            }
        } catch { $regBatchFailed = $true; $script:GraphFailures['MFA registration details'] = "authentication methods per user: $(Get-GraphErrorText $_)"; Write-Warning "Could not read authentication methods per user: $(Get-GraphErrorText $_)" }
    }
    if (-not $regBatchFailed -and $regFallback.Count) {
        $regDetails = $regFallback.ToArray(); $script:GraphFailures.Remove('MFA registration details'); $script:RegFromMethods = $true
        $RegMap = @{}; foreach ($r in $regDetails) { $RegMap[$r.id] = $r }
        Write-Host "  Registration facts built for $($regDetails.Count) accounts from their authentication methods (Entra ID Free fallback)." -ForegroundColor DarkGray
    }
}
$UserById = @{}; foreach ($u in $users) { $UserById[$u.id] = $u }

# One device list: every Entra ID device object joined with its Intune record (plus Intune devices with no Entra object).
$StaleDays = 90
$DevicesByUser = @{}
function Add-DeviceOwner { param([string]$Upn) if ($Upn) { $k = $Upn.ToLower(); if ($DevicesByUser.ContainsKey($k)) { $DevicesByUser[$k]++ } else { $DevicesByUser[$k] = 1 } } }
function Get-LastLoggedOnUser { param($md)
    if (-not $md -or -not $md.usersLoggedOn) { return '' }
    $lu = @($md.usersLoggedOn | Sort-Object lastLogOnDateTime | Select-Object -Last 1)[0]
    if ($lu.userId -and $UserById.ContainsKey($lu.userId)) { return $UserById[$lu.userId].userPrincipalName } ; return [string]$lu.userId
}
$allDeviceRows = New-Object System.Collections.Generic.List[object]
foreach ($ad in ($entraDevices | Sort-Object displayName)) {
    $md = if ($ad.deviceId -and $ManagedByAadId.ContainsKey($ad.deviceId)) { $ManagedByAadId[$ad.deviceId] } else { $null }
    $owners = @(@($ad.registeredOwners) | ForEach-Object { $_.userPrincipalName } | Where-Object { $_ })
    if (-not $owners.Count -and $md.userPrincipalName) { $owners = @($md.userPrincipalName) }
    foreach ($o in ($owners | Select-Object -Unique)) { Add-DeviceOwner $o }
    $lastUser = Get-LastLoggedOnUser $md
    $acts = @($ad.approximateLastSignInDateTime, $md.lastSyncDateTime) | Where-Object { $_ } | ForEach-Object { [datetime]$_ }
    $lastAct = if ($acts.Count) { ($acts | Sort-Object | Select-Object -Last 1) } else { $null }
    $isStale = (-not $lastAct) -or ($lastAct -lt (Get-Date).AddDays(-$StaleDays))
    $nonCompliant = ($md -and $md.complianceState -eq 'noncompliant') -or ($ad.isCompliant -eq $false)
    $tags = @()
    if ($isStale) { $tags += 'stale' }; if (-not $md) { $tags += 'notintune' }; if ($nonCompliant) { $tags += 'noncompliant' }
    if ($ad.operatingSystem -match 'Windows') { $tags += 'windows' }; if (-not $ad.accountEnabled) { $tags += 'disabled' }
    if (($md -and $md.managedDeviceOwnerType -eq 'personal') -or $ad.trustType -eq 'Workplace') { $tags += 'personal' }
    if ($md -and $md.isEncrypted -eq $false) { $tags += 'notencrypted' }
    $allDeviceRows.Add([ordered]@{
        'Device'            = $ad.displayName
        'OS'                = ("$($ad.operatingSystem) $($ad.operatingSystemVersion)").Trim()
        'Join type'         = Get-TrustLabel $ad.trustType
        'Enabled'           = $ad.accountEnabled
        'Intune'            = $(if ($md) { 'Enrolled' } else { 'Not enrolled' })
        'Compliance'        = $(if ($md) { $md.complianceState } elseif ($null -ne $ad.isCompliant) { $(if ($ad.isCompliant) { 'compliant' } else { 'noncompliant' }) } else { 'Not evaluated' })
        'Owner / primary user' = (Join-List ($owners | Select-Object -Unique) '; ')
        'Last user'         = $lastUser
        'Last activity'     = $(if ($lastAct) { $lastAct.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { 'Never' })
        'Enrolled'          = $(if ($md) { Format-DateText $md.enrolledDateTime } else { '' })
        'Encrypted'         = $(if ($md) { $md.isEncrypted } else { '' })
        'Manufacturer / model' = (("$(if ($md) { $md.manufacturer } else { $ad.manufacturer }) $(if ($md) { $md.model } else { $ad.model })").Trim())
        'Serial'            = $(if ($md) { $md.serialNumber } else { '' })
        '_tags'             = ($tags -join ' ')
    })
}
foreach ($md in ($managedDevices | Where-Object { -not ($_.azureADDeviceId -and $EntraByDeviceId.ContainsKey($_.azureADDeviceId)) } | Sort-Object deviceName)) {
    if ($md.userPrincipalName) { Add-DeviceOwner $md.userPrincipalName }
    $lastAct = if ($md.lastSyncDateTime) { [datetime]$md.lastSyncDateTime } else { $null }
    $tags = @(); if ((-not $lastAct) -or ($lastAct -lt (Get-Date).AddDays(-$StaleDays))) { $tags += 'stale' }
    if ($md.complianceState -eq 'noncompliant') { $tags += 'noncompliant' }; if ($md.operatingSystem -match 'Windows') { $tags += 'windows' }
    if ($md.managedDeviceOwnerType -eq 'personal') { $tags += 'personal' }; if ($md.isEncrypted -eq $false) { $tags += 'notencrypted' }
    $allDeviceRows.Add([ordered]@{
        'Device' = $md.deviceName; 'OS' = ("$($md.operatingSystem) $($md.osVersion)").Trim(); 'Join type' = 'Intune only (no Entra object)'; 'Enabled' = ''
        'Intune' = 'Enrolled'; 'Compliance' = $md.complianceState; 'Owner / primary user' = $md.userPrincipalName; 'Last user' = (Get-LastLoggedOnUser $md)
        'Last activity' = $(if ($lastAct) { $lastAct.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { 'Never' }); 'Enrolled' = (Format-DateText $md.enrolledDateTime)
        'Encrypted' = $md.isEncrypted; 'Manufacturer / model' = (("$($md.manufacturer) $($md.model)").Trim()); 'Serial' = $md.serialNumber; '_tags' = ($tags -join ' ')
    })
}
$staleDeviceCount = @($allDeviceRows | Where-Object { " $($_['_tags']) " -like '* stale *' }).Count
$deviceStats['All devices (Entra ID and Intune)'] = $allDeviceRows.Count
$deviceStats["Stale devices (no activity in $StaleDays days)"] = $staleDeviceCount
$DeviceFilters = @(
    @{ Label = 'Stale'; Tag = 'stale' }, @{ Label = 'Not in Intune'; Tag = 'notintune' }, @{ Label = 'Noncompliant'; Tag = 'noncompliant' },
    @{ Label = 'Not encrypted'; Tag = 'notencrypted' }, @{ Label = 'Personal'; Tag = 'personal' }, @{ Label = 'Windows'; Tag = 'windows' }, @{ Label = 'Disabled'; Tag = 'disabled' }
)

$signInAvailable = [bool]@($users | Where-Object { $_.signInActivity } | Select-Object -First 1).Count
$userRows = foreach ($u in ($users | Sort-Object displayName)) {
    $reg = $RegMap[$u.id]
    $lic = ($u.assignedLicenses | ForEach-Object { if ($SkuMap.ContainsKey($_.skuId)) { $SkuMap[$_.skuId] } else { $_.skuId } }) | Sort-Object -Unique
    $lastTimes = @($u.signInActivity.lastSignInDateTime, $u.signInActivity.lastNonInteractiveSignInDateTime) | Where-Object { $_ } | ForEach-Object { [datetime]$_ }
    $lastAny = if ($lastTimes.Count) { ($lastTimes | Sort-Object | Select-Object -Last 1) } else { $null }
    $isMember = $u.accountEnabled -and $u.userType -eq 'Member'
    $isStale = $signInAvailable -and $isMember -and ((-not $lastAny) -or ($lastAny -lt (Get-Date).AddDays(-$StaleDays)))
    $isAdmin = $AdminIds.ContainsKey($u.id) -or ($reg -and $reg.isAdmin)
    $noMfa   = $isMember -and $reg -and -not $reg.isMfaRegistered
    $tags = @()
    if ($isStale) { $tags += 'stale' }; if ($noMfa) { $tags += 'nomfa' }; if ($isAdmin) { $tags += 'admin' }
    if (-not $u.accountEnabled) { $tags += 'disabled' }; if ($u.userType -eq 'Guest') { $tags += 'guest' }
    if (@($lic).Count) { $tags += 'licensed' }; if ($u.onPremisesSyncEnabled) { $tags += 'synced' }
    $devKey = if ($u.userPrincipalName) { $u.userPrincipalName.ToLower() } else { '' }
    [ordered]@{
        'Display name'        = $u.displayName
        'User principal name' = $u.userPrincipalName
        'Enabled'             = $u.accountEnabled
        'Department'          = $u.department
        'Created'             = (Format-DateText $u.createdDateTime)
        'Last sign-in'        = $(if ($signInAvailable) { $(if ($lastAny) { $lastAny.ToLocalTime().ToString('yyyy-MM-dd HH:mm') } else { 'Never' }) } else { 'Not available' })
        'Licenses'            = (Join-List $lic '; ')
        'MFA registered'      = $(if ($reg) { $(if ($reg.isMfaRegistered) { 'Yes' } else { 'No' }) } else { 'Not available' })
        'MFA methods'         = $(if ($reg) { Join-List $reg.methodsRegistered '; ' } else { '' })
        'Default method'      = $(if ($reg) { $reg.defaultMfaMethod } else { '' })
        'Admin'               = $(if ($isAdmin) { 'Yes' } else { 'No' })
        'Devices'             = $(if ($devKey -and $DevicesByUser.ContainsKey($devKey)) { $DevicesByUser[$devKey] } else { 0 })
        '_tags'               = ($tags -join ' ')
    }
}
$UserFilters = @(@{ Label = 'No MFA'; Tag = 'nomfa' }, @{ Label = 'Admin'; Tag = 'admin' }, @{ Label = 'Guests'; Tag = 'guest' }, @{ Label = 'Disabled'; Tag = 'disabled' }, @{ Label = 'Licensed'; Tag = 'licensed' })
if ($signInAvailable) { $UserFilters = @(@{ Label = "Stale ($StaleDays days)"; Tag = 'stale' }) + $UserFilters }
$userStats = [ordered]@{
    'Total users'                        = @($users).Count
    'Enabled members'                    = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' }).Count
    'Disabled accounts'                  = @($users | Where-Object { -not $_.accountEnabled }).Count
    'Guest accounts'                     = @($users | Where-Object { $_.userType -eq 'Guest' }).Count
    'Enabled members with no MFA registered' = $(if (@($regDetails).Count) { @($userRows | Where-Object { " $($_['_tags']) " -like '* nomfa *' }).Count } else { 'Not available' })
    'Enabled members with no sign-in in 90 days' = $(if ($users | Where-Object { $_.signInActivity }) { @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' -and ( -not $_.signInActivity.lastSignInDateTime -or ([datetime]$_.signInActivity.lastSignInDateTime) -lt (Get-Date).AddDays(-90) ) }).Count } else { 'Not available' })
}

# ----------------------------------------------------------------------------------------------
# 8. Privileged roles
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting directory role members..." -ForegroundColor Cyan
$roleRows = @()
$privHtml = New-Object System.Text.StringBuilder
$RolePriority = @('Global Administrator', 'Privileged Role Administrator', 'Privileged Authentication Administrator', 'Security Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'Intune Administrator', 'Conditional Access Administrator', 'Application Administrator', 'Cloud Application Administrator', 'User Administrator', 'Authentication Administrator', 'Helpdesk Administrator', 'Global Reader')
$activeRoles = @($roles | Where-Object { @($_.members).Count -gt 0 } | Sort-Object @{ Expression = { $idx = $RolePriority.IndexOf($_.displayName); if ($idx -lt 0) { 999 } else { $idx } } }, displayName)
$pimRules = Get-GraphAllSafe 'v1.0/roleManagement/directory/roleEligibilityScheduleInstances?$expand=principal' 'PIM eligible assignments'
foreach ($r in $activeRoles) {
    $memberRows = foreach ($m in ($r.members | Sort-Object displayName)) {
        $reg = $RegMap[$m.id]
        $mu = $UserById[$m.id]
        $row = [ordered]@{
            'Member'        = $m.displayName
            'UPN / Type'    = $(if ($m.userPrincipalName) { $m.userPrincipalName } else { ($m.'@odata.type' -replace '#microsoft.graph.', '') })
            'Enabled'       = $(if ($null -ne $m.accountEnabled) { $m.accountEnabled } else { '' })
            'Assignment'    = 'Active (permanent)'
            'MFA registered' = $(if ($reg) { $(if ($reg.isMfaRegistered) { 'Yes' } else { 'No' }) } else { '' })
            'Last sign-in'  = $(if ($mu) { Format-DateText $mu.signInActivity.lastSignInDateTime } else { '' })
            '_issues'       = $(if ($reg -and -not $reg.isMfaRegistered) { @{ 'MFA registered' = 'admin-no-mfa' } } else { $null })
        }
        $full = [ordered]@{ 'Role' = $r.displayName }
        foreach ($k in $row.Keys) { $full[$k] = $row[$k] }
        $roleRows += $full
        $row
    }
    $gaIcon = if ($r.displayName -eq 'Global Administrator') { $(if (@($r.members).Count -gt 4) { New-InfoIcon 'ga-count-high' } elseif (@($r.members).Count -eq 1) { New-InfoIcon 'ga-single' } else { '' }) } else { '' }
    [void]$privHtml.Append("<h3>$(HtmlEnc $r.displayName) <span class='count'>($(@($r.members).Count))</span>$gaIcon</h3>")
    [void]$privHtml.Append((New-HtmlTable $memberRows -Columns 'Member', 'UPN / Type', 'Enabled', 'Assignment', 'MFA registered', 'Last sign-in'))
}
if (@($pimRules).Count) {
    $pimRows = foreach ($e in $pimRules) { [ordered]@{ 'Role' = (Resolve-RoleTemplateNames -Ids @($e.roleDefinitionId)); 'Principal' = $(if ($e.principal.userPrincipalName) { "$($e.principal.displayName) ($($e.principal.userPrincipalName))" } else { $e.principal.displayName }); 'Eligible from' = (Format-DateText $e.startDateTime); 'Eligible until' = $(if ($e.endDateTime) { Format-DateText $e.endDateTime } else { 'Permanent' }) } }
    [void]$privHtml.Append("<h3>PIM eligible assignments (not currently active)</h3>")
    [void]$privHtml.Append((New-HtmlTable $pimRows -Columns 'Role', 'Principal', 'Eligible from', 'Eligible until'))
}
$privStats = [ordered]@{
    'Roles with active members'              = @($activeRoles).Count
    'Total active role assignments'          = @($roleRows).Count
    'Global Administrators'                  = @($roleRows | Where-Object { $_.Role -eq 'Global Administrator' }).Count
    'Privileged accounts without MFA registered' = $(if (@($regDetails).Count) { @($roleRows | Where-Object { $_.'MFA registered' -eq 'No' }).Count } else { 'Not available' })
    'PIM eligible assignments'               = @($pimRules).Count
}

# ----------------------------------------------------------------------------------------------
# 9. Authentication methods, TAP, security defaults, device registration policy
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting authentication method policy and tenant security settings..." -ForegroundColor Cyan
$secDefaults = (Get-GraphAllSafe 'v1.0/policies/identitySecurityDefaultsEnforcementPolicy' 'security defaults') | Select-Object -First 1
$authMethodsPolicy = (Get-GraphAllSafe 'v1.0/policies/authenticationMethodsPolicy' 'authentication methods policy') | Select-Object -First 1
$authMethods = @($authMethodsPolicy.authenticationMethodConfigurations | Where-Object { $_ })
$devRegPolicy = (Get-GraphAllSafe 'beta/policies/deviceRegistrationPolicy' 'device registration policy') | Select-Object -First 1

function Resolve-AuthTargets { param($targets)
    if (-not $targets -or @($targets).Count -eq 0) { return 'None' }
    return (Join-List (@($targets) | ForEach-Object { if ($_.id -eq 'all_users') { 'All users' } else { Resolve-ObjectNames -Ids @($_.id) } }) '; ')
}
$authMethodRows = foreach ($m in $authMethods) {
    $extra = @()
    foreach ($prop in $m.PSObject.Properties) {
        if ($prop.Name -in @('id', '@odata.type', 'state', 'includeTargets', 'excludeTargets', 'includeTargets@odata.context', 'excludeTargets@odata.context')) { continue }
        $val = Format-SettingValue $prop.Value
        if ($val -notin @('Not configured', 'None')) { $extra += "$(Format-SettingName $prop.Name)=$val" }
    }
    [ordered]@{ 'Method' = $m.id; 'State' = $m.state; 'Included targets' = (Resolve-AuthTargets $m.includeTargets); 'Excluded targets' = (Resolve-AuthTargets $m.excludeTargets); 'Settings' = (Join-List $extra '; ') }
}
$tap = $authMethods | Where-Object { $_.id -eq 'TemporaryAccessPass' } | Select-Object -First 1
$tapKv = if ($tap) {
    [ordered]@{
        'Temporary Access Pass state' = $tap.state
        'Enabled for'                 = Resolve-AuthTargets $tap.includeTargets
        'Default lifetime (minutes)'  = $tap.defaultLifetimeInMinutes
        'Minimum / maximum lifetime (minutes)' = "$($tap.minimumLifetimeInMinutes) / $($tap.maximumLifetimeInMinutes)"
        'One-time use only'           = $tap.isUsableOnce
        'Default length'              = $tap.defaultLength
    }
} else { [ordered]@{ 'Temporary Access Pass state' = "Could not read. $(Get-GraphFailure 'authentication methods policy')" } }

$devRegKv = if ($devRegPolicy) {
    [ordered]@{
        'Users may join devices to Entra ID'           = $devRegPolicy.azureADJoin.allowedToJoin.'@odata.type' -replace '#microsoft.graph.', '' -replace 'DeviceRegistrationMembership', ''
        'Users may register devices'                    = $devRegPolicy.azureADRegistration.allowedToRegister.'@odata.type' -replace '#microsoft.graph.', '' -replace 'DeviceRegistrationMembership', ''
        'MFA required to join or register devices'      = $devRegPolicy.multiFactorAuthConfiguration
        'Maximum devices per user'                      = $devRegPolicy.userDeviceQuota
        'Local admin: global admins added'              = $devRegPolicy.localAdminPassword.isEnabled
    }
} else { [ordered]@{ 'Device registration policy' = 'Could not read' } }

# ----------------------------------------------------------------------------------------------
# 10. Conditional Access evaluation: MFA, legacy auth, device-based access
# ----------------------------------------------------------------------------------------------
function Test-CaAllUsers { param($p) return (@($p.conditions.users.includeUsers) -contains 'All') }
function Test-CaAllApps  { param($p) return (@($p.conditions.applications.includeApplications) -contains 'All') }
function Get-CaScopeText { param($p)
    $inc = Resolve-CaTargets $p.conditions.users.includeUsers
    if ($p.conditions.users.includeGroups) { $inc += '; groups: ' + (Resolve-CaTargets $p.conditions.users.includeGroups) }
    if ($p.conditions.users.includeRoles)  { $inc += '; roles: ' + (Resolve-CaTargets $p.conditions.users.includeRoles -Roles) }
    $exc = @()
    if ($p.conditions.users.excludeUsers)  { $exc += Resolve-CaTargets $p.conditions.users.excludeUsers }
    if ($p.conditions.users.excludeGroups) { $exc += Resolve-CaTargets $p.conditions.users.excludeGroups }
    if ($p.conditions.users.excludeRoles)  { $exc += Resolve-CaTargets $p.conditions.users.excludeRoles -Roles }
    return "Include: $inc | Exclude: $(if ($exc.Count) { Join-List $exc '; ' } else { 'None' })"
}
$enabledCa = @($caPolicies | Where-Object { $_.state -eq 'enabled' })
$reportOnlyCa = @($caPolicies | Where-Object { $_.state -eq 'enabledForReportingButNotEnforced' })

$mfaPolicies = @($caPolicies | Where-Object { ($_.grantControls.builtInControls -contains 'mfa') -or $_.grantControls.authenticationStrength })
$mfaRows = foreach ($p in $mfaPolicies) { [ordered]@{ 'Policy' = $p.displayName; 'State' = $p.state; 'Scope' = (Get-CaScopeText $p); 'Apps' = (Resolve-CaTargets $p.conditions.applications.includeApplications); 'Control' = $(if ($p.grantControls.authenticationStrength) { "Authentication strength: $($AuthStrengthMap[$p.grantControls.authenticationStrength.id])" } else { 'Require MFA' }) } }
$mfaAllUsersEnforced = @($mfaPolicies | Where-Object { $_.state -eq 'enabled' -and (Test-CaAllUsers $_) -and (Test-CaAllApps $_) })

$legacyPolicies = @($caPolicies | Where-Object { ($_.conditions.clientAppTypes -contains 'other' -or $_.conditions.clientAppTypes -contains 'exchangeActiveSync') -and ($_.grantControls.builtInControls -contains 'block') })
$legacyRows = foreach ($p in $legacyPolicies) { [ordered]@{ 'Policy' = $p.displayName; 'State' = $p.state; 'Scope' = (Get-CaScopeText $p); 'Client app types' = (Join-List $p.conditions.clientAppTypes); 'Control' = 'Block' } }
$legacyBlockedAll = @($legacyPolicies | Where-Object { $_.state -eq 'enabled' -and (Test-CaAllUsers $_) -and (Test-CaAllApps $_) })

$devicePolicies = @($caPolicies | Where-Object { ($_.grantControls.builtInControls -contains 'compliantDevice') -or ($_.grantControls.builtInControls -contains 'domainJoinedDevice') -or ($_.conditions.devices.deviceFilter.rule -match 'trustType|isCompliant') })
$deviceCaRows = foreach ($p in $devicePolicies) {
    $ctrl = @()
    if ($p.grantControls.builtInControls -contains 'compliantDevice')    { $ctrl += 'Require compliant device' }
    if ($p.grantControls.builtInControls -contains 'domainJoinedDevice') { $ctrl += 'Require Entra hybrid joined device' }
    if ($p.grantControls.builtInControls -contains 'block')              { $ctrl += 'Block' }
    [ordered]@{ 'Policy' = $p.displayName; 'State' = $p.state; 'Scope' = (Get-CaScopeText $p); 'Apps' = (Resolve-CaTargets $p.conditions.applications.includeApplications); 'Platforms' = $(if ($p.conditions.platforms) { Join-List $p.conditions.platforms.includePlatforms } else { 'Any' }); 'Device filter' = $(if ($p.conditions.devices.deviceFilter) { "$($p.conditions.devices.deviceFilter.mode): $($p.conditions.devices.deviceFilter.rule)" } else { '' }); 'Controls' = (Join-List $ctrl " $($p.grantControls.operator) ") }
}
$deviceEnforcedAll = @($devicePolicies | Where-Object { $_.state -eq 'enabled' -and (Test-CaAllUsers $_) })

# MFA registration coverage
$enabledMembers = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' })
$noMfaRows = @()
if (@($regDetails).Count) {
    $noMfaRows = foreach ($u in ($enabledMembers | Where-Object { -not $RegMap[$_.id] -or -not $RegMap[$_.id].isMfaRegistered } | Sort-Object displayName)) {
        [ordered]@{ 'User' = $u.displayName; 'UPN' = $u.userPrincipalName; 'Licensed' = $(if (@($u.assignedLicenses).Count) { 'Yes' } else { 'No' }); 'Last sign-in' = (Format-DateText $u.signInActivity.lastSignInDateTime); 'Methods registered' = $(if ($RegMap[$u.id]) { Join-List $RegMap[$u.id].methodsRegistered '; ' } else { '' }) }
    }
}
$mfaRegistered = @($enabledMembers | Where-Object { $RegMap[$_.id] -and $RegMap[$_.id].isMfaRegistered }).Count

# ---- MFA product in use and passkey (phishing-resistant) rollout, from the registration report and the methods policy
$MethodLabels = [ordered]@{
    'microsoftAuthenticatorPasswordless' = 'Microsoft Authenticator passwordless (phone sign-in)'; 'passKeyDeviceBoundAuthenticator' = 'Passkey in Microsoft Authenticator'
    'passKeyDeviceBound' = 'Passkey (device-bound)'; 'passKeyDeviceBoundWindowsHello' = 'Passkey (Windows Hello)'; 'fido2SecurityKey' = 'FIDO2 security key'
    'windowsHelloForBusiness' = 'Windows Hello for Business'; 'microsoftAuthenticatorPush' = 'Microsoft Authenticator (push / number matching)'
    'softwareOneTimePasscode' = 'Software OATH token (authenticator app code)'; 'hardwareOneTimePasscode' = 'Hardware OATH token'
    'mobilePhone' = 'SMS or call to mobile phone'; 'alternateMobilePhone' = 'SMS or call to alternate phone'; 'officePhone' = 'Call to office phone'
    'email' = 'Email (self-service password reset only)'; 'securityQuestion' = 'Security questions (self-service password reset only)'
    'temporaryAccessPass' = 'Temporary Access Pass'; 'x509Certificate' = 'Certificate-based authentication'; 'externalAuthMethod' = 'External (third-party) authentication method'
}
$PhishingResistantMethods = @('passKeyDeviceBoundAuthenticator', 'passKeyDeviceBound', 'passKeyDeviceBoundWindowsHello', 'fido2SecurityKey', 'windowsHelloForBusiness', 'x509Certificate')
$WeakMethods = @('mobilePhone', 'alternateMobilePhone', 'officePhone')
$regEnabled = @($enabledMembers | ForEach-Object { $RegMap[$_.id] } | Where-Object { $_ })
$methodCounts = @{}
foreach ($rd in $regEnabled) { foreach ($m in @($rd.methodsRegistered)) { if ($m) { if (-not $methodCounts.ContainsKey($m)) { $methodCounts[$m] = 0 }; $methodCounts[$m]++ } } }
$methodCountRows = @($methodCounts.GetEnumerator() | Sort-Object { $_.Value } -Descending | ForEach-Object {
    $kind = if ($_.Key -in $PhishingResistantMethods) { 'Phishing-resistant' } elseif ($_.Key -in $WeakMethods) { 'Weak (phone-based)' } elseif ($_.Key -in @('email', 'securityQuestion')) { 'Recovery only' } else { 'Standard MFA' }
    [ordered]@{ 'Method' = $(if ($MethodLabels.Contains($_.Key)) { $MethodLabels[$_.Key] } else { $_.Key }); 'Strength' = $kind; 'Users registered' = $_.Value; 'Share of enabled users' = "$([math]::Round(100 * $_.Value / [Math]::Max(1, $enabledMembers.Count)))%"; '_issues' = $(if ($kind -eq 'Weak (phone-based)') { @{ 'Strength' = 'auth-weak-method' } } else { $null }); '_tags' = $kind.ToLower().Split(' ')[0] }
})
$passkeyUsers = @($regEnabled | Where-Object { $rd = $_; @($rd.methodsRegistered | Where-Object { $_ -in $PhishingResistantMethods }).Count }).Count
$passwordlessCapable = @($regEnabled | Where-Object { $_.isPasswordlessCapable }).Count
$weakOnlyUsers = @($regEnabled | Where-Object { $rd = $_; $rd.isMfaRegistered -and -not @($rd.methodsRegistered | Where-Object { $_ -notin $WeakMethods -and $_ -notin @('email', 'securityQuestion') }).Count }).Count
$passkeyPct = if ($regEnabled.Count) { [math]::Round(100 * $passkeyUsers / $enabledMembers.Count) } else { $null }
$adminUsers = @($enabledMembers | Where-Object { $AdminIds.ContainsKey($_.id) })
$adminPhishRes = @($adminUsers | Where-Object { $rd = $RegMap[$_.id]; $rd -and @($rd.methodsRegistered | Where-Object { $_ -in $PhishingResistantMethods }).Count }).Count
$adminWithoutPhishRes = @($adminUsers | Where-Object { $rd = $RegMap[$_.id]; -not ($rd -and @($rd.methodsRegistered | Where-Object { $_ -in $PhishingResistantMethods }).Count) } | ForEach-Object { $_.displayName })
$fido2 = $authMethods | Where-Object { $_.id -eq 'Fido2' } | Select-Object -First 1
$authApp = $authMethods | Where-Object { $_.id -eq 'MicrosoftAuthenticator' } | Select-Object -First 1
$smsMethod = $authMethods | Where-Object { $_.id -eq 'Sms' } | Select-Object -First 1
$voiceMethod = $authMethods | Where-Object { $_.id -eq 'Voice' } | Select-Object -First 1
$eamConfigs = @($authMethods | Where-Object { ([string]$_.'@odata.type') -match 'externalAuthenticationMethod' })
$federatedDomains = @($org.verifiedDomains | Where-Object { $_.authenticationType -eq 'Federated' } | ForEach-Object { $_.name })
$mfaProduct = if ($eamConfigs.Count) { "Third-party MFA through Entra external authentication methods: $(Join-List ($eamConfigs | ForEach-Object { "$($_.displayName) ($($_.state))" }))" }
              elseif ($federatedDomains.Count) { "Federated identity provider for $(Join-List $federatedDomains) (MFA is enforced by that provider, for example ADFS, Okta or Duo); Microsoft Entra MFA for any managed domains" }
              else { 'Microsoft Entra multifactor authentication (Microsoft Authenticator and the methods below)' }
$topMethod = if ($methodCountRows.Count) { $methodCountRows[0]['Method'] } else { 'Not available' }
$mfaProductKv = [ordered]@{
    'Registration data source'                         = $(if ($script:RegFromMethods) { 'Per-user authentication methods (Entra ID Free; the registration report needs Entra ID P1)' } elseif (@($regDetails).Count) { 'Entra ID registration report (userRegistrationDetails)' } else { 'Not available' })
    'MFA product in use'                               = $mfaProduct
    'Most registered method'                           = $topMethod
    'Default MFA method (most common)'                 = $(if ($regEnabled.Count) { $d = @($regEnabled | ForEach-Object { [string]$_.defaultMfaMethod } | Where-Object { $_ -and $_ -ne 'none' } | Group-Object | Sort-Object Count -Descending | Select-Object -First 1); if ($d) { "$(if ($MethodLabels.Contains($d[0].Name)) { $MethodLabels[$d[0].Name] } else { $d[0].Name }) ($($d[0].Count) users)" } else { 'Not set' } } else { 'Not available' })
    'Users whose only MFA method is a phone (SMS / voice)' = $(if ($regEnabled.Count) { $weakOnlyUsers } else { 'Not available' })
    'SMS method policy / Voice method policy'          = "$(if ($smsMethod) { $smsMethod.state } else { 'n/a' }) / $(if ($voiceMethod) { $voiceMethod.state } else { 'n/a' })"
    'Microsoft Authenticator policy'                   = $(if ($authApp) { "$($authApp.state)$(if ($authApp.featureSettings) { "; number matching: $(Format-SettingValue $authApp.featureSettings.numberMatchingRequiredState.state); app context: $(Format-SettingValue $authApp.featureSettings.displayAppInformationRequiredState.state); location context: $(Format-SettingValue $authApp.featureSettings.displayLocationInformationRequiredState.state)" })" } else { 'Not available' })
    'System-preferred MFA (strongest registered method is prompted first)' = $(if ($authMethodsPolicy -and $authMethodsPolicy.systemCredentialPreferences) { [string]$authMethodsPolicy.systemCredentialPreferences.state } else { 'Not available' })
}
$mfaProductIssues = @{}
if ($regEnabled.Count -and $weakOnlyUsers -gt 0) { $mfaProductIssues['Users whose only MFA method is a phone (SMS / voice)'] = 'auth-weak-method' }
if ($smsMethod -and $smsMethod.state -eq 'enabled') { $mfaProductIssues['SMS method policy / Voice method policy'] = 'auth-weak-method' }
$passkeyKv = [ordered]@{
    'Passkey (FIDO2) method policy'                    = $(if ($fido2) { "$($fido2.state); targets: $(Resolve-AuthTargets $fido2.includeTargets); self-service registration: $(Format-SettingValue $fido2.isSelfServiceRegistrationAllowed); attestation enforced: $(Format-SettingValue $fido2.isAttestationEnforced); key restrictions: $(if ($fido2.keyRestrictions -and $fido2.keyRestrictions.isEnforced) { "$($fido2.keyRestrictions.enforcementType) $(@($fido2.keyRestrictions.aaGuids).Count) model(s)" } else { 'none (any passkey allowed)' })" } else { 'Not available' })
    'Enabled users with a phishing-resistant method (passkey, security key, Windows Hello, certificate)' = $(if ($regEnabled.Count) { "$passkeyUsers of $($enabledMembers.Count) ($passkeyPct%)" } else { 'Not available (requires the registration report)' })
    'Enabled users capable of passwordless sign-in'    = $(if ($regEnabled.Count) { "$passwordlessCapable of $($enabledMembers.Count) ($([math]::Round(100 * $passwordlessCapable / [Math]::Max(1, $enabledMembers.Count)))%)" } else { 'Not available' })
    'Administrators with a phishing-resistant method'  = $(if ($regEnabled.Count) { "$adminPhishRes of $($adminUsers.Count)$(if ($adminWithoutPhishRes.Count) { "; without: $(Join-List ($adminWithoutPhishRes | Select-Object -First 8))$(if ($adminWithoutPhishRes.Count -gt 8) { ' ...' })" })" } else { 'Not available' })
    'Rollout stage'                                    = $(if (-not $regEnabled.Count) { 'Not available' } elseif (-not $fido2 -or $fido2.state -ne 'enabled') { 'Not started: passkey method disabled' } elseif ($passkeyPct -ge 90) { 'Complete' } elseif ($passkeyUsers -gt 0) { "In progress ($passkeyPct%)" } else { 'Not started: method enabled, no registrations' })
}
$passkeyIssues = @{}
if ($regEnabled.Count -and ([string]$passkeyKv['Rollout stage']) -like 'Not started*') { $passkeyIssues['Rollout stage'] = 'passkey-not-started' }
if ($regEnabled.Count -and $adminUsers.Count -and $adminPhishRes -lt $adminUsers.Count) { $passkeyIssues['Administrators with a phishing-resistant method'] = 'passkey-admins' }
$passkeyStatus = if (-not $regEnabled.Count) { 'Not available' } elseif ($passkeyPct -ge 90) { 'Complete' } elseif ($passkeyUsers -gt 0) { "In progress ($passkeyPct% of users)" } else { 'Not started' }
$passkeyEvidence = "$($passkeyKv['Passkey (FIDO2) method policy']). Users with a phishing-resistant method: $($passkeyKv['Enabled users with a phishing-resistant method (passkey, security key, Windows Hello, certificate)']); passwordless-capable: $($passkeyKv['Enabled users capable of passwordless sign-in'])."
$adminPhishStatus = if (-not $regEnabled.Count) { 'Not available' } elseif ($adminUsers.Count -eq 0) { 'Not available' } elseif ($adminPhishRes -eq $adminUsers.Count) { "Complete ($adminPhishRes of $($adminUsers.Count) administrators)" } elseif ($adminPhishRes -gt 0) { "Partial ($adminPhishRes of $($adminUsers.Count) administrators)" } else { "None (0 of $($adminUsers.Count) administrators)" }
$adminPhishEvidence = "Administrators (any directory role) with a passkey, FIDO2 key, Windows Hello for Business or certificate registered: $adminPhishRes of $($adminUsers.Count).$(if ($adminWithoutPhishRes.Count) { " Without: $(Join-List ($adminWithoutPhishRes | Select-Object -First 10))." }) MFA product: $mfaProduct."
$MethodFilters = @(@{ Label = 'Phishing-resistant'; Tag = 'phishing-resistant' }, @{ Label = 'Standard'; Tag = 'standard' }, @{ Label = 'Weak'; Tag = 'weak' }, @{ Label = 'Recovery only'; Tag = 'recovery' })


# ----------------------------------------------------------------------------------------------
# 11. Exchange Online: external auto-forwarding and legacy protocols
# ----------------------------------------------------------------------------------------------
$exoKv = [ordered]@{}
$fwdMailboxRows = @(); $fwdRuleRows = @(); $fwdTransportRows = @(); $remoteDomainRows = @(); $casRows = @()
$dkimConfigs = @(); $spamRows = @(); $malwareRows = @(); $phishRows = @(); $safeLinksRows = @(); $safeAttachRows = @(); $presetRows = @(); $atpGlobal = $null; $connFilter = $null; $defenderCollected = $false
$exoCollected = $false
$suspiciousRuleRows = @(); $fullAccessRows = @(); $sendAsRows = @(); $recentTransportRows = @(); $mailboxAuditDisabled = $null; $exoSingleSignOn = $false
$ualCollected = $false; $ualNote = 'Unified audit log not searched (Exchange Online was not collected).'; $ualEnabled = $null; $sharingRows = @(); $anonLinkCount = 0; $extShareEventCount = 0
$downloadRows = @(); $massDownloadUrgent = 0; $massDownloadReview = 0; $downloadTotal = 0; $dlBy = @{}; $dlMedian = 0; $downloadDays = [Math]::Min($AuditDays, 7); $MassDownloadThreshold = 500
$auditRuleRows = @(); $auditRuleDanger = 0; $urlClickRows = @(); $urlClickedThrough = 0; $urlBlocked = 0; $urlClicksCollected = $false
$calendarShareRows = @(); $sharingPolicyText = ''; $sharedMailboxCount = 0
if (-not $SkipExchange) {
    if (Get-Module -ListAvailable ExchangeOnlineManagement) {
        try {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
            Import-Module ExchangeOnlineManagement
            $exoParams = @{ ShowBanner = $false }
            if ($TenantId) { $exoParams.Organization = $TenantId }
            if ($ctx.Account) { $exoParams.UserPrincipalName = $ctx.Account }   # reuse the same account signed in to Graph
            # Single sign-on: reuse the browser sign-in's refresh token for an Exchange Online token so there is no second prompt.
            $exoToken = $null
            if ($script:GraphToken -and $script:GraphToken.RefreshToken) {
                try { $exoToken = Request-ResourceToken 'https://outlook.office365.com/.default' -RequireScope 'Exchange.Manage' }
                catch {
                    $m = $_.Exception.Message
                    try { $b = $_.ErrorDetails.Message; if ($b) { $j = $b | ConvertFrom-Json; if ($j.error_description) { $m = ([string]$j.error_description -split "`r?`n")[0] } } } catch { Write-Verbose 'Token error body was not JSON.' }
                    if ($m -match 'consent|AADSTS65001|not been approved') { $script:ConsentNeeded['Exchange Online'] = 'https://outlook.office365.com/Exchange.Manage'; Write-Warning "Exchange Online: the connector's read-only Exchange permission is not approved in this tenant, so Exchange cannot reuse your sign-in. $(if ($NonInteractive) { 'Exchange is skipped in -NonInteractive mode.' } else { 'Signing in to Exchange separately (a second browser prompt).' }) The one-time approval link is printed at the end of the run." }
                    else { Write-Warning "Exchange Online single sign-on is not available ($m). $(if ($NonInteractive) { 'Exchange is skipped in -NonInteractive mode.' } else { 'A second sign-in prompt will appear.' })" }
                }
            }
            if ($exoToken) { $exoParams.AccessToken = $exoToken; $exoParams.Remove('Organization'); $exoSingleSignOn = $true }
            elseif ($NonInteractive) { throw 'Exchange Online needs its own sign-in in this tenant, and -NonInteractive does not allow a prompt.' }
            Connect-ExchangeOnline @exoParams
            $exoCollected = $true

            $remoteDomains = Get-RemoteDomain
            $remoteDomainRows = foreach ($rd in $remoteDomains) { [ordered]@{ 'Remote domain' = $rd.Name; 'Domain' = $rd.DomainName; 'Auto-forward allowed' = $rd.AutoForwardEnabled; 'Auto-reply allowed' = $rd.AutoReplyEnabled } }
            $defaultRd = $remoteDomains | Where-Object { $_.DomainName -eq '*' } | Select-Object -First 1

            $outboundSpam = Get-HostedOutboundSpamFilterPolicy
            $spamText = Join-List ($outboundSpam | ForEach-Object { "$($_.Name): $($_.AutoForwardingMode)" }) '; '

            $transportRules = Get-TransportRule
            $fwdTransportRows = foreach ($tr in ($transportRules | Where-Object { $_.RedirectMessageTo -or $_.BlindCopyTo -or $_.AddToRecipients -or $_.CopyTo })) {
                [ordered]@{ 'Rule' = $tr.Name; 'State' = $tr.State; 'Redirect to' = (Join-List $tr.RedirectMessageTo); 'Bcc to' = (Join-List $tr.BlindCopyTo); 'Copy / add to' = (Join-List (@($tr.CopyTo) + @($tr.AddToRecipients))); 'Mode' = $tr.Mode }
            }

            $mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails UserMailbox, SharedMailbox
            $acceptedDomains = @((Get-AcceptedDomain).DomainName | ForEach-Object { [string]$_ })
            function Test-ExternalAddress { param([string]$addr) if (-not $addr) { return $false } ; $d = ($addr -replace '^smtp:', '' -split '@')[-1] ; return -not ($acceptedDomains -contains $d) }
            $fwdMailboxRows = foreach ($mb in ($mailboxes | Where-Object { $_.ForwardingSmtpAddress -or $_.ForwardingAddress })) {
                [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Forwarding SMTP address' = $mb.ForwardingSmtpAddress; 'Forwarding address (internal)' = $mb.ForwardingAddress; 'External' = $(if (Test-ExternalAddress ([string]$mb.ForwardingSmtpAddress)) { 'Yes' } else { 'No' }); 'Keep copy in mailbox' = $mb.DeliverToMailboxAndForward }
            }

            if (-not $SkipInboxRules) {
                $i = 0
                foreach ($mb in $mailboxes) {
                    $i++
                    Write-Progress -Activity 'Scanning inbox rules for forwarding' -Status $mb.PrimarySmtpAddress -PercentComplete ([int](100 * $i / [Math]::Max(1, @($mailboxes).Count)))
                    try {
                        $allRules = @(Get-InboxRule -Mailbox $mb.Identity -ErrorAction Stop)
                        foreach ($r in ($allRules | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo })) {
                            $targets = @(@($r.ForwardTo) + @($r.ForwardAsAttachmentTo) + @($r.RedirectTo) | Where-Object { $_ })
                            $ext = @($targets | Where-Object { Test-ExternalAddress ([string]$_) }).Count -gt 0
                            $fwdRuleRows += [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Rule' = $r.Name; 'Enabled' = $r.Enabled; 'Forwards to' = (Join-List ($targets | ForEach-Object { [string]$_ }) '; '); 'External' = $(if ($ext) { 'Yes' } else { 'No' }) }
                        }
                        # Indications of compromise: rules that hide, delete or exfiltrate mail (classic business email compromise tradecraft)
                        foreach ($r in $allRules) {
                            $why = @()
                            if ($r.DeleteMessage) { $why += 'deletes mail' }
                            if ($r.SoftDeleteMessage) { $why += 'soft-deletes mail' }
                            if ($r.MoveToFolder -and ([string]$r.MoveToFolder) -match '(?i)RSS|Archive|Deleted Items|Junk|Conversation History|Notes') { $why += "moves mail to '$(([string]$r.MoveToFolder) -replace '^.*\\', '')'" }
                            if ($r.MarkAsRead -and ($r.ForwardTo -or $r.RedirectTo -or $r.ForwardAsAttachmentTo -or $r.MoveToFolder -or $r.DeleteMessage)) { $why += 'marks as read' }
                            if (([string]$r.Name).Trim() -match '^[\W_]{0,3}$') { $why += 'blank or punctuation-only name' }
                            if (@(@($r.ForwardTo) + @($r.ForwardAsAttachmentTo) + @($r.RedirectTo) | Where-Object { $_ -and (Test-ExternalAddress ([string]$_)) }).Count) { $why += 'forwards outside the organization' }
                            if ($why.Count) {
                                $conds = @()
                                if ($r.From) { $conds += "from: $(Join-List ($r.From | ForEach-Object { [string]$_ }))" }
                                if ($r.SubjectContainsWords) { $conds += "subject contains: $(Join-List $r.SubjectContainsWords)" }
                                if ($r.SubjectOrBodyContainsWords) { $conds += "subject or body contains: $(Join-List $r.SubjectOrBodyContainsWords)" }
                                if ($r.BodyContainsWords) { $conds += "body contains: $(Join-List $r.BodyContainsWords)" }
                                $suspiciousRuleRows += [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Rule' = $r.Name; 'Enabled' = $r.Enabled; 'Why it is suspicious' = (Join-List $why '; '); 'Conditions' = $(if ($conds.Count) { Join-List $conds '; ' } else { 'All messages' }); '_issues' = @{ 'Why it is suspicious' = 'ioc-inbox-rule-hide' }; '_tags' = $(if ($r.Enabled) { 'enabled' } else { 'disabled' }) }
                            }
                        }
                        # Sharing and access: calendar folder open to everyone in the organization (Default) or to anonymous users
                        try {
                            foreach ($cp in @(Get-MailboxFolderPermission -Identity "$($mb.PrimarySmtpAddress):\Calendar" -ErrorAction Stop | Where-Object { ([string]$_.User) -in @('Default', 'Anonymous') -and ([string]$_.AccessRights) -notmatch '^(None|AvailabilityOnly)$' })) {
                                $calendarShareRows += [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Level' = [string]$cp.User; 'Access rights' = (Join-List (@($cp.AccessRights) | ForEach-Object { [string]$_ })) }
                            }
                        } catch { Write-Verbose "Calendar permissions not readable for $($mb.PrimarySmtpAddress): $($_.Exception.Message)" }
                        # Indications of compromise: Full Access delegates that are not the owner or the system
                        foreach ($perm in @(Get-MailboxPermission -Identity $mb.Identity -ErrorAction Stop | Where-Object { -not $_.IsInherited -and $_.Deny -ne $true -and ([string]$_.User) -notmatch '^(NT AUTHORITY|S-1-5)' -and ((@($_.AccessRights) | ForEach-Object { [string]$_ }) -join ',') -match 'FullAccess' })) {
                            $fullAccessRows += [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Mailbox type' = [string]$mb.RecipientTypeDetails; 'Delegate (Full Access)' = [string]$perm.User; 'Rights' = (Join-List (@($perm.AccessRights) | ForEach-Object { [string]$_ })); '_issues' = @{ 'Delegate (Full Access)' = 'ioc-delegate' }; '_tags' = $(if ([string]$mb.RecipientTypeDetails -eq 'SharedMailbox') { 'shared' } else { 'user' }) }
                        }
                    } catch { Write-Warning "Inbox rules or permissions not readable for $($mb.PrimarySmtpAddress): $($_.Exception.Message)" }
                }
                Write-Progress -Activity 'Scanning inbox rules for forwarding' -Completed
            }

            $transportConfig = Get-TransportConfig
            $orgConfig = Get-OrganizationConfig
            $casMailboxes = Get-CASMailbox -ResultSize Unlimited
            $casRows = foreach ($c in ($casMailboxes | Where-Object { $_.PopEnabled -or $_.ImapEnabled -or ($_.SmtpClientAuthenticationDisabled -eq $false) } | Sort-Object PrimarySmtpAddress)) {
                [ordered]@{ 'Mailbox' = $c.PrimarySmtpAddress; 'POP enabled' = $c.PopEnabled; 'IMAP enabled' = $c.ImapEnabled; 'SMTP AUTH' = $(if ($c.SmtpClientAuthenticationDisabled -eq $false) { 'Enabled (mailbox override)' } elseif ($c.SmtpClientAuthenticationDisabled -eq $true) { 'Disabled' } else { 'Follows tenant default' }); 'ActiveSync' = $c.ActiveSyncEnabled; 'MAPI' = $c.MAPIEnabled }
            }
            $authPolicies = Get-AuthenticationPolicy
            $defaultAuthPolicy = $orgConfig.DefaultAuthenticationPolicy

            # ---- Email protection: Exchange Online Protection and Defender for Office 365 policies (each cmdlet tolerated separately)
            Write-Host "Collecting Exchange Online Protection and Defender for Office 365 policies..." -ForegroundColor Cyan
            function Invoke-ExoSafe { param([scriptblock]$Script, [string]$What) try { return @(& $Script) } catch { Write-Verbose "Could not read $What : $($_.Exception.Message)"; return @() } }
            $dkimConfigs = Invoke-ExoSafe { Get-DkimSigningConfig } 'DKIM signing configuration'
            $spamPolicies = Invoke-ExoSafe { Get-HostedContentFilterPolicy } 'anti-spam policies'
            $spamRows = foreach ($p in $spamPolicies) {
                $is = @{}
                if ($p.HighConfidenceSpamAction -ne 'Quarantine') { $is['High-confidence spam'] = 'spam-action-weak' }
                if ($p.PhishSpamAction -ne 'Quarantine') { $is['Phishing'] = 'spam-action-weak' }
                if ($p.HighConfidencePhishAction -ne 'Quarantine') { $is['High-confidence phishing'] = 'spam-action-weak' }
                if ([int]$p.BulkThreshold -gt 6) { $is['Bulk threshold'] = 'bulk-threshold' }
                if ($p.SpamZapEnabled -eq $false -or $p.PhishZapEnabled -eq $false) { $is['ZAP (spam / phish)'] = 'zap-off' }
                [ordered]@{ 'Policy' = $p.Name; 'Default' = $(if ($p.IsDefault) { 'Yes' } else { '' }); 'Spam' = $p.SpamAction; 'High-confidence spam' = $p.HighConfidenceSpamAction; 'Phishing' = $p.PhishSpamAction; 'High-confidence phishing' = $p.HighConfidencePhishAction; 'Bulk threshold' = $p.BulkThreshold; 'Bulk action' = $p.BulkSpamAction; 'ZAP (spam / phish)' = "$($p.SpamZapEnabled) / $($p.PhishZapEnabled)"; 'Quarantine retention (days)' = $p.QuarantineRetentionPeriod; 'End-user notifications' = $p.EnableEndUserSpamNotifications; '_issues' = $(if ($is.Count) { $is } else { $null }) }
            }
            $malwarePolicies = Invoke-ExoSafe { Get-MalwareFilterPolicy } 'anti-malware policies'
            $malwareRows = foreach ($p in $malwarePolicies) {
                $is = @{}; if ($p.EnableFileFilter -ne $true) { $is['Common attachments filter'] = 'malware-filefilter-off' }
                [ordered]@{ 'Policy' = $p.Name; 'Default' = $(if ($p.IsDefault) { 'Yes' } else { '' }); 'Common attachments filter' = $p.EnableFileFilter; 'Blocked file types' = @($p.FileTypes).Count; 'File filter action' = $p.FileTypeAction; 'ZAP for malware' = $p.ZapEnabled; 'Notify admins (internal / external senders)' = "$($p.EnableInternalSenderAdminNotifications) / $($p.EnableExternalSenderAdminNotifications)"; '_issues' = $(if ($is.Count) { $is } else { $null }) }
            }
            $phishPolicies = Invoke-ExoSafe { Get-AntiPhishPolicy } 'anti-phishing policies'
            $phishRows = foreach ($p in $phishPolicies) {
                $is = @{}
                if ($script:HasDefenderP1 -and $p.Enabled -and -not ($p.EnableTargetedUserProtection -and $p.EnableTargetedDomainsProtection -and $p.EnableMailboxIntelligenceProtection)) { $is['Impersonation protection (users / domains / mailbox intelligence)'] = 'antiphish-impersonation-off' }
                if ($p.Enabled -and [int]$p.PhishThresholdLevel -lt 2) { $is['Phishing threshold'] = 'phish-threshold-low' }
                [ordered]@{ 'Policy' = $p.Name; 'Default' = $(if ($p.IsDefault) { 'Yes' } else { '' }); 'Enabled' = $p.Enabled; 'Spoof intelligence' = $p.EnableSpoofIntelligence; 'Phishing threshold' = $p.PhishThresholdLevel; 'Impersonation protection (users / domains / mailbox intelligence)' = "$($p.EnableTargetedUserProtection) / $($p.EnableTargetedDomainsProtection) / $($p.EnableMailboxIntelligenceProtection)"; 'Protected users' = @($p.TargetedUsersToProtect).Count; 'Impersonation actions (user / domain / mailbox intelligence)' = "$($p.TargetedUserProtectionAction) / $($p.TargetedDomainProtectionAction) / $($p.MailboxIntelligenceProtectionAction)"; 'Safety tips (first contact / similar user / similar domain)' = "$($p.EnableFirstContactSafetyTips) / $($p.EnableSimilarUsersSafetyTips) / $($p.EnableSimilarDomainsSafetyTips)"; 'Unauthenticated sender indicator' = $p.EnableUnauthenticatedSender; '_issues' = $(if ($is.Count) { $is } else { $null }) }
            }
            foreach ($r in @($phishRows)) { if ($r) { foreach ($k in @($r.Keys)) { if ($k -ne '_issues' -and ([string]$r[$k]) -match '^[\s/]*$') { $r[$k] = $(if ($script:HasDefenderP1) { '' } else { 'n/a (Defender for Office 365 only)' }) } } } }
            $safeLinksPolicies = Invoke-ExoSafe { Get-SafeLinksPolicy } 'Safe Links policies'
            $safeLinksRows = foreach ($p in $safeLinksPolicies) { [ordered]@{ 'Policy' = $p.Name; 'Email' = $p.EnableSafeLinksForEmail; 'Teams' = $p.EnableSafeLinksForTeams; 'Office apps' = $p.EnableSafeLinksForOffice; 'Scan URLs in real time' = $p.ScanUrls; 'Wait for scan before delivery' = $p.DeliverMessageAfterScan; 'Track clicks' = $p.TrackClicks; 'Users may click through' = $p.AllowClickThrough; 'Internal mail' = $p.EnableForInternalSenders } }
            $safeAttachPolicies = Invoke-ExoSafe { Get-SafeAttachmentPolicy } 'Safe Attachments policies'
            $safeAttachRows = foreach ($p in $safeAttachPolicies) { [ordered]@{ 'Policy' = $p.Name; 'Enabled' = $p.Enable; 'Action' = $p.Action; 'Redirect blocked mail' = $p.Redirect; 'Redirect to' = $p.RedirectAddress; 'If scan errors or times out' = $p.ActionOnError } }
            $atpGlobal = (Invoke-ExoSafe { Get-AtpPolicyForO365 } 'Defender for Office 365 global settings') | Select-Object -First 1
            $connFilter = (Invoke-ExoSafe { Get-HostedConnectionFilterPolicy -Identity Default } 'connection filter policy') | Select-Object -First 1
            $presetRows = @()
            foreach ($rule in (Invoke-ExoSafe { Get-EOPProtectionPolicyRule } 'EOP preset policy rules')) { $presetRows += [ordered]@{ 'Preset' = $rule.Name; 'Scope' = 'Exchange Online Protection (anti-spam, anti-malware, anti-phishing)'; 'State' = $rule.State; 'Priority' = $rule.Priority; 'Applies to' = (Join-List (@($rule.SentTo) + @($rule.SentToMemberOf) + @($rule.RecipientDomainIs) | Where-Object { $_ }) '; ') } }
            foreach ($rule in (Invoke-ExoSafe { Get-ATPProtectionPolicyRule } 'Defender preset policy rules')) { $presetRows += [ordered]@{ 'Preset' = $rule.Name; 'Scope' = 'Defender for Office 365 (Safe Links, Safe Attachments)'; 'State' = $rule.State; 'Priority' = $rule.Priority; 'Applies to' = (Join-List (@($rule.SentTo) + @($rule.SentToMemberOf) + @($rule.RecipientDomainIs) | Where-Object { $_ }) '; ') } }
            $defenderCollected = $true

            # ---- Indications of compromise from Exchange: Send As grants, recently changed transport rules, mailbox auditing
            $sendAsRows = @(Invoke-ExoSafe { Get-RecipientPermission -ResultSize Unlimited } 'Send As permissions' | Where-Object { $_ -and ([string]$_.Trustee) -notmatch '^(NT AUTHORITY|S-1-5)' -and ((@($_.AccessRights) | ForEach-Object { [string]$_ }) -join ',') -match 'SendAs' -and ([string]$_.AccessControlType) -ne 'Deny' } | ForEach-Object { [ordered]@{ 'Mailbox' = [string]$_.Identity; 'Trustee (can send as)' = [string]$_.Trustee; '_issues' = @{ 'Trustee (can send as)' = 'ioc-sendas' } } })
            $recentTransportRows = @($transportRules | Where-Object { $_.WhenChanged -and ([datetime]$_.WhenChanged) -ge (Get-Date).AddDays(-$AuditDays) } | ForEach-Object { [ordered]@{ 'Rule' = $_.Name; 'State' = [string]$_.State; 'Changed' = (Format-DateText $_.WhenChanged); 'Created' = (Format-DateText $_.WhenCreated); 'Actions' = (Join-List (@($_.Actions) | ForEach-Object { [string]$_ }) '; '); '_issues' = @{ 'Changed' = 'ioc-transport-change' } } })
            $mailboxAuditDisabled = [bool]$orgConfig.AuditDisabled
            $sharedMailboxCount = @($mailboxes | Where-Object { [string]$_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
            $sharingPolicies = Invoke-ExoSafe { Get-SharingPolicy } 'sharing policies'
            $sharingPolicyText = Join-List ($sharingPolicies | ForEach-Object { "$($_.Name) ($(if ($_.Enabled) { 'enabled' } else { 'disabled' })): $(Join-List (@($_.Domains) | ForEach-Object { [string]$_ }) ', ')" }) ' | '

            # ---- Unified audit log (Microsoft Purview Audit): sharing, mass downloads, mailbox rule and forwarding changes
            $ualConfig = (Invoke-ExoSafe { Get-AdminAuditLogConfig } 'audit log configuration') | Select-Object -First 1
            $ualEnabled = if ($ualConfig) { [bool]$ualConfig.UnifiedAuditLogIngestionEnabled } else { $null }
            if ($SkipAuditLog) { $ualNote = 'Audit log search skipped (-SkipAuditLog).' }
            elseif ($ualEnabled -eq $false) { $ualNote = 'The unified audit log is turned off for this tenant (Get-AdminAuditLogConfig UnifiedAuditLogIngestionEnabled = False), so sharing, download and mailbox-rule activity cannot be reviewed.' }
            else {
                $ualNote = ''
                Write-Host "Searching the unified audit log (sharing, downloads, mailbox rules; last $AuditDays days)..." -ForegroundColor Cyan
                function Search-AuditLogSafe {
                    # Pages through Search-UnifiedAuditLog with a session; returns parsed AuditData objects (read-only).
                    param([string[]]$Operations, [int]$Days, [int]$MaxRecords, [string]$What)
                    $out = New-Object System.Collections.Generic.List[object]
                    $sid = [guid]::NewGuid().ToString()
                    try {
                        do {
                            $page = @(Search-UnifiedAuditLog -StartDate (Get-Date).AddDays(-$Days) -EndDate (Get-Date).AddMinutes(5) -Operations $Operations -SessionId $sid -SessionCommand ReturnLargeSet -ResultSize 5000 -ErrorAction Stop)
                            foreach ($rec in $page) { try { $d = $rec.AuditData | ConvertFrom-Json; if ($d) { $out.Add($d) } } catch { Write-Verbose 'Audit record was not JSON.' } }
                        } while ($page.Count -eq 5000 -and $out.Count -lt $MaxRecords)
                    } catch { $script:GraphFailures["audit log: $What"] = $_.Exception.Message; Write-Warning "Audit log search ($What) failed: $($_.Exception.Message)" }
                    return $out.ToArray()
                }
                $ualCollected = $true
                $sharingEvents = Search-AuditLogSafe -Operations @('AnonymousLinkCreated', 'SharingInvitationCreated', 'AddedToSecureLink', 'SecureLinkCreated', 'SharingSet', 'CompanyLinkCreated') -Days $AuditDays -MaxRecords 20000 -What 'sharing'
                $sharingRows = @($sharingEvents | Where-Object { $_.Operation -in @('AnonymousLinkCreated', 'SharingInvitationCreated', 'AddedToSecureLink', 'SecureLinkCreated') -or ([string]$_.TargetUserOrGroupType) -eq 'Guest' } | Sort-Object CreationTime -Descending | Select-Object -First 300 | ForEach-Object {
                    $anon = $_.Operation -eq 'AnonymousLinkCreated'
                    $isExternal = $anon -or ([string]$_.TargetUserOrGroupType) -eq 'Guest' -or (([string]$_.TargetUserOrGroupName) -and (Test-ExternalAddress ([string]$_.TargetUserOrGroupName)))
                    [ordered]@{ 'Date (local)' = (Format-DateText $_.CreationTime); 'User' = $_.UserId; 'Action' = $(switch ($_.Operation) { 'AnonymousLinkCreated' { 'Created an Anyone link (no sign-in needed)' } 'SharingInvitationCreated' { 'Invited an outside person' } 'AddedToSecureLink' { 'Added a person to a sharing link' } 'SecureLinkCreated' { 'Created a specific-people link' } default { $_ } }); 'Shared with' = $(if ($anon) { 'Anyone with the link' } else { [string]$_.TargetUserOrGroupName }); 'Item' = $(if ($_.SourceFileName) { $_.SourceFileName } else { ([string]$_.ObjectId -split '/')[-1] }); 'Site' = $_.SiteUrl; 'IP address' = $_.ClientIP; '_issues' = $(if ($anon) { @{ 'Action' = 'ioc-anon-link' } } elseif ($isExternal) { @{ 'Shared with' = 'ioc-ext-share' } } else { $null }); '_tags' = $(if ($anon) { 'anonymous' } elseif ($isExternal) { 'external' } else { 'internal' }) }
                })
                $anonLinkCount = @($sharingRows | Where-Object { $_['_tags'] -eq 'anonymous' }).Count
                $extShareEventCount = @($sharingRows | Where-Object { $_['_tags'] -eq 'external' }).Count

                $downloadDays = [Math]::Min($AuditDays, 7)
                $downloadEvents = Search-AuditLogSafe -Operations @('FileDownloaded', 'FileSyncDownloadedFull') -Days $downloadDays -MaxRecords 50000 -What 'downloads'
                $dlBy = @{}
                foreach ($ev in $downloadEvents) {
                    $u = [string]$ev.UserId; if (-not $u -or $u -match '^(app@sharepoint|SHAREPOINT\\system)') { continue }
                    if (-not $dlBy.ContainsKey($u)) { $dlBy[$u] = @{ Count = 0; Sync = 0; Ips = @{}; Sites = @{}; Last = $null; Agents = @{} } }
                    $dlBy[$u].Count++; if ($ev.Operation -eq 'FileSyncDownloadedFull') { $dlBy[$u].Sync++ }
                    $dlBy[$u].Ips[[string]$ev.ClientIP] = $true; $dlBy[$u].Sites[[string]$ev.SiteUrl] = $true; $dlBy[$u].Agents[[string]$ev.UserAgent] = $true
                    if (-not $dlBy[$u].Last -or ([datetime]$ev.CreationTime) -gt $dlBy[$u].Last) { $dlBy[$u].Last = [datetime]$ev.CreationTime }
                }
                $dlCounts = @($dlBy.Values | ForEach-Object { $_.Count } | Sort-Object)
                $dlMedian = if ($dlCounts.Count) { $dlCounts[[int][math]::Floor($dlCounts.Count / 2)] } else { 0 }
                $MassDownloadThreshold = 500
                $downloadRows = @($dlBy.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | Select-Object -First 25 | ForEach-Object {
                    $urgent = $_.Value.Count -ge $MassDownloadThreshold
                    $review = -not $urgent -and $dlMedian -gt 0 -and $_.Value.Count -ge [Math]::Max(100, 5 * $dlMedian)
                    [ordered]@{ 'User' = $_.Key; "Files downloaded ($downloadDays days)" = $_.Value.Count; 'Of which full sync downloads' = $_.Value.Sync; 'Distinct IP addresses' = $_.Value.Ips.Count; 'Sites' = $_.Value.Sites.Count; 'Clients' = (Join-List ($_.Value.Agents.Keys | Where-Object { $_ } | ForEach-Object { $(if ($_ -match 'OneDriveSync|Microsoft SkyDriveSync') { 'OneDrive sync' } elseif ($_ -match 'rclone') { 'rclone' } elseif ($_ -match 'Mozilla|Chrome|Edg|Safari') { 'Browser' } else { $_.Substring(0, [Math]::Min(30, $_.Length)) }) } | Select-Object -Unique -First 3)); 'Last' = (Format-DateText $_.Value.Last); '_issues' = $(if ($urgent -or $review) { @{ "Files downloaded ($downloadDays days)" = 'ioc-mass-download' } } else { $null }); '_tags' = $(if ($urgent) { 'urgent' } elseif ($review) { 'review' } else { 'normal' }) }
                })
                $massDownloadUrgent = @($downloadRows | Where-Object { $_['_tags'] -eq 'urgent' }).Count
                $massDownloadReview = @($downloadRows | Where-Object { $_['_tags'] -eq 'review' }).Count
                $downloadTotal = $downloadEvents.Count

                $ruleEvents = Search-AuditLogSafe -Operations @('New-InboxRule', 'Set-InboxRule', 'UpdateInboxRules', 'Set-Mailbox', 'Add-MailboxPermission', 'Add-RecipientPermission', 'New-TransportRule', 'Set-TransportRule', 'Set-MailboxJunkEmailConfiguration', 'Remove-MailboxAuditBypassAssociation', 'Set-MailboxAuditBypassAssociation') -Days $AuditDays -MaxRecords 20000 -What 'mailbox rules and permissions'
                $auditRuleRows = @($ruleEvents | ForEach-Object {
                    $parms = @(); foreach ($pp in @($_.Parameters)) { if ($pp.Name -and $pp.Name -notin @('Identity', 'Mailbox', 'Force', 'AlwaysDeleteOutlookRulesBlob')) { $parms += "$($pp.Name)=$([string]$pp.Value)" } }
                    $ptext = Join-List $parms '; '
                    if ($_.Operation -eq 'Set-Mailbox' -and $ptext -notmatch 'Forwarding|DeliverToMailboxAndForward|AuditEnabled|GrantSendOnBehalfTo') { return }
                    $danger = $ptext -match '(?i)ForwardTo|ForwardAsAttachmentTo|RedirectTo|ForwardingSmtpAddress|ForwardingAddress|DeleteMessage|MoveToFolder|SoftDeleteMessage|MarkAsRead|FullAccess|SendAs|RedirectMessageTo|BlindCopyTo|AuditEnabled=False|TrustedSendersAndDomains|BypassSpam'
                    [ordered]@{ 'Date (local)' = (Format-DateText $_.CreationTime); 'Performed by' = $_.UserId; 'Operation' = $_.Operation; 'Mailbox / object' = $(if ($_.ObjectId) { ([string]$_.ObjectId -split '\\')[-1] } else { '' }); 'Parameters' = $(if ($ptext.Length -gt 220) { $ptext.Substring(0, 220) + '...' } else { $ptext }); 'IP address' = $_.ClientIP; 'Client' = $(if ($_.ClientInfoString) { ([string]$_.ClientInfoString).Substring(0, [Math]::Min(40, ([string]$_.ClientInfoString).Length)) } else { '' }); '_issues' = $(if ($danger) { @{ 'Parameters' = 'ioc-audit-rule-change' } } else { $null }); '_tags' = $(if ($danger) { 'danger' } else { 'other' }) }
                } | Where-Object { $_ } | Sort-Object { $_['Date (local)'] } -Descending | Select-Object -First 300)
                $auditRuleDanger = @($auditRuleRows | Where-Object { $_['_tags'] -eq 'danger' }).Count
            }

            # ---- Safe Links: users who clicked malicious URLs (Defender for Office 365 Plan 1 or 2)
            if ($script:HasDefenderP1) {
                $clickEvents = Invoke-ExoSafe { Get-SafeLinksDetailReport -StartDate (Get-Date).AddDays(-7) -EndDate (Get-Date) -Action Blocked, ClickedEvenBlocked, ClickedDuringScan -PageSize 5000 } 'Safe Links detail report'
                $urlClickRows = @($clickEvents | Sort-Object ClickTime -Descending | Select-Object -First 200 | ForEach-Object {
                    $clicked = ([string]$_.Action) -ne 'Blocked'
                    [ordered]@{ 'Date (local)' = (Format-DateText $_.ClickTime); 'User' = $_.RecipientAddress; 'Result' = $(switch ([string]$_.Action) { 'Blocked' { 'Blocked (user stopped)' } 'ClickedEvenBlocked' { 'Clicked through the warning' } 'ClickedDuringScan' { 'Clicked while the URL was still being scanned' } default { $_ } }); 'URL' = $(if (([string]$_.Url).Length -gt 90) { ([string]$_.Url).Substring(0, 90) + '...' } else { [string]$_.Url }); 'Where' = $_.AppName; 'Source' = $_.SourceId; '_issues' = @{ 'Result' = 'ioc-url-click' }; '_tags' = $(if ($clicked) { 'clicked' } else { 'blocked' }) }
                })
                $urlClickedThrough = @($urlClickRows | Where-Object { $_['_tags'] -eq 'clicked' }).Count
                $urlBlocked = @($urlClickRows | Where-Object { $_['_tags'] -eq 'blocked' }).Count
                $urlClicksCollected = -not (Get-GraphFailure 'Safe Links detail report')
            }

            $exoKv = [ordered]@{
                'Default remote domain (*) allows auto-forward'   = $(if ($defaultRd) { $defaultRd.AutoForwardEnabled } else { 'Not found' })
                'Outbound spam policy automatic forwarding'       = $spamText
                'Mailboxes forwarding to an external address'     = @($fwdMailboxRows | Where-Object { $_.External -eq 'Yes' }).Count
                'Inbox rules forwarding to an external address'   = $(if ($SkipInboxRules) { 'Not scanned' } else { @($fwdRuleRows | Where-Object { $_.External -eq 'Yes' }).Count })
                'Transport rules that redirect, copy, or bcc mail' = @($fwdTransportRows).Count
                'Connected to Exchange Online with'               = $(if ($exoSingleSignOn) { 'Single sign-on token from the browser sign-in (no second prompt)' } else { 'A separate Exchange Online sign-in prompt' })
                'Modern authentication (OAuth2) enabled'          = $orgConfig.OAuth2ClientProfileEnabled
                'SMTP AUTH disabled tenant-wide'                  = $transportConfig.SmtpClientAuthenticationDisabled
                'Mailboxes with POP enabled'                      = @($casMailboxes | Where-Object { $_.PopEnabled }).Count
                'Mailboxes with IMAP enabled'                     = @($casMailboxes | Where-Object { $_.ImapEnabled }).Count
                'Default authentication policy'                   = $(if ($defaultAuthPolicy) { [string]$defaultAuthPolicy } else { 'None set' })
                'Authentication policies defined'                 = $(if ($authPolicies) { Join-List ($authPolicies | ForEach-Object { $ap = $_; $basic = @($ap.PSObject.Properties | Where-Object { $_.Name -like 'AllowBasicAuth*' -and $_.Value -eq $true } | ForEach-Object { $_.Name -replace 'AllowBasicAuth', '' }); "$($ap.Name) (basic auth allowed: $(if ($basic.Count) { $basic -join ',' } else { 'none' }))" }) '; ' } else { 'None' })
            }
        } catch {
            $exoErr = [string]$_.Exception.Message
            if ($exoErr -match '(?i)unauthori[sz]ed|\b401\b|access ?denied') { $exoErr = "Exchange Online refused $($ctx.Account) ($exoErr). The account needs an Exchange role: Global Reader, View-Only Organization Management or Exchange Administrator." }
            Write-Warning "Exchange Online collection failed: $exoErr"
            $exoKv = [ordered]@{ 'Exchange Online' = "Collection failed: $($_.Exception.Message)" }
        } finally { try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'Exchange Online session was already closed.' } }
    } else {
        Write-Warning "ExchangeOnlineManagement module not installed; skipping forwarding and POP/IMAP/SMTP AUTH checks. Install-Module ExchangeOnlineManagement -Scope CurrentUser"
        $exoKv = [ordered]@{ 'Exchange Online' = 'Not collected: ExchangeOnlineManagement module not installed on the machine that ran this report' }
    }
} else { $exoKv = [ordered]@{ 'Exchange Online' = 'Not collected (-SkipExchange)' } }

# ----------------------------------------------------------------------------------------------
# 12. Security Controls Summary (computed)
# ----------------------------------------------------------------------------------------------
# ----------------------------------------------------------------------------------------------
# Email authentication: SPF, DKIM, DMARC per email-enabled verified domain (public DNS, plus Exchange DKIM signing state)
# ----------------------------------------------------------------------------------------------
Write-Host "Checking SPF, DKIM and DMARC records..." -ForegroundColor Cyan
$dnsAvailable = [bool](Get-Command Resolve-DnsName -ErrorAction SilentlyContinue)
function Get-DnsTxt { param([string]$Name) if (-not $dnsAvailable) { return @() } ; try { return @(Resolve-DnsName -Name $Name -Type TXT -ErrorAction Stop | Where-Object { $_.Strings } | ForEach-Object { ($_.Strings -join '') }) } catch { return @() } }
function Get-DnsCname { param([string]$Name) if (-not $dnsAvailable) { return $null } ; try { return (Resolve-DnsName -Name $Name -Type CNAME -ErrorAction Stop | Where-Object { $_.NameHost } | Select-Object -First 1).NameHost } catch { return $null } }
function Get-DnsMx { param([string]$Name) if (-not $dnsAvailable) { return @() } ; try { return @(Resolve-DnsName -Name $Name -Type MX -ErrorAction Stop | Where-Object { $_.NameExchange } | Sort-Object Preference | ForEach-Object { $_.NameExchange }) } catch { return @() } }
$DkimByDomain = @{}; foreach ($d in $dkimConfigs) { if ($d.Domain) { $DkimByDomain[$d.Domain.ToLower()] = $d } }
$emailDomains = @($org.verifiedDomains | Where-Object { $_.capabilities -match 'Email' -and $_.name -notmatch '\.onmicrosoft\.com$' } | ForEach-Object { $_.name })
$emailAuthRows = foreach ($dom in ($emailDomains | Sort-Object)) {
    $mx    = Get-DnsMx $dom
    $spf   = (Get-DnsTxt $dom) | Where-Object { $_ -match '^v=spf1' } | Select-Object -First 1
    $dmarc = (Get-DnsTxt "_dmarc.$dom") | Where-Object { $_ -match '^v=DMARC1' } | Select-Object -First 1
    $sel1  = Get-DnsCname "selector1._domainkey.$dom"; $sel2 = Get-DnsCname "selector2._domainkey.$dom"
    $dkimCfg = $DkimByDomain[$dom.ToLower()]
    $spfMode = if (-not $spf) { 'Missing' } elseif ($spf -match '-all') { 'Hard fail (-all)' } elseif ($spf -match '~all') { 'Soft fail (~all)' } elseif ($spf -match '\?all') { 'Neutral (?all)' } elseif ($spf -match '\+all') { 'Pass all (+all)' } else { 'No all mechanism' }
    $dmarcPolicy = if (-not $dmarc) { 'Missing' } elseif ($dmarc -match 'p=reject') { 'reject' } elseif ($dmarc -match 'p=quarantine') { 'quarantine' } elseif ($dmarc -match 'p=none') { 'none (monitor only)' } else { 'unknown' }
    $dmarcPct = if ($dmarc -match 'pct=(\d+)') { "$($Matches[1])%" } else { $(if ($dmarc) { '100%' } else { '' }) }
    $dmarcRua = if ($dmarc -match 'rua=') { 'Yes' } elseif ($dmarc) { 'No' } else { '' }
    $dkimSigning = if ($dkimCfg) { $(if ($dkimCfg.Enabled) { 'Enabled' } else { 'Disabled' }) } elseif ($exoCollected) { 'Not configured' } else { 'Not collected' }
    $dkimDns = if ($sel1 -and $sel2) { 'Both selectors published' } elseif ($sel1 -or $sel2) { 'One selector published' } else { 'Not published' }
    $is = @{}
    if (-not $spf) { $is['SPF'] = 'spf-missing' } elseif ($spfMode -notmatch 'Hard fail') { $is['SPF'] = 'spf-weak' }
    if ($dkimSigning -ne 'Enabled' -or $dkimDns -ne 'Both selectors published') { $is['DKIM signing (Exchange)'] = 'dkim-missing' }
    if (-not $dmarc) { $is['DMARC policy'] = 'dmarc-missing' } elseif ($dmarcPolicy -like 'none*') { $is['DMARC policy'] = 'dmarc-none' }
    [ordered]@{ 'Domain' = $dom; 'MX' = (Join-List $mx '; '); 'SPF' = $spfMode; 'SPF record' = $spf; 'DKIM signing (Exchange)' = $dkimSigning; 'DKIM DNS (selector1 / selector2)' = $dkimDns; 'DMARC policy' = $dmarcPolicy; 'DMARC coverage' = $dmarcPct; 'DMARC reporting (rua)' = $dmarcRua; 'DMARC record' = $dmarc; '_issues' = $(if ($is.Count) { $is } else { $null }) }
}
$emailAuthRows = @($emailAuthRows)
$spfBad   = @($emailAuthRows | Where-Object { $_['SPF'] -ne 'Hard fail (-all)' })
$dkimBad  = @($emailAuthRows | Where-Object { $_['DKIM signing (Exchange)'] -ne 'Enabled' -or $_['DKIM DNS (selector1 / selector2)'] -ne 'Both selectors published' })
$dmarcBad = @($emailAuthRows | Where-Object { $_['DMARC policy'] -notin @('reject', 'quarantine') })
function Get-EmailAuthStatus { param($Bad, [scriptblock]$HasRecord)
    # Complete: every domain passes. Partial: some pass, or a record exists but is weak. Missing: no domain has the record at all.
    if (-not $dnsAvailable) { return 'Not available' }
    if ($emailAuthRows.Count -eq 0) { return 'No email domains' }
    if ($Bad.Count -eq 0) { return 'Complete' }
    if ($Bad.Count -lt $emailAuthRows.Count) { return 'Partial' }
    if ($HasRecord -and @($Bad | Where-Object $HasRecord).Count) { return 'Partial' }
    return 'Missing'
}
$spfStatus   = Get-EmailAuthStatus $spfBad   { $_['SPF'] -ne 'Missing' }
$dkimStatus  = Get-EmailAuthStatus $dkimBad  { $_['DKIM signing (Exchange)'] -eq 'Enabled' -or $_['DKIM DNS (selector1 / selector2)'] -ne 'Not published' }
$dmarcStatus = Get-EmailAuthStatus $dmarcBad { $_['DMARC policy'] -ne 'Missing' }

# ----------------------------------------------------------------------------------------------
# Defender for Endpoint (separate API; delegated token for the same user, only after a browser sign-in)
# ----------------------------------------------------------------------------------------------
$mdeRows = @(); $mdeKv = [ordered]@{}; $mdeCollected = $false; $mdeNote = ''; $mdeCheck = ''
# Admin consent for exactly the three Defender permissions; works even when the app registration does not list them.
$mdeConsentUrl = "https://login.microsoftonline.com/$($ctx.TenantId)/v2.0/adminconsent?client_id=$($script:GraphClientId)&redirect_uri=http://localhost&scope=https://api.securitycenter.microsoft.com/Machine.Read%20https://api.securitycenter.microsoft.com/Software.Read%20https://api.securitycenter.microsoft.com/Vulnerability.Read"
if ($script:HasMDE -eq $false) { $mdeNote = 'Not collected: no Defender for Endpoint or Defender for Business service plan in the subscribed SKUs.' }
elseif ($script:GraphToken -and $script:GraphToken.RefreshToken) {
    Write-Host "Collecting Defender for Endpoint device status..." -ForegroundColor Cyan
    try {
        $mdeToken = Request-ResourceToken 'https://api.securitycenter.microsoft.com/Machine.Read https://api.securitycenter.microsoft.com/Software.Read https://api.securitycenter.microsoft.com/Vulnerability.Read'
        if ($mdeToken) {
            $mdeHeaders = @{ Authorization = "Bearer $mdeToken" }
            $machines = @()
            $next = 'https://api.securitycenter.microsoft.com/api/machines'
            while ($next) { $r = Invoke-RestMethod -Method Get -Uri $next -Headers $mdeHeaders -ErrorAction Stop; $machines += @($r.value); $next = $r.'@odata.nextLink' }
            $mdeCollected = $true
            $mdeRows = foreach ($m in ($machines | Sort-Object computerDnsName)) {
                $is = @{}; $tags = @()
                if ($m.riskScore -eq 'High') { $is['Risk'] = 'mde-high-risk'; $tags += 'highrisk' }
                if ($m.healthStatus -ne 'Active') { $is['Sensor health'] = 'mde-inactive'; $tags += 'inactive' }
                if ($m.exposureLevel -eq 'High') { $tags += 'highexposure' }
                if ($m.lastSeen -and ([datetime]$m.lastSeen) -lt (Get-Date).AddDays(-$StaleDays)) { $tags += 'stale' }
                [ordered]@{ 'Device' = $m.computerDnsName; 'OS' = "$($m.osPlatform) $($m.version)"; 'Onboarding' = $m.onboardingStatus; 'Sensor health' = $m.healthStatus; 'Risk' = $m.riskScore; 'Exposure' = $m.exposureLevel; 'Last seen' = (Format-DateText $m.lastSeen); 'Managed by' = $m.managedBy; 'Device value' = $m.deviceValue; 'Tags' = (Join-List $m.machineTags '; '); '_tags' = ($tags -join ' '); '_issues' = $(if ($is.Count) { $is } else { $null }) }
            }
            $mdeRows = @($mdeRows)
            $expo = try { (Invoke-RestMethod -Method Get -Uri 'https://api.securitycenter.microsoft.com/api/exposureScore' -Headers $mdeHeaders -ErrorAction Stop).score } catch { $null }
            $conf = try { (Invoke-RestMethod -Method Get -Uri 'https://api.securitycenter.microsoft.com/api/configurationScore' -Headers $mdeHeaders -ErrorAction Stop).score } catch { $null }
            $mdeKv = [ordered]@{
                'Devices onboarded to Defender for Endpoint' = @($mdeRows | Where-Object { $_['Onboarding'] -in @('Onboarded', '', $null) }).Count
                'Devices discovered on the network but not onboarded' = @($mdeRows | Where-Object { $_['Onboarding'] -notin @('Onboarded', '', $null) }).Count
                'Sensor health: active / inactive / impaired' = "$(@($mdeRows | Where-Object { $_['Sensor health'] -eq 'Active' }).Count) / $(@($mdeRows | Where-Object { $_['Sensor health'] -eq 'Inactive' }).Count) / $(@($mdeRows | Where-Object { $_['Sensor health'] -like 'Impaired*' }).Count)"
                'Devices at high risk (active alerts)'       = @($mdeRows | Where-Object { $_['Risk'] -eq 'High' }).Count
                'Devices with high exposure (vulnerabilities)' = @($mdeRows | Where-Object { $_['Exposure'] -eq 'High' }).Count
                'Organization exposure score (0-100, lower is better)' = $(if ($null -ne $expo) { [math]::Round([double]$expo, 1) } else { 'Not available' })
                'Microsoft Secure Score for Devices (configuration score)' = $(if ($null -ne $conf) { [math]::Round([double]$conf, 1) } else { 'Not available' })
            }
            $mdeIntuneWin = @($allDeviceRows | Where-Object { $_['Intune'] -eq 'Enrolled' -and $_['OS'] -match 'Windows' }).Count
            if ($mdeIntuneWin -gt $mdeRows.Count) { $mdeKv['Intune Windows devices not seen by Defender for Endpoint'] = $mdeIntuneWin - $mdeRows.Count }
        } else { $mdeNote = 'Not collected: device-code sign-in was used (no refresh token for the Defender API).'; $mdeCheck = 'devicecode' }
    } catch {
        $msg = $_.Exception.Message
        try { $body = $_.ErrorDetails.Message; if ($body) { $j = $body | ConvertFrom-Json; if ($j.error_description) { $msg = "$msg $($j.error_description)" } elseif ($j.error.message) { $msg = "$msg $($j.error.message)" } } } catch { Write-Verbose 'Defender error body was not JSON.' }
        $mdeCheck = if ($msg -match 'AADSTS65001|AADSTS650|consent') { 'consent' } elseif ($msg -match '403|Forbidden|Unauthorized|401') { 'role' } else { 'error' }
        $mdeNote = if ($mdeCheck -eq 'consent') { "The Defender for Endpoint permissions (Machine.Read, Software.Read, Vulnerability.Read) have not been approved for this app in the tenant. A Global Administrator approves them once at $mdeConsentUrl" } elseif ($msg -match '403|Forbidden|Unauthorized|401') { "Defender for Endpoint API refused the request ($($msg -replace '\s+', ' ')). The signed-in account lacks a Defender role (Security Reader), or Defender has never been set up: open https://security.microsoft.com once as an administrator to finish setup." } else { "Not collected: $($msg -replace '\s+', ' ')" }
        if ($msg -match '(?i)no active license|not licensed') { $script:HasMDE = $false; $mdeCheck = 'licence'; $mdeNote = 'Not licensed: the Defender for Endpoint API reports no active licence for this tenant (the subscription includes Defender for Office 365 but not Defender for Endpoint / Defender for Business device onboarding).' }
        Write-Verbose "Defender for Endpoint API: $mdeNote"
        if ($mdeCheck -eq 'consent') { $script:ConsentNeeded['Defender for Endpoint'] = 'https://api.securitycenter.microsoft.com/Machine.Read https://api.securitycenter.microsoft.com/Software.Read https://api.securitycenter.microsoft.com/Vulnerability.Read' }
    }
} else { $mdeNote = 'Not collected: device-code sign-in was used (no refresh token for the Defender API).'; $mdeCheck = 'devicecode' }
# Coverage: active Windows and macOS devices known to Entra ID / Intune that should report to Defender for Endpoint.
$mdeEligible = @($allDeviceRows.ToArray() | Where-Object { $_['OS'] -match '^(Windows|macOS|Mac OS)' -and $_['Enabled'] -ne $false -and [string]$_['_tags'] -notmatch '\bstale\b' })
$mdeOnboarded = @($mdeRows | Where-Object { $_['Onboarding'] -in @('Onboarded', '', $null) })
$mdeNames = @{}; foreach ($r in $mdeOnboarded) { $n = ([string]$r['Device'] -split '\.')[0].ToUpper(); if ($n) { $mdeNames[$n] = $true } }
$mdeMissingRows = @(foreach ($d in $mdeEligible) { if (-not $mdeNames.ContainsKey(([string]$d['Device']).ToUpper())) { [ordered]@{ 'Device' = $d['Device']; 'OS' = $d['OS']; 'Intune' = $d['Intune']; 'Owner / primary user' = $d['Owner / primary user']; 'Last activity' = $d['Last activity'] } } })
$mdeCovered = $mdeEligible.Count - $mdeMissingRows.Count
# Licence detail and Intune onboarding policy describe licensing, setup and deployment even when the Defender API cannot be read.
$mdePlanRx = '^WINDEFATP|^MDE_SMB|^DEFENDER_ENDPOINT|^MDE_LITE|Microsoft_Defender_For_Endpoint'
$mdeSkus = @(@($skus) | Where-Object { $_ -and $_.capabilityStatus -eq 'Enabled' -and @($_.servicePlans | Where-Object { $_.servicePlanName -match $mdePlanRx -and $_.provisioningStatus -ne 'Disabled' }).Count })
$mdeAssigned = [int](@($mdeSkus | ForEach-Object { [int]$_.consumedUnits }) | Measure-Object -Sum).Sum
$mdePurchased = [int](@($mdeSkus | ForEach-Object { [int]$_.prepaidUnits.enabled }) | Measure-Object -Sum).Sum
$mdeSeatsText = if ($mdeSkus.Count) { "$mdeAssigned of $mdePurchased licences assigned, $(Join-List ($mdeSkus | ForEach-Object { $_.skuPartNumber }) ', ')" } else { 'No licence detail' }
$mdeProvisioning = @($mdeSkus | ForEach-Object { $_.servicePlans } | Where-Object { $_.servicePlanName -match $mdePlanRx -and $_.provisioningStatus -ne 'Disabled' } | ForEach-Object { [string]$_.provisioningStatus } | Select-Object -Unique)
$mdeSetupPending = [bool]($script:HasMDE -and $mdeProvisioning.Count -and -not ($mdeProvisioning -contains 'Success'))
$mdeEdrPolicies = @(@($settingsCatalog) | Where-Object { $_ -and $_.templateReference.templateFamily -eq 'endpointSecurityEndpointDetectionAndResponse' })
$mdeEdrAssigned = @($mdeEdrPolicies | Where-Object { @($_.assignments).Count -gt 0 })
$mdeEdrText = if (-not $script:HasIntune) { 'Intune is not licensed, so there is no Intune onboarding policy to check' } elseif (-not $mdeEdrPolicies.Count) { 'No endpoint detection and response (EDR) onboarding policy in Intune' } else { "$($mdeEdrPolicies.Count) EDR onboarding policy(ies) in Intune, $($mdeEdrAssigned.Count) assigned to devices ($(Join-List ($mdeEdrPolicies | ForEach-Object { $_.name }) '; '))" }
$mdeVerifyNote = switch ($mdeCheck) {
    'consent'    { 'Device-level detail was not read because the read-only Defender permissions are not approved for the report in this tenant (approval link at the end of the run and in section 16).' }
    'role'       { 'Device-level detail was not read because Defender refused the account that ran the report (it needs Security Reader), or Defender has never been opened to finish setup.' }
    'devicecode' { 'Device-level detail was not read because the device-code sign-in cannot reach the Defender API.' }
    default      { if ($mdeNote -and -not $mdeCollected) { "Device-level detail was not read: $mdeNote" } else { '' } }
}
$mdeVerifyTodoHtml = switch ($mdeCheck) {
    'consent'    { " To see every device, a Global Administrator approves the read-only Defender permissions once: <a href='$(HtmlEnc $mdeConsentUrl)' target='_blank' rel='noopener'>approve Defender for Endpoint read access</a> (the browser then shows a localhost page that does not load; that is expected), and the report is run again." }
    'role'       { " To see every device, give the account that runs the report the <b>Security Reader</b> role and run the report again." }
    'devicecode' { ' To see every device, run the report with the normal browser sign-in (without -DeviceCode).' }
    default      { '' }
}
$mdeOnboardHtml = "Onboard devices through Intune: Endpoint security &gt; Endpoint detection and response, one policy for Windows and one for macOS, with the Defender connector turned on (<a href='https://security.microsoft.com/securitysettings/endpoints/onboarding' target='_blank' rel='noopener'>Defender onboarding settings</a>). Without Intune, use the onboarding package from the same page."
# Deployment stage: licensed? set up? assigned? deployed to every active Windows / macOS device? healthy?
$mdeAttention = @($mdeOnboarded | Where-Object { $_['Risk'] -eq 'High' -or $_['Sensor health'] -ne 'Active' }).Count
$mdeStatus = if ($script:HasMDE -eq $false) { 'Not licensed' }
    elseif ($mdeCollected) { if ($mdeOnboarded.Count -eq 0) { $(if ($mdeSetupPending) { 'Licensed, setup not finished' } else { 'Licensed, not deployed' }) } elseif ($mdeAttention) { 'Attention needed' } elseif ($mdeMissingRows.Count) { 'Partly deployed' } else { 'Deployed' } }
    elseif ($mdeSetupPending) { 'Licensed, setup not finished' }
    elseif ($mdeSkus.Count -and $mdeAssigned -eq 0) { 'Licensed, not assigned' }
    elseif ($script:HasIntune -and $mdeEdrPolicies.Count -eq 0) { 'Licensed, no onboarding policy' }
    elseif ($mdeEdrAssigned.Count) { 'Onboarding policy in place (devices not verified)' }
    else { 'Licensed, deployment not verified' }
$mdeCoverageText = if ($mdeEligible.Count) { "$mdeCovered of $($mdeEligible.Count) active Windows and macOS devices onboarded ($([math]::Round(100 * $mdeCovered / $mdeEligible.Count))%)" } else { "$($mdeOnboarded.Count) device(s) onboarded; no active Windows or macOS devices in Entra ID or Intune to compare against" }
$mdeEvidence = switch -Regex ($mdeStatus) {
    '^Not licensed'                  { 'No Defender for Endpoint or Defender for Business plan in the subscribed licences, so devices have no Microsoft endpoint detection and response (EDR). A third-party EDR product would not show here.'; break }
    '^Licensed, setup not finished'  { "$($script:MdePlan) is licensed ($mdeSeatsText), but Microsoft reports the plan as $(Join-List $mdeProvisioning ', '), so Defender has not been set up and no device is protected."; break }
    '^Licensed, not assigned'        { "$($script:MdePlan) is in the subscription, but none of its $mdePurchased licences is assigned to a user. $mdeVerifyNote"; break }
    '^Licensed, not deployed'        { "$($script:MdePlan) is licensed ($mdeSeatsText) but no devices are onboarded, so the licence is not protecting anything yet. $($mdeEligible.Count) active Windows and macOS device(s) should be onboarded."; break }
    '^(Licensed, no onboarding|Onboarding policy|Licensed, deployment)' { "$($script:MdePlan) is licensed ($mdeSeatsText). $mdeEdrText. $mdeVerifyNote"; break }
    default                          { "$($script:MdePlan): $mdeCoverageText; $($mdeKv['Devices at high risk (active alerts)']) at high risk; sensors active / inactive / impaired $($mdeKv['Sensor health: active / inactive / impaired']); exposure score $($mdeKv['Organization exposure score (0-100, lower is better)'])." }
}
$mdeEvidence = ([string]$mdeEvidence).Trim()
$mdeMeaning = switch -Regex ($mdeStatus) {
    '^Not licensed'                  { 'Devices have no Microsoft endpoint detection and response. If another product (for example CrowdStrike or SentinelOne) protects them, record it in the minutes; this report cannot see third-party tools.'; break }
    '^Licensed, setup not finished'  { 'Defender is paid for but has not been switched on, so no device is protected.'; break }
    '^Licensed, not assigned'        { 'The licence is not assigned to anyone, so devices are almost certainly not protected and the institution is not licensed to protect them.'; break }
    '^Licensed, not deployed'        { 'The licence is paid for, but no device sends data to Defender. Attacks on workstations would not be detected or contained.'; break }
    '^Licensed, no onboarding'       { 'Devices are probably not onboarded: Intune is the usual way to onboard them and it has no policy for it. Devices onboarded another way (script or Group Policy) could not be confirmed.'; break }
    '^Onboarding policy'             { 'Onboarding is set up in Intune, but the report could not confirm which devices actually report to Defender.'; break }
    '^Licensed, deployment'          { 'Defender is licensed, but the report could not see how or whether devices are onboarded.'; break }
    '^Partly deployed'               { 'Some devices are protected; the devices listed below are not, and an attacker only needs one.'; break }
    '^Attention needed'              { 'Defender is deployed, but some devices are at high risk or have stopped reporting.'; break }
    default                          { 'Every active Windows and macOS device reports to Defender and none is at high risk.' }
}
$mdeTodoHtml = switch -Regex ($mdeStatus) {
    '^Not licensed'                  { 'Decide whether to license Defender for Business (included in Microsoft 365 Business Premium, or as an add-on) or Defender for Endpoint, or record the endpoint protection product in use instead.'; break }
    '^Licensed, setup not finished'  { "Open the <a href='https://security.microsoft.com' target='_blank' rel='noopener'>Defender portal</a> as a Global or Security Administrator to finish setting up $(HtmlEnc $script:MdePlan). Then $($mdeOnboardHtml.Substring(0,1).ToLower() + $mdeOnboardHtml.Substring(1))"; break }
    '^Licensed, not assigned'        { "Assign the licence to every user whose device should be protected (Microsoft 365 admin center &gt; Users &gt; Active users &gt; Licenses and apps), then onboard devices.$mdeVerifyTodoHtml"; break }
    '^Licensed, not deployed'        { $mdeOnboardHtml; break }
    '^Licensed, no onboarding'       { "$mdeOnboardHtml$mdeVerifyTodoHtml"; break }
    '^Onboarding policy'             { "Confirm in the <a href='https://security.microsoft.com/machines' target='_blank' rel='noopener'>Defender portal</a> that every device appears.$mdeVerifyTodoHtml"; break }
    '^Licensed, deployment'          { "Confirm how devices are onboarded and check the <a href='https://security.microsoft.com/machines' target='_blank' rel='noopener'>Defender portal</a> device list.$mdeVerifyTodoHtml"; break }
    '^Partly deployed'               { 'Onboard the devices listed below, or retire them if they are no longer in use.'; break }
    '^Attention needed'              { "Investigate the high-risk devices and fix inactive sensors in the <a href='https://security.microsoft.com/machines' target='_blank' rel='noopener'>Defender portal</a>."; break }
    default                          { 'Keep it: review high-risk devices and new devices every month.' }
}
$mdeStatusHtml = "<table class='kv'><tbody><tr><th>Deployment status</th><td>$(New-StatusBadge $mdeStatus)</td></tr><tr><th>Licence</th><td>$(HtmlEnc $(if ($script:HasMDE) { $script:MdePlan } else { 'Not licensed' }))</td></tr>" +
    $(if ($script:HasMDE) { "<tr><th>Licences assigned</th><td>$(HtmlEnc $mdeSeatsText)</td></tr><tr><th>Intune onboarding policy</th><td>$(HtmlEnc $mdeEdrText)</td></tr>" } else { '' }) +
    $(if ($mdeCollected) { "<tr><th>Coverage</th><td>$(HtmlEnc $mdeCoverageText)</td></tr>" } elseif ($script:HasMDE -and $mdeVerifyNote) { "<tr><th>Device-level detail</th><td>$(HtmlEnc $mdeVerifyNote)</td></tr>" } else { '' }) +
    "<tr><th>What this means</th><td>$(HtmlEnc $mdeMeaning)</td></tr><tr><th>What to do</th><td>$mdeTodoHtml</td></tr></tbody></table>"
# Console: lead with what was found about licensing, setup and deployment; the permission gap is a secondary note.
$mdeConsole = "Defender for Endpoint: $mdeStatus. $mdeEvidence"
if ((Get-StatusKind $mdeStatus) -in @('warn', 'mid')) { Write-Warning $mdeConsole } else { Write-Host $mdeConsole -ForegroundColor DarkGray }
$MdeFilters = @(@{ Label = 'High risk'; Tag = 'highrisk' }, @{ Label = 'High exposure'; Tag = 'highexposure' }, @{ Label = 'Inactive sensor'; Tag = 'inactive' }, @{ Label = 'Stale'; Tag = 'stale' })

$UserByUpn = @{}; foreach ($u in $users) { if ($u.userPrincipalName) { $UserByUpn[([string]$u.userPrincipalName).ToLower()] = $u } }
# ----------------------------------------------------------------------------------------------
# Possible indications of compromise (section 2): what an attacker typically leaves behind, from the
# directory audit log, the sign-in log, Entra ID Protection, app registrations and Exchange (collected above).
# Nothing here is a confirmed incident; each item is something to confirm with the people involved.
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting indications of compromise (last $AuditDays days)..." -ForegroundColor Cyan
$iocSinceDate = (Get-Date).AddDays(-$AuditDays)
function ConvertTo-IocAuditRow {
    param($e, [string]$Issue)
    $actor = if ($e.initiatedBy.user) { $(if ($e.initiatedBy.user.displayName) { "$($e.initiatedBy.user.displayName) ($($e.initiatedBy.user.userPrincipalName))" } else { [string]$e.initiatedBy.user.userPrincipalName }) } elseif ($e.initiatedBy.app) { "App: $($e.initiatedBy.app.displayName)" } else { 'Unknown' }
    $targets = @($e.targetResources | Where-Object { $_ })
    $tUser = $targets | Where-Object { $_.type -eq 'User' } | Select-Object -First 1
    $tOther = $targets | Where-Object { $_.type -ne 'User' } | Select-Object -First 1
    $target = Join-List (@($tUser.userPrincipalName, $tUser.displayName, $tOther.displayName) | Where-Object { $_ } | Select-Object -First 2) ' / '
    $detail = @()
    foreach ($t in $targets) {
        foreach ($mp in @($t.modifiedProperties | Where-Object { $_ })) {
            if ($mp.displayName -match '^(Role\.DisplayName|Role\.TemplateId|Domain|LiveType|FederationBrandName|PasswordPolicies|StrongAuthenticationMethod|StrongAuthenticationUserDetails|KeyDescription|KeyType|KeyUsage|Included Updated Properties|AppAddress|TargetId\.UserType|Group\.DisplayName|ConditionalAccessPolicy|PolicyName|AuthenticationMethod|IsAdminConsent|AccountEnabled|UserPrincipalName)$') {
                $nv = ([string]$mp.newValue) -replace '^"|"$', ''
                if ($nv -and $nv -ne '[]') { $detail += "$($mp.displayName -replace '^Role\.', ''): $nv" }
            }
        }
    }
    $self = $tUser -and $e.initiatedBy.user -and $tUser.userPrincipalName -and ([string]$tUser.userPrincipalName -eq [string]$e.initiatedBy.user.userPrincipalName)
    [ordered]@{ 'Date (local)' = (Format-DateText $e.activityDateTime); 'Activity' = $e.activityDisplayName; 'Performed by' = $actor; 'Target' = $target; 'Detail' = (Join-List ($detail | Select-Object -First 4) '; '); 'Result' = $e.result; '_issues' = @{ 'Activity' = $Issue }; '_tags' = $(if ($self) { 'self' } else { 'other' }) }
}
function Get-IocAudit {
    param([string]$Filter, [string]$What, [string]$Issue, [scriptblock]$Keep)
    $ev = Get-GraphAllSafe "v1.0/auditLogs/directoryAudits?`$filter=$([uri]::EscapeDataString("activityDateTime ge $since and ($Filter)"))&`$top=999" $What -MaxItems 2000
    if ($Keep) { $ev = @($ev | Where-Object $Keep) }
    return @($ev | Sort-Object activityDateTime -Descending | ForEach-Object { ConvertTo-IocAuditRow $_ $Issue })
}
$iocRoleRows   = Get-IocAudit "category eq 'RoleManagement'" 'role assignment audit events' 'ioc-role-add' { $_.activityDisplayName -match '^Add (member|eligible member|scoped member) to role' -and $_.result -eq 'success' }
$iocMfaRows    = Get-IocAudit "loggedByService eq 'Authentication Methods'" 'security info audit events' 'ioc-mfa-change' { $_.activityDisplayName -match 'registered|deleted|changed default|Admin (registered|deleted|updated)' -and $_.result -eq 'success' }
$iocPwRows     = Get-IocAudit "activityDisplayName eq 'Reset user password' or activityDisplayName eq 'Reset password (by admin)' or activityDisplayName eq 'Change user password' or activityDisplayName eq 'Set force change user password' or activityDisplayName eq 'Reset password (self-service)'" 'password audit events' 'ioc-pw-reset' { $_.result -eq 'success' }
$iocPwAdminRows = @($iocPwRows | Where-Object { $_['_tags'] -eq 'other' })
$iocCaRows     = Get-IocAudit "activityDisplayName eq 'Add conditional access policy' or activityDisplayName eq 'Update conditional access policy' or activityDisplayName eq 'Delete conditional access policy' or activityDisplayName eq 'Add named location' or activityDisplayName eq 'Update named location' or activityDisplayName eq 'Delete named location' or activityDisplayName eq 'Update policy' or activityDisplayName eq 'Update authorization policy' or activityDisplayName eq 'Update authentication methods policy'" 'policy change audit events' 'ioc-ca-change' { $_.result -eq 'success' }
$iocDomainRows = Get-IocAudit "activityDisplayName eq 'Set federation settings on domain' or activityDisplayName eq 'Set domain authentication' or activityDisplayName eq 'Add unverified domain' or activityDisplayName eq 'Verify domain' or activityDisplayName eq 'Delete domain' or activityDisplayName eq 'Update domain' or activityDisplayName eq 'Add domain to company' or activityDisplayName eq 'Remove domain from company'" 'domain audit events' 'ioc-domain-change' { $_.result -eq 'success' }
$iocAppRows    = Get-IocAudit "activityDisplayName eq 'Add service principal credentials' or activityDisplayName eq 'Remove service principal credentials' or activityDisplayName eq 'Add owner to application' or activityDisplayName eq 'Add owner to service principal' or activityDisplayName eq 'Add application' or activityDisplayName eq 'Update application – Certificates and secrets management' or activityDisplayName eq 'Update application - Certificates and secrets management' or activityDisplayName eq 'Add app role assignment grant to user'" 'application change audit events' 'ioc-app-cred' { $_.result -eq 'success' }
$iocUserRows   = Get-IocAudit "activityDisplayName eq 'Add user' or activityDisplayName eq 'Invite external user' or activityDisplayName eq 'Enable account' or activityDisplayName eq 'Delete user' or activityDisplayName eq 'Restore user' or activityDisplayName eq 'Hard Delete user'" 'account lifecycle audit events' 'ioc-new-user' { $_.result -eq 'success' }
foreach ($r in $iocUserRows) { if ($r['Activity'] -eq 'Invite external user') { $r['_issues'] = @{ 'Activity' = 'ioc-guest' } } }
$iocAuditAvailable = -not (Get-GraphFailure 'role assignment audit events')

# App registrations with secrets or certificates added in the window
$appRegs = Get-GraphAllSafe 'v1.0/applications?$select=id,appId,displayName,createdDateTime,passwordCredentials,keyCredentials&$top=999' 'application registrations'
# who added each credential: the audit log names the actor on "Update application - Certificates and secrets management" / "Add service principal credentials"
$CredActorByApp = @{}
foreach ($ev in $iocAppRows) {
    if ($ev['Activity'] -match 'Certificates and secrets|service principal credentials') {
        $appName = ([string]$ev['Target'] -split ' / ')[0]
        if ($appName) { if (-not $CredActorByApp.ContainsKey($appName)) { $CredActorByApp[$appName] = @() }; $CredActorByApp[$appName] += "$($ev['Performed by']) on $($ev['Date (local)'])" }
    }
}
$iocNewCredRows = @()
foreach ($a in $appRegs) {
    foreach ($c in @($a.passwordCredentials) + @($a.keyCredentials)) {
        if ($c -and $c.startDateTime -and ([datetime]$c.startDateTime) -ge $iocSinceDate) {
            $addedBy = if ($CredActorByApp.ContainsKey([string]$a.displayName)) { Join-List (@($CredActorByApp[[string]$a.displayName]) | Select-Object -Unique -First 3) '; ' } else { 'Not in audit log window' }
            $iocNewCredRows += [ordered]@{ 'Application' = $a.displayName; 'Credential' = $(if ($c.secretText -or $c.hint -or $c.PSObject.Properties['hint']) { 'Client secret' } elseif ($c.type) { "Certificate ($($c.type))" } else { 'Credential' }); 'Name' = $c.displayName; 'Added by' = $addedBy; 'Valid from' = (Format-DateText $c.startDateTime); 'Expires' = (Format-DateText $c.endDateTime); '_issues' = @{ 'Credential' = 'ioc-app-cred' } }
        }
    }
}
# Enterprise applications (service principals) added in the window, with the permissions already resolved for section 10
$AppRowByName = @{}; foreach ($r in $appRows) { if ($r -and $r['Application']) { $AppRowByName[[string]$r['Application']] = $r } }
$iocNewAppRows = @($servicePrincipals | Where-Object { $_.createdDateTime -and ([datetime]$_.createdDateTime) -ge $iocSinceDate } | Sort-Object createdDateTime -Descending | ForEach-Object {
    $ar = $AppRowByName[[string]$_.displayName]
    [ordered]@{ 'Application' = $_.displayName; 'Publisher' = $(if ($ar) { $ar['Publisher'] } elseif ($_.publisherName) { $_.publisherName } else { 'Unknown' }); 'Added' = (Format-DateText $_.createdDateTime); 'Added / consented by' = $(if ($ConsentBySp.ContainsKey($_.id)) { Join-List (@($ConsentBySp[$_.id]) | Select-Object -First 3) ' | ' } elseif ($ar) { $ar['Approved by / when'] } else { 'Not in audit log window' }); 'Delegated permissions' = $(if ($ar) { Join-List (@($ar['Admin-consented delegated permissions'], $ar['User-consented delegated permissions']) | Where-Object { $_ -and $_ -ne 'None' }) '; ' } else { '' }); 'Application permissions' = $(if ($ar) { $ar['Application permissions'] } else { '' }); '_issues' = @{ 'Application' = 'ioc-new-app' } }
})
# Accounts created in the window (from the user list already collected)
$iocNewUserRows = @($users | Where-Object { $_.createdDateTime -and ([datetime]$_.createdDateTime) -ge $iocSinceDate } | Sort-Object createdDateTime -Descending | ForEach-Object {
    [ordered]@{ 'User' = $_.displayName; 'UPN' = $_.userPrincipalName; 'Type' = $_.userType; 'Enabled' = $_.accountEnabled; 'Created' = (Format-DateText $_.createdDateTime); 'Licensed' = $(if (@($_.assignedLicenses).Count) { 'Yes' } else { 'No' }); 'Admin' = $(if ($AdminIds.ContainsKey($_.id)) { 'Yes' } else { 'No' }); '_issues' = @{ 'Created' = $(if ($_.userType -eq 'Guest') { 'ioc-guest' } else { 'ioc-new-user' }) }; '_tags' = $(if ($_.userType -eq 'Guest') { 'guest' } else { 'member' }) }
})

# Sign-in log (Entra ID P1): countries, legacy protocols, MFA denials, Conditional Access blocks, password spray, lockouts
$signInDays = [Math]::Min($AuditDays, 30)
$signInSince = (Get-Date).AddDays(-$signInDays).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$signInSelect = 'id,createdDateTime,userPrincipalName,userDisplayName,appDisplayName,clientAppUsed,ipAddress,location,status,riskLevelDuringSignIn,riskLevelAggregated,riskState,conditionalAccessStatus,deviceDetail'
$iocSignInsCollected = $false
$countryRows = @(); $multiCountryRows = @(); $legacySignInRows = @(); $mfaDeniedRows = @(); $caBlockedRows = @(); $sprayRows = @(); $lockoutRows = @(); $riskySignInRows = @()
$iocSignInNote = ''
if ($script:HasEntraP1 -eq $false) { $iocSignInNote = "Sign-in log analysis requires Entra ID P1 or P2, which is not licensed in this tenant." }
else {
    Write-Host "Analysing sign-ins (last $signInDays days)..." -ForegroundColor Cyan
    $okSignIns = Get-GraphAllSafe "v1.0/auditLogs/signIns?`$filter=$([uri]::EscapeDataString("createdDateTime ge $signInSince and status/errorCode eq 0"))&`$select=$signInSelect&`$top=999" 'successful sign-ins' -MaxItems 8000
    if (-not (Get-GraphFailure 'successful sign-ins')) {
        $iocSignInsCollected = $true
        $capped = $okSignIns.Count -ge 8000
        $byCountry = @{}; $userCountries = @{}
        foreach ($si in $okSignIns) {
            $c = [string]$si.location.countryOrRegion; if (-not $c) { $c = 'Unknown' }
            if (-not $byCountry.ContainsKey($c)) { $byCountry[$c] = @{ Count = 0; Users = @{} } }
            $byCountry[$c].Count++; $byCountry[$c].Users[[string]$si.userPrincipalName] = $true
            $u = [string]$si.userPrincipalName
            if (-not $userCountries.ContainsKey($u)) { $userCountries[$u] = @{} }
            if (-not $userCountries[$u].ContainsKey($c)) { $userCountries[$u][$c] = @{ Count = 0; Last = $null; Cities = @{}; Ips = @{} } }
            $userCountries[$u][$c].Count++; $userCountries[$u][$c].Cities[[string]$si.location.city] = $true; $userCountries[$u][$c].Ips[[string]$si.ipAddress] = $true
            if (-not $userCountries[$u][$c].Last -or ([datetime]$si.createdDateTime) -gt $userCountries[$u][$c].Last) { $userCountries[$u][$c].Last = [datetime]$si.createdDateTime }
        }
        $homeCountry = ($byCountry.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | Select-Object -First 1).Key
        $countryRows = @($byCountry.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object {
            [ordered]@{ 'Country' = $_.Key; 'Successful sign-ins' = $_.Value.Count; 'Distinct users' = $_.Value.Users.Count; 'Share' = "$([math]::Round(100 * $_.Value.Count / [Math]::Max(1, $okSignIns.Count)))%"; '_issues' = $(if ($_.Key -ne $homeCountry -and $_.Key -ne 'Unknown') { @{ 'Country' = 'ioc-multi-country' } } else { $null }) }
        })
        foreach ($u in ($userCountries.Keys | Sort-Object)) {
            $cs = @($userCountries[$u].Keys | Where-Object { $_ -ne 'Unknown' })
            if ($cs.Count -gt 1) {
                $uo = $UserByUpn[$u.ToLower()]
                $multiCountryRows += [ordered]@{ 'User' = $(if ($uo) { $uo.displayName } else { $u }); 'UPN' = $u; 'Countries' = (Join-List ($cs | Sort-Object { -$userCountries[$u][$_].Count } | ForEach-Object { "$_ ($($userCountries[$u][$_].Count))" }) ', '); 'Cities' = (Join-List (@($cs | ForEach-Object { $userCountries[$u][$_].Cities.Keys }) | Where-Object { $_ } | Select-Object -Unique -First 6) ', '); 'Last sign-in outside home country' = (Format-DateText (($cs | Where-Object { $_ -ne $homeCountry } | ForEach-Object { $userCountries[$u][$_].Last }) | Sort-Object -Descending | Select-Object -First 1)); 'Admin' = $(if ($uo -and $AdminIds.ContainsKey($uo.id)) { 'Yes' } else { 'No' }); '_issues' = @{ 'Countries' = 'ioc-multi-country' }; '_tags' = $(if ($uo -and $AdminIds.ContainsKey($uo.id)) { 'admin' } else { 'user' }) }
            }
        }
        $legacyClients = 'Exchange ActiveSync|IMAP4|POP3|^SMTP|Authenticated SMTP|Exchange Web Services|MAPI Over HTTP|Offline Address Book|Other clients|Outlook Anywhere|Reporting Web Services|Universal Outlook|AutoDiscover|Exchange Online PowerShell'
        $legacyBy = @{}
        foreach ($si in ($okSignIns | Where-Object { [string]$_.clientAppUsed -match $legacyClients })) {
            $k = "$($si.userPrincipalName)|$($si.clientAppUsed)"
            if (-not $legacyBy.ContainsKey($k)) { $legacyBy[$k] = @{ User = $si.userPrincipalName; Client = $si.clientAppUsed; Count = 0; Last = $null; Apps = @{}; Countries = @{} } }
            $legacyBy[$k].Count++; $legacyBy[$k].Apps[[string]$si.appDisplayName] = $true; $legacyBy[$k].Countries[[string]$si.location.countryOrRegion] = $true
            if (-not $legacyBy[$k].Last -or ([datetime]$si.createdDateTime) -gt $legacyBy[$k].Last) { $legacyBy[$k].Last = [datetime]$si.createdDateTime }
        }
        $legacySignInRows = @($legacyBy.Values | Sort-Object { $_.Count } -Descending | ForEach-Object { [ordered]@{ 'User' = $_.User; 'Protocol / client' = $_.Client; 'Successful sign-ins' = $_.Count; 'Applications' = (Join-List ($_.Apps.Keys | Where-Object { $_ })); 'Countries' = (Join-List ($_.Countries.Keys | Where-Object { $_ })); 'Last' = (Format-DateText $_.Last); '_issues' = @{ 'Protocol / client' = 'ioc-legacy-signin' } } })
        $riskySignInRows = @($okSignIns | Where-Object { ([string]$_.riskLevelDuringSignIn -match '^(low|medium|high)$') -or ([string]$_.riskLevelAggregated -match '^(medium|high)$') } | Sort-Object createdDateTime -Descending | Select-Object -First 100 | ForEach-Object { [ordered]@{ 'Date (local)' = (Format-DateText $_.createdDateTime); 'User' = $_.userPrincipalName; 'Risk during sign-in' = $_.riskLevelDuringSignIn; 'Aggregated risk' = $_.riskLevelAggregated; 'Risk state' = $_.riskState; 'Application' = $_.appDisplayName; 'IP address' = $_.ipAddress; 'Location' = (Join-List (@($_.location.city, $_.location.countryOrRegion) | Where-Object { $_ }) ', '); '_issues' = @{ 'Risk during sign-in' = 'ioc-risk-detection' } } })

        function Get-FailedSignIns { param([int]$Code, [string]$What) return @(Get-GraphAllSafe "v1.0/auditLogs/signIns?`$filter=$([uri]::EscapeDataString("createdDateTime ge $signInSince and status/errorCode eq $Code"))&`$select=$signInSelect&`$top=999" $What -MaxItems 5000) }
        function Group-SignInsByUser { param($SignIns, [string]$Issue)
            $g = @{}
            foreach ($si in $SignIns) { $u = [string]$si.userPrincipalName; if (-not $g.ContainsKey($u)) { $g[$u] = @{ Count = 0; Last = $null; Ips = @{}; Countries = @{}; Apps = @{} } }; $g[$u].Count++; $g[$u].Ips[[string]$si.ipAddress] = $true; $g[$u].Countries[[string]$si.location.countryOrRegion] = $true; $g[$u].Apps[[string]$si.appDisplayName] = $true; if (-not $g[$u].Last -or ([datetime]$si.createdDateTime) -gt $g[$u].Last) { $g[$u].Last = [datetime]$si.createdDateTime } }
            return @($g.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object { $uo = $UserByUpn[$_.Key.ToLower()]; [ordered]@{ 'User' = $_.Key; 'Attempts' = $_.Value.Count; 'Distinct IP addresses' = $_.Value.Ips.Count; 'Countries' = (Join-List ($_.Value.Countries.Keys | Where-Object { $_ })); 'Applications' = (Join-List ($_.Value.Apps.Keys | Where-Object { $_ } | Select-Object -First 4)); 'Last' = (Format-DateText $_.Value.Last); 'Admin' = $(if ($uo -and $AdminIds.ContainsKey($uo.id)) { 'Yes' } else { 'No' }); '_issues' = @{ 'Attempts' = $Issue } } })
        }
        $mfaDeniedRows = Group-SignInsByUser (Get-FailedSignIns 500121 'MFA failed sign-ins') 'ioc-mfa-denied'
        $caBlockedRows = Group-SignInsByUser (Get-FailedSignIns 53003 'Conditional Access blocked sign-ins') 'ioc-ca-blocked'
        $lockoutRows   = Group-SignInsByUser (Get-FailedSignIns 50053 'locked-out sign-ins') 'ioc-lockout'
        $badPw = Get-FailedSignIns 50126 'wrong-password sign-ins'
        $byIp = @{}
        foreach ($si in $badPw) { $ip = [string]$si.ipAddress; if (-not $ip) { continue }; if (-not $byIp.ContainsKey($ip)) { $byIp[$ip] = @{ Count = 0; Users = @{}; Country = [string]$si.location.countryOrRegion; Last = $null } }; $byIp[$ip].Count++; $byIp[$ip].Users[[string]$si.userPrincipalName] = $true; if (-not $byIp[$ip].Last -or ([datetime]$si.createdDateTime) -gt $byIp[$ip].Last) { $byIp[$ip].Last = [datetime]$si.createdDateTime } }
        $sprayRows = @($byIp.GetEnumerator() | Where-Object { $_.Value.Users.Count -ge 3 -or $_.Value.Count -ge 20 } | Sort-Object { $_.Value.Users.Count } -Descending | Select-Object -First 50 | ForEach-Object { [ordered]@{ 'Source IP address' = $_.Key; 'Country' = $_.Value.Country; 'Failed passwords' = $_.Value.Count; 'Distinct accounts targeted' = $_.Value.Users.Count; 'Accounts (sample)' = (Join-List ($_.Value.Users.Keys | Select-Object -First 5)); 'Last attempt' = (Format-DateText $_.Value.Last); '_issues' = @{ 'Distinct accounts targeted' = 'ioc-spray' } } })
        $badPwTotal = $badPw.Count; $badPwUsers = @($badPw | ForEach-Object { $_.userPrincipalName } | Select-Object -Unique).Count
        if ($capped) { $iocSignInNote = "Successful sign-ins were capped at 8,000 records for the country analysis; the tenant has more in the window." }
    } else { $iocSignInNote = "Sign-in log could not be read: $(Get-GraphFailure 'successful sign-ins')" }
}

# Entra ID Protection (P2): risky users and risk detections
$riskyUserRows = @(); $riskDetectionRows = @(); $iocRiskNote = ''
if ($script:HasEntraP2 -eq $false) { $iocRiskNote = 'Entra ID Protection (risky users, risk detections) requires Entra ID P2, which is not licensed in this tenant.' }
else {
    $riskyUsers = Get-GraphAllSafe "v1.0/identityProtection/riskyUsers?`$filter=$([uri]::EscapeDataString("riskState eq 'atRisk' or riskState eq 'confirmedCompromised'"))" 'risky users'
    $riskyUserRows = @($riskyUsers | Sort-Object riskLastUpdatedDateTime -Descending | ForEach-Object { [ordered]@{ 'User' = $_.userDisplayName; 'UPN' = $_.userPrincipalName; 'Risk level' = $_.riskLevel; 'Risk state' = $_.riskState; 'Detail' = $_.riskDetail; 'Last updated' = (Format-DateText $_.riskLastUpdatedDateTime); '_issues' = @{ 'Risk level' = 'ioc-risky-user' } } })
    $riskDetections = Get-GraphAllSafe "v1.0/identityProtection/riskDetections?`$filter=$([uri]::EscapeDataString("detectedDateTime ge $since"))" 'risk detections' -MaxItems 500
    $riskDetectionRows = @($riskDetections | Sort-Object detectedDateTime -Descending | Select-Object -First 200 | ForEach-Object { [ordered]@{ 'Detected (local)' = (Format-DateText $_.detectedDateTime); 'User' = $_.userPrincipalName; 'Detection' = $_.riskEventType; 'Level' = $_.riskLevel; 'State' = $_.riskState; 'IP address' = $_.ipAddress; 'Location' = (Join-List (@($_.location.city, $_.location.countryOrRegion) | Where-Object { $_ }) ', '); '_issues' = @{ 'Detection' = 'ioc-risk-detection' } } })
    if (Get-GraphFailure 'risky users') { $iocRiskNote = "Entra ID Protection could not be read: $(Get-GraphFailure 'risky users')" }
}

# Known-abused ("traitorware") OAuth applications seen in business email compromise: legitimate products attackers
# consent to for mailbox export, bulk mailing, contact harvesting and file exfiltration. Source lists: Huntress RogueApps
# and the community MaliciousOauthAppDetections set; matched by application (client) ID first, then by name.
$RogueApps = @{
    'ff8d92dc-3d82-41d6-bcbd-b9174d163620' = @{ Name = 'PerfectData Software (Email Backup Wizard)'; Tactic = 'Mailbox export / exfiltration via EWS' }
    '2ef68ccc-8a4d-42ff-ae88-2d7bb89ad139' = @{ Name = 'Mail_Backup (renamed PerfectData Software)'; Tactic = 'Mailbox export / exfiltration' }
    'a245e8c0-b53c-4b67-9b45-751d1dff8e6b' = @{ Name = 'Newsletter Software Supermailer'; Tactic = 'Bulk phishing sent from the compromised mailbox' }
    '4761b959-9780-4c2d-87a3-512b4638f767' = @{ Name = 'rclone'; Tactic = 'OneDrive / SharePoint bulk download and exfiltration' }
    'e9a7fea1-1cc0-4cd9-a31b-9137ca5deedd' = @{ Name = 'eM Client'; Tactic = 'Mailbox sync, hidden sent items, inbox rules, persistence' }
    'a43e5392-f48b-46a4-a0f1-098b5eeb4757' = @{ Name = 'CloudSponge'; Tactic = 'Contact / address book harvesting' }
    '858d7e42-35f0-44b7-9033-df309239a47f' = @{ Name = 'ZoomInfo Login'; Tactic = 'Address book exfiltration' }
    '497ac034-5120-4c1a-929a-0351f5c09918' = @{ Name = 'ZoomInfo Communitiez Login'; Tactic = 'Contact and mail reconnaissance for phishing' }
    'caffae8c-0882-4c81-9a27-d1803af53a40' = @{ Name = 'SigParser'; Tactic = 'Address book exfiltration' }
    '77468577-4f6e-40e7-b745-11d3d0c28095' = @{ Name = 'Fastmail'; Tactic = 'Mailbox exfiltration and persistence' }
    '179d5108-412b-4c95-8e34-06786784ab39' = @{ Name = 'PostBox'; Tactic = 'Mailbox exfiltration and persistence (IMAP/POP/SMTP)' }
    '946c777c-bc85-489e-b034-392389ae23d6' = @{ Name = 'Spike'; Tactic = 'Mailbox exfiltration, persistence, phishing' }
}
$RogueAppNamePattern = '(?i)^(PERFECTDATA SOFTWARE|Mail_Backup|Email Backup Wizard|Newsletter Software Supermailer|rclone|eM Client|CloudSponge|SigParser|Fastmail|PostBox|Spike|Zoominfo( Communitiez)? Login)$'
$rogueAppRows = @()
foreach ($sp in $servicePrincipals) {
    $hit = $null
    if ($sp.appId -and $RogueApps.ContainsKey([string]$sp.appId)) { $hit = $RogueApps[[string]$sp.appId] }
    elseif (([string]$sp.displayName) -match $RogueAppNamePattern) { $hit = @{ Name = $sp.displayName; Tactic = 'Known-abused application (matched by name)' } }
    if ($hit) {
        $ar = $AppRowByName[[string]$sp.displayName]
        $rogueAppRows += [ordered]@{ 'Application' = $sp.displayName; 'Known as' = $hit.Name; 'Typical abuse' = $hit.Tactic; 'Added' = (Format-DateText $sp.createdDateTime); 'Enabled' = $sp.accountEnabled; 'Permissions granted' = $(if ($ar) { Join-List (@($ar['Admin-consented delegated permissions'], $ar['User-consented delegated permissions'], $ar['Application permissions']) | Where-Object { $_ -and $_ -ne 'None' }) '; ' } else { '' }); 'Who consented' = $(if ($ar) { $ar['Approved by / when'] } else { '' }); '_issues' = @{ 'Known as' = 'ioc-rogue-app' } }
    }
}
$rogueAppRows = @($rogueAppRows)
# Applications from unverified publishers holding mail, file, contact or directory permissions
$RiskyPermPattern = '(?i)Mail\.|MailboxSettings|Contacts\.|Files\.|Sites\.|EWS\.|full_access_as_app|IMAP|POP\.|SMTP\.|Directory\.Read|User\.Read\.All|offline_access.*Mail'
$unverifiedAppRows = @($appRows | Where-Object { $_ -and ([string]$_['Publisher']) -match 'unverified|Unknown' -and ((@($_['Admin-consented delegated permissions'], $_['User-consented delegated permissions'], $_['Application permissions']) -join ' ') -match $RiskyPermPattern) -and ([string]$_['Application']) -notmatch $RogueAppNamePattern } | ForEach-Object {
    [ordered]@{ 'Application' = $_['Application']; 'Publisher' = $_['Publisher']; 'Added to tenant' = $_['Added to tenant']; 'Delegated permissions' = (Join-List (@($_['Admin-consented delegated permissions'], $_['User-consented delegated permissions']) | Where-Object { $_ -and $_ -ne 'None' }) '; '); 'Application permissions' = $_['Application permissions']; 'Who consented' = $_['Approved by / when']; '_issues' = @{ 'Publisher' = 'ioc-unverified-app' } }
})
# External forwarding already measured in the Exchange section; it is also the single most common compromise trace
$fwdMailboxExternal = @($fwdMailboxRows | Where-Object { $_ -and $_['External'] -eq 'Yes' })
$fwdRuleExternal    = @($fwdRuleRows | Where-Object { $_ -and $_['External'] -eq 'Yes' })
$fwdTransportAll    = @($fwdTransportRows | Where-Object { $_ })


# ----------------------------------------------------------------------------------------------
# Microsoft Teams (section 17): guest and external access policy, every team with owners, members, guests, activity
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting Microsoft Teams..." -ForegroundColor Cyan
$teamGroups = Get-GraphAllSafe "v1.0/groups?`$filter=$([uri]::EscapeDataString("resourceProvisioningOptions/Any(x:x eq 'Team')"))&`$select=id,displayName,visibility,createdDateTime,description,mail,classification,membershipRule,groupTypes&`$top=999" 'teams'
$teamsActivity = @{}
foreach ($r in (Get-GraphReportCsv "getTeamsTeamActivityDetail(period='D90')" 'Teams activity report')) { if ($r.'Team Id') { $teamsActivity[[string]$r.'Team Id'] = $r } }
$groupSettings = Get-GraphAllSafe 'v1.0/groupSettings' 'group settings'
$unified = $groupSettings | Where-Object { $_.displayName -eq 'Group.Unified' } | Select-Object -First 1
function Get-GroupSettingValue { param([string]$Name) if (-not $unified) { return $null } ; $v = ($unified.values | Where-Object { $_.name -eq $Name } | Select-Object -First 1).value; return $v }
$allowAddGuests = Get-GroupSettingValue 'AllowToAddGuests'; $allowGuestsAccess = Get-GroupSettingValue 'AllowGuestsToAccessGroups'
$teamRows = @(); $teamsWithGuests = 0; $teamsPublic = 0; $teamsNoOwner = 0; $teamsInactive = 0; $teamsArchived = 0; $teamGuestTotal = 0; $TeamMembersByName = @{}
$teamLimit = 400; $teamIdx = 0
foreach ($g in ($teamGroups | Sort-Object displayName)) {
    $teamIdx++
    if ($teamIdx -gt $teamLimit) { break }
    Write-Progress -Activity 'Reading teams' -Status $g.displayName -PercentComplete ([int](100 * $teamIdx / [Math]::Max(1, [Math]::Min($teamLimit, @($teamGroups).Count))))
    $owners  = @(Get-GraphAllSafe "v1.0/groups/$($g.id)/owners?`$select=id,displayName,userPrincipalName" "owners of $($g.displayName)")
    $members = @(Get-GraphAllSafe "v1.0/groups/$($g.id)/members?`$select=id,displayName,userPrincipalName,userType&`$top=999" "members of $($g.displayName)")
    $guests  = @($members | Where-Object { $_.userType -eq 'Guest' -or ([string]$_.userPrincipalName) -match '#EXT#' })
    $TeamMembersByName[[string]$g.displayName] = @(@($members) + @($owners) | ForEach-Object { $(if ($_.userPrincipalName) { ([string]$_.userPrincipalName).ToLower() } elseif ($_.mail) { ([string]$_.mail).ToLower() }) } | Where-Object { $_ } | Select-Object -Unique)
    $team    = (Get-GraphAllSafe "v1.0/teams/$($g.id)?`$select=isArchived,memberSettings,guestSettings,messagingSettings" "settings of team $($g.displayName)" -NoPaging) | Select-Object -First 1
    $act     = $teamsActivity[[string]$g.id]
    $lastAct = if ($act -and $act.'Last Activity Date') { try { [datetime]$act.'Last Activity Date' } catch { $null } } else { $null }
    $inactive = ($null -eq $lastAct -and $g.createdDateTime -and ([datetime]$g.createdDateTime) -lt (Get-Date).AddDays(-$StaleDays)) -or ($lastAct -and $lastAct -lt (Get-Date).AddDays(-$StaleDays))
    $is = @{}; $tags = @()
    if ($guests.Count) { $is['Guests'] = 'teams-guests'; $tags += 'guests'; $teamsWithGuests++; $teamGuestTotal += $guests.Count }
    if ($g.visibility -eq 'Public') { $is['Visibility'] = 'teams-public'; $tags += 'public'; $teamsPublic++ }
    if ($owners.Count -eq 0) { $is['Owners'] = 'teams-no-owner'; $tags += 'noowner'; $teamsNoOwner++ } elseif ($owners.Count -eq 1) { $is['Owners'] = 'teams-single-owner'; $tags += 'singleowner' }
    if ($inactive) { $is['Last activity'] = 'teams-inactive'; $tags += 'inactive'; $teamsInactive++ }
    if ($team -and $team.isArchived) { $tags += 'archived'; $teamsArchived++ }
    $teamRows += [ordered]@{
        'Team' = $g.displayName; 'Visibility' = $(if ($g.visibility) { $g.visibility } else { 'Private' }); 'Archived' = $(if ($team) { [bool]$team.isArchived } else { '' })
        'Owners' = $(if ($owners.Count -le 3) { Join-List ($owners | ForEach-Object { $_.displayName }) } else { "$($owners.Count) owners" }); 'Members' = $members.Count; 'Guests' = $guests.Count
        'Guest names' = (Join-List ($guests | ForEach-Object { $_.displayName } | Select-Object -First 5) '; ')
        'Last activity' = $(if ($lastAct) { $lastAct.ToString('yyyy-MM-dd') } elseif ($act) { 'None in 90 days' } else { '' }); 'Active users (90 days)' = $(if ($act) { $act.'Active Users' } else { '' })
        'Members may add apps / connectors' = $(if ($team -and $team.memberSettings) { "$($team.memberSettings.allowAddRemoveApps) / $($team.memberSettings.allowCreateUpdateRemoveConnectors)" } else { '' })
        'Guests may create channels' = $(if ($team -and $team.guestSettings) { [bool]$team.guestSettings.allowCreateUpdateChannels } else { '' })
        'Created' = (Format-DateText $g.createdDateTime); 'Classification' = $g.classification
        '_issues' = $is; '_tags' = ($tags -join ' ')
    }
}
Write-Progress -Activity 'Reading teams' -Completed
$teamRows = @($teamRows)
$TeamFilters = @(@{ Label = 'Guests'; Tag = 'guests' }, @{ Label = 'Public'; Tag = 'public' }, @{ Label = 'No owner'; Tag = 'noowner' }, @{ Label = 'Single owner'; Tag = 'singleowner' }, @{ Label = "Inactive ($StaleDays days)"; Tag = 'inactive' }, @{ Label = 'Archived'; Tag = 'archived' })
$teamsConcealed = $teamsActivity.Count -and (Test-ConcealedReportNames @($teamsActivity.Values) 'Team Name')

# Tenant-wide Teams policies need the MicrosoftTeams module (optional) and the Skype and Teams Tenant Admin API permission
$teamsPolicyKv = [ordered]@{}; $teamsPolicyIssues = @{}; $teamsPolicyNote = ''; $teamsPolicyCollected = $false
if (Get-Module -ListAvailable MicrosoftTeams) {
    if ($script:GraphToken -and $script:GraphToken.RefreshToken) {
        try {
            Write-Host "Collecting Teams tenant policies..." -ForegroundColor Cyan
            $teamsToken = Request-ResourceToken '48ac35b8-9aa8-4d74-927d-1f4a14a0b239/.default' -RequireScope 'user_impersonation'
            Import-Module MicrosoftTeams -ErrorAction Stop
            Connect-MicrosoftTeams -AccessTokens @($script:GraphToken.AccessToken, $teamsToken) -ErrorAction Stop | Out-Null
            function Invoke-TeamsSafe { param([scriptblock]$Script, [string]$What) try { return (& $Script) } catch { Write-Verbose "Could not read $What : $($_.Exception.Message)"; return $null } }
            $fed  = Invoke-TeamsSafe { Get-CsTenantFederationConfiguration } 'federation configuration'
            $tcc  = Invoke-TeamsSafe { Get-CsTeamsClientConfiguration -Identity Global } 'client configuration'
            $mtg  = Invoke-TeamsSafe { Get-CsTeamsMeetingPolicy -Identity Global } 'global meeting policy'
            if ($fed -or $tcc -or $mtg) {
                $teamsPolicyCollected = $true
                $allowedDomains = if ($fed) { $ad = $fed.AllowedDomains; if ($ad -and $ad.PSObject.Properties['AllowedDomain'] -and @($ad.AllowedDomain).Count) { Join-List (@($ad.AllowedDomain) | ForEach-Object { [string]$_.Domain }) } elseif (([string]$ad) -match 'AllowAllKnownDomains') { 'All external domains' } else { [string]$ad } } else { '' }
                $teamsPolicyKv['External access (federation) with other organizations'] = $(if ($fed) { $(if ($fed.AllowFederatedUsers) { "Enabled; allowed domains: $allowedDomains$(if (@($fed.BlockedDomains).Count) { "; blocked: $(Join-List (@($fed.BlockedDomains) | ForEach-Object { [string]$_.Domain }))" })" } else { 'Disabled' }) } else { 'Not collected' })
                $teamsPolicyKv['Chat with Teams personal (consumer) accounts']      = $(if ($fed) { "$($fed.AllowTeamsConsumer) (inbound: $($fed.AllowTeamsConsumerInbound))" } else { 'Not collected' })
                $teamsPolicyKv['Chat with Skype (consumer) users']                   = $(if ($fed) { [string]$fed.AllowPublicUsers } else { 'Not collected' })
                $teamsPolicyKv['Guest access to Teams']                              = $(if ($tcc) { [string]$tcc.AllowGuestUser } else { 'Not collected' })
                $teamsPolicyKv['Third-party cloud storage in Teams (Dropbox / Box / Google Drive / ShareFile / Egnyte)'] = $(if ($tcc) { "$($tcc.AllowDropBox) / $($tcc.AllowBox) / $($tcc.AllowGoogleDrive) / $($tcc.AllowShareFile) / $($tcc.AllowEgnyte)" } else { 'Not collected' })
                $teamsPolicyKv['Email into channels']                                = $(if ($tcc) { [string]$tcc.AllowEmailIntoChannel } else { 'Not collected' })
                $teamsPolicyKv['Anonymous users can join meetings (global policy)']  = $(if ($mtg) { [string]$mtg.AllowAnonymousUsersToJoinMeeting } else { 'Not collected' })
                $teamsPolicyKv['Anonymous users can start meetings']                 = $(if ($mtg) { [string]$mtg.AllowAnonymousUsersToStartMeeting } else { 'Not collected' })
                $teamsPolicyKv['Who bypasses the lobby']                             = $(if ($mtg) { [string]$mtg.AutoAdmittedUsers } else { 'Not collected' })
                $teamsPolicyKv['Dial-in users bypass the lobby']                     = $(if ($mtg) { [string]$mtg.AllowPSTNUsersToBypassLobby } else { 'Not collected' })
                $teamsPolicyKv['External participants can take control']             = $(if ($mtg) { [string]$mtg.AllowExternalParticipantGiveRequestControl } else { 'Not collected' })
                $teamsPolicyKv['Cloud recording allowed']                            = $(if ($mtg) { [string]$mtg.AllowCloudRecording } else { 'Not collected' })
                if ($fed -and $fed.AllowFederatedUsers -and $allowedDomains -eq 'All external domains') { $teamsPolicyIssues['External access (federation) with other organizations'] = 'teams-federation-open' }
                if ($fed -and $fed.AllowTeamsConsumer) { $teamsPolicyIssues['Chat with Teams personal (consumer) accounts'] = 'teams-federation-open' }
                if ($mtg -and $mtg.AllowAnonymousUsersToJoinMeeting -and ([string]$mtg.AutoAdmittedUsers) -match 'Everyone') { $teamsPolicyIssues['Anonymous users can join meetings (global policy)'] = 'teams-anon-join' }
            }
        } catch {
            $msg = $_.Exception.Message
            try { $body = $_.ErrorDetails.Message; if ($body) { $j = $body | ConvertFrom-Json; if ($j.error_description) { $msg = ([string]$j.error_description -split "`r?`n")[0] } } } catch { Write-Verbose 'Teams error body was not JSON.' }
            if ($msg -match 'AADSTS65001|AADSTS650|consent') { $script:ConsentNeeded['Teams tenant policies'] = '48ac35b8-9aa8-4d74-927d-1f4a14a0b239/user_impersonation' }
            $teamsPolicyNote = if ($msg -match 'AADSTS65001|AADSTS650|consent') { "The Skype and Teams Tenant Admin API permission has not been consented for the connector. A Global Administrator can grant it at https://login.microsoftonline.com/$($ctx.TenantId)/adminconsent?client_id=$($script:GraphClientId)" } else { "Teams tenant policies could not be read: $msg" }
            Write-Warning "Teams policies: $teamsPolicyNote"
        } finally { try { Disconnect-MicrosoftTeams -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch { Write-Verbose 'Teams session was already closed.' } }
    } else { $teamsPolicyNote = 'Teams tenant policies were not collected: device-code sign-in was used (no refresh token for the Teams admin API).' }
} else { $teamsPolicyNote = 'Teams tenant policies (external access, guest access, meeting policy) were not collected: the MicrosoftTeams PowerShell module is not installed. Install-Module MicrosoftTeams -Scope CurrentUser adds them to the next report.' }
if (-not $teamsPolicyCollected) { $teamsPolicyKv['Tenant-wide Teams policies'] = $teamsPolicyNote; $teamsPolicyIssues['Tenant-wide Teams policies'] = 'teams-policies-na' }
$teamsKv = [ordered]@{
    'Teams in the tenant'                                   = $(if (Get-GraphFailure 'teams') { "Not available: $(Get-GraphFailure 'teams')" } else { @($teamGroups).Count })
    'Teams with guests / total guest memberships'           = "$teamsWithGuests / $teamGuestTotal"
    'Public teams (anyone in the organization can join)'    = $teamsPublic
    'Teams with no owner'                                   = $teamsNoOwner
    "Inactive teams (no activity in $StaleDays days)"       = $(if ($teamsActivity.Count) { $teamsInactive } else { "Activity report not available$(if (Get-GraphFailure 'Teams activity report') { ": $(Get-GraphFailure 'Teams activity report')" })" })
    'Archived teams'                                        = $teamsArchived
    'Members may invite guests to Teams and Microsoft 365 Groups (Group.Unified AllowToAddGuests)' = $(if ($null -ne $allowAddGuests) { $allowAddGuests } elseif ($unified) { 'True (default)' } else { 'True (default; no Group.Unified setting object)' })
    'Guests may access group content (AllowGuestsToAccessGroups)' = $(if ($null -ne $allowGuestsAccess) { $allowGuestsAccess } else { 'True (default)' })
}
$teamsKvIssues = @{}
if ($teamsWithGuests) { $teamsKvIssues['Teams with guests / total guest memberships'] = 'teams-guests' }
if ($teamsPublic) { $teamsKvIssues['Public teams (anyone in the organization can join)'] = 'teams-public' }
if ($teamsNoOwner) { $teamsKvIssues['Teams with no owner'] = 'teams-no-owner' }
if ($teamsActivity.Count -and $teamsInactive) { $teamsKvIssues["Inactive teams (no activity in $StaleDays days)"] = 'teams-inactive' }
if (([string]$teamsKv['Members may invite guests to Teams and Microsoft 365 Groups (Group.Unified AllowToAddGuests)']) -match '^True') { $teamsKvIssues['Members may invite guests to Teams and Microsoft 365 Groups (Group.Unified AllowToAddGuests)'] = 'groups-guests-open' }
$teamsStatus = if (Get-GraphFailure 'teams') { 'Not available' } elseif (([string]$teamsKv['Members may invite guests to Teams and Microsoft 365 Groups (Group.Unified AllowToAddGuests)']) -match '^False') { 'Restricted' } elseif ($teamsWithGuests) { "Guests present in $teamsWithGuests team(s)" } else { 'Allowed (no guests today)' }
$teamsEvidence = "$(@($teamGroups).Count) teams; $teamsWithGuests with guests ($teamGuestTotal guest memberships); $teamsPublic public; $teamsNoOwner without an owner. Guest invitations by members: $($teamsKv['Members may invite guests to Teams and Microsoft 365 Groups (Group.Unified AllowToAddGuests)']). $(if ($teamsPolicyCollected) { "External access: $($teamsPolicyKv['External access (federation) with other organizations']); guest access: $($teamsPolicyKv['Guest access to Teams']); anonymous meeting join: $($teamsPolicyKv['Anonymous users can join meetings (global policy)'])." } else { 'Tenant-wide Teams policies not collected.' })"

# ----------------------------------------------------------------------------------------------
# SharePoint and OneDrive (section 18): tenant sharing settings, every site, every OneDrive, who shares externally
# ----------------------------------------------------------------------------------------------
Write-Host "Collecting SharePoint and OneDrive..." -ForegroundColor Cyan
$spoSettings = (Get-GraphAllSafe 'v1.0/admin/sharepoint/settings' 'SharePoint tenant settings' -NoPaging) | Select-Object -First 1
$spoKv = [ordered]@{}; $spoKvIssues = @{}
$sharingLabel = @{ 'disabled' = 'Only people in the organization (external sharing off)'; 'existingExternalUserSharingOnly' = 'Existing guests only'; 'externalUserSharingOnly' = 'New and existing guests (sign-in required)'; 'externalUserAndGuestSharing' = 'Anyone (anonymous links allowed)' }
if ($spoSettings) {
    $cap = [string]$spoSettings.sharingCapability
    $spoKv['External sharing (SharePoint)']              = $(if ($sharingLabel.ContainsKey($cap)) { $sharingLabel[$cap] } else { $cap })
    $spoKv['Sharing domain restriction']                 = $(switch ([string]$spoSettings.sharingDomainRestrictionMode) { 'allowList' { "Allow list: $(Join-List $spoSettings.sharingAllowedDomainList)" } 'blockList' { "Block list: $(Join-List $spoSettings.sharingBlockedDomainList)" } default { 'None (any domain)' } })
    $spoKv['Guests can re-share items they do not own']  = $spoSettings.isResharingByExternalUsersEnabled
    $spoKv['Guests must sign in with the invited account'] = $spoSettings.isRequireAcceptingUserToMatchInvitedUserEnabled
    $spoKv['Legacy authentication protocols allowed']    = $spoSettings.isLegacyAuthProtocolsEnabled
    $spoKv['OneDrive sync restricted to managed devices'] = $spoSettings.isUnmanagedSyncAppForTenantRestricted
    $spoKv['Mac sync client allowed']                    = $spoSettings.isMacSyncAppEnabled
    $spoKv['Idle session sign-out']                      = $(if ($spoSettings.idleSessionSignOut -and $spoSettings.idleSessionSignOut.isEnabled) { "Enabled: warn after $([int]$spoSettings.idleSessionSignOut.warnAfterInSeconds / 60) min, sign out after $([int]$spoSettings.idleSessionSignOut.signOutAfterInSeconds / 60) min" } else { 'Disabled' })
    $spoKv['Sync excluded file types']                   = (Join-List $spoSettings.excludedFileExtensionsForSyncApp)
    $spoKv['Default OneDrive storage per user (GB)']     = $(if ($spoSettings.personalSiteDefaultStorageLimitInMB) { [math]::Round([double]$spoSettings.personalSiteDefaultStorageLimitInMB / 1024) } else { '' })
    $spoKv['OneDrive retention after a user is deleted (days)'] = $spoSettings.deletedUserPersonalSiteRetentionPeriodInDays
    $spoKv['Users can create sites']                     = $spoSettings.isSiteCreationEnabled
    $spoKv['Loop components enabled']                    = $spoSettings.isLoopEnabled
    if ($cap -eq 'externalUserAndGuestSharing') { $spoKvIssues['External sharing (SharePoint)'] = 'spo-anyone-links' } elseif ($cap -ne 'disabled') { $spoKvIssues['External sharing (SharePoint)'] = 'spo-external-sharing' }
    if ($spoSettings.isResharingByExternalUsersEnabled -eq $true -and $cap -ne 'disabled') { $spoKvIssues['Guests can re-share items they do not own'] = 'spo-resharing' }
    if ($spoSettings.isLegacyAuthProtocolsEnabled -eq $true) { $spoKvIssues['Legacy authentication protocols allowed'] = 'spo-legacy-auth' }
    if ($spoSettings.isUnmanagedSyncAppForTenantRestricted -eq $false) { $spoKvIssues['OneDrive sync restricted to managed devices'] = 'spo-unmanaged-sync' }
    if (-not ($spoSettings.idleSessionSignOut -and $spoSettings.idleSessionSignOut.isEnabled)) { $spoKvIssues['Idle session sign-out'] = 'spo-idle-timeout-off' }
} else { $spoKv['SharePoint tenant settings'] = "Not available: $(Get-GraphFailure 'SharePoint tenant settings')" }

$siteReport = @(Get-GraphReportCsv "getSharePointSiteUsageDetail(period='D90')" 'SharePoint site usage report')
$siteConcealed = Test-ConcealedReportNames $siteReport 'Site URL'
function ConvertTo-Gb { param($bytes) if ($null -eq $bytes -or $bytes -eq '') { return '' } ; return [math]::Round([double]$bytes / 1GB, 2) }
function Get-SiteTemplateLabel { param([string]$t) switch -Regex ($t) { '^GROUP' { 'Team site (Microsoft 365 group)' } '^SITEPAGEPUBLISHING' { 'Communication site' } '^STS' { 'Team site (no group)' } '^TEAMCHANNEL' { 'Teams channel site' } '^SPSPERS|^SPSMSITE' { 'OneDrive' } '^APPCATALOG' { 'App catalog' } '^SRCHCEN|^SRCHCENTERLITE' { 'Search center' } '^POINTPUBLISHING|^PWA|^EHS|^RedirectSite' { 'System' } default { $(if ($t) { $t } else { '' }) } } }
$siteRows = @(); $sitesInactive = 0; $sitesNoOwner = 0; $sitesDeleted = 0
foreach ($r in ($siteReport | Sort-Object { [double]$(if ($_.'Storage Used (Byte)') { $_.'Storage Used (Byte)' } else { 0 }) } -Descending)) {
    $last = if ($r.'Last Activity Date') { try { [datetime]$r.'Last Activity Date' } catch { $null } } else { $null }
    $inactive = ($null -eq $last) -or ($last -lt (Get-Date).AddDays(-$StaleDays))
    $deleted = ([string]$r.'Is Deleted') -eq 'True'
    $is = @{}; $tags = @()
    if ($deleted) { $tags += 'deleted'; $sitesDeleted++ }
    elseif ($inactive) { $is['Last activity'] = 'spo-site-inactive'; $tags += 'inactive'; $sitesInactive++ }
    if (-not $r.'Owner Principal Name' -and -not $r.'Owner Display Name') { $is['Owner'] = 'spo-site-no-owner'; $tags += 'noowner'; $sitesNoOwner++ }
    $tpl = Get-SiteTemplateLabel ([string]$r.'Root Web Template')
    if ($tpl -like 'Team site*') { $tags += 'team' } elseif ($tpl -eq 'Communication site') { $tags += 'comms' }
    $siteRows += [ordered]@{ 'Site' = $r.'Site URL'; 'Type' = $tpl; 'Owner' = $(if ($r.'Owner Display Name') { "$($r.'Owner Display Name') ($($r.'Owner Principal Name'))" } else { [string]$r.'Owner Principal Name' }); 'Last activity' = $(if ($last) { $last.ToString('yyyy-MM-dd') } else { 'None in 90 days' }); 'Files' = $r.'File Count'; 'Active files (90 days)' = $r.'Active File Count'; 'Storage used (GB)' = (ConvertTo-Gb $r.'Storage Used (Byte)'); 'Page views (90 days)' = $r.'Page View Count'; 'Deleted' = $(if ($deleted) { 'Yes' } else { '' }); '_issues' = $is; '_tags' = ($tags -join ' ') }
}
$siteRows = @($siteRows)
$SiteFilters = @(@{ Label = "Inactive ($StaleDays days)"; Tag = 'inactive' }, @{ Label = 'No owner'; Tag = 'noowner' }, @{ Label = 'Team sites'; Tag = 'team' }, @{ Label = 'Communication sites'; Tag = 'comms' }, @{ Label = 'Deleted'; Tag = 'deleted' })

$odReport = @(Get-GraphReportCsv "getOneDriveUsageAccountDetail(period='D90')" 'OneDrive usage report')
$odConcealed = Test-ConcealedReportNames $odReport 'Owner Principal Name'
$odRows = @(); $odInactive = 0; $odOwnerDisabled = 0; $odOwnerMissing = 0
foreach ($r in ($odReport | Sort-Object { [double]$(if ($_.'Storage Used (Byte)') { $_.'Storage Used (Byte)' } else { 0 }) } -Descending)) {
    $last = if ($r.'Last Activity Date') { try { [datetime]$r.'Last Activity Date' } catch { $null } } else { $null }
    $owner = $UserByUpn[([string]$r.'Owner Principal Name').ToLower()]
    $is = @{}; $tags = @()
    if (([string]$r.'Is Deleted') -eq 'True') { $tags += 'deleted' }
    if (-not $odConcealed -and $r.'Owner Principal Name' -and -not $owner) { $is['Owner'] = 'od-owner-disabled'; $tags += 'ownermissing'; $odOwnerMissing++ }
    elseif ($owner -and -not $owner.accountEnabled) { $is['Owner'] = 'od-owner-disabled'; $tags += 'ownerdisabled'; $odOwnerDisabled++ }
    if (($null -eq $last) -or ($last -lt (Get-Date).AddDays(-$StaleDays))) { $tags += 'inactive'; $odInactive++ }
    $odRows += [ordered]@{ 'Owner' = $(if ($r.'Owner Display Name') { "$($r.'Owner Display Name') ($($r.'Owner Principal Name'))" } else { [string]$r.'Owner Principal Name' }); 'Account' = $(if ($odConcealed) { 'n/a (names concealed)' } elseif (-not $owner) { 'Not in directory' } elseif ($owner.accountEnabled) { 'Enabled' } else { 'Disabled' }); 'Last activity' = $(if ($last) { $last.ToString('yyyy-MM-dd') } else { 'None in 90 days' }); 'Files' = $r.'File Count'; 'Active files (90 days)' = $r.'Active File Count'; 'Storage used (GB)' = (ConvertTo-Gb $r.'Storage Used (Byte)'); 'Deleted' = $(if (([string]$r.'Is Deleted') -eq 'True') { 'Yes' } else { '' }); '_issues' = $is; '_tags' = ($tags -join ' ') }
}
$odRows = @($odRows)
$OdFilters = @(@{ Label = 'Owner disabled'; Tag = 'ownerdisabled' }, @{ Label = 'Owner not in directory'; Tag = 'ownermissing' }, @{ Label = "Inactive ($StaleDays days)"; Tag = 'inactive' }, @{ Label = 'Deleted'; Tag = 'deleted' })

$spoActivity = @(Get-GraphReportCsv "getSharePointActivityUserDetail(period='D30')" 'SharePoint user activity report')
$odActivity  = @(Get-GraphReportCsv "getOneDriveActivityUserDetail(period='D30')" 'OneDrive user activity report')
$extShareBy = @{}
foreach ($r in $spoActivity) { $n = [int]$(if ($r.'Shared Externally File Count') { $r.'Shared Externally File Count' } else { 0 }); if ($n -gt 0) { $u = [string]$r.'User Principal Name'; if (-not $extShareBy.ContainsKey($u)) { $extShareBy[$u] = @{ Spo = 0; Od = 0; Internal = 0 } }; $extShareBy[$u].Spo += $n; $extShareBy[$u].Internal += [int]$(if ($r.'Shared Internally File Count') { $r.'Shared Internally File Count' } else { 0 }) } }
foreach ($r in $odActivity)  { $n = [int]$(if ($r.'Shared Externally File Count') { $r.'Shared Externally File Count' } else { 0 }); if ($n -gt 0) { $u = [string]$r.'User Principal Name'; if (-not $extShareBy.ContainsKey($u)) { $extShareBy[$u] = @{ Spo = 0; Od = 0; Internal = 0 } }; $extShareBy[$u].Od += $n } }
$extShareRows = @($extShareBy.GetEnumerator() | Sort-Object { $_.Value.Spo + $_.Value.Od } -Descending | ForEach-Object { $uo = $UserByUpn[$_.Key.ToLower()]; [ordered]@{ 'User' = $(if ($uo) { $uo.displayName } else { $_.Key }); 'UPN' = $_.Key; 'Files shared externally from SharePoint (30 days)' = $_.Value.Spo; 'Files shared externally from OneDrive (30 days)' = $_.Value.Od; 'Files shared internally (30 days)' = $_.Value.Internal; '_issues' = @{ 'UPN' = 'spo-external-shares' } } })
$extShareTotal = 0; foreach ($v in $extShareBy.Values) { $extShareTotal += $v.Spo + $v.Od }
$spoKv['Sites (excluding OneDrive)']                 = $(if ($siteReport.Count) { $siteRows.Count } else { "Usage report not available$(if (Get-GraphFailure 'SharePoint site usage report') { ": $(Get-GraphFailure 'SharePoint site usage report')" })" })
$spoKv["Inactive sites (no activity in $StaleDays days)"] = $(if ($siteReport.Count) { $sitesInactive } else { '' })
$spoKv['Sites with no owner']                        = $(if ($siteReport.Count) { $sitesNoOwner } else { '' })
$spoKv['OneDrive accounts']                          = $(if ($odReport.Count) { $odRows.Count } else { "Usage report not available$(if (Get-GraphFailure 'OneDrive usage report') { ": $(Get-GraphFailure 'OneDrive usage report')" })" })
$spoKv['OneDrives whose owner is disabled or gone']  = $(if ($odReport.Count) { $odOwnerDisabled + $odOwnerMissing } else { '' })
$spoKv['Files shared externally in the last 30 days (users / files)'] = $(if ($spoActivity.Count -or $odActivity.Count) { "$($extShareRows.Count) / $extShareTotal" } else { 'Activity report not available' })
if ($siteReport.Count -and $sitesInactive) { $spoKvIssues["Inactive sites (no activity in $StaleDays days)"] = 'spo-site-inactive' }
if ($siteReport.Count -and $sitesNoOwner) { $spoKvIssues['Sites with no owner'] = 'spo-site-no-owner' }
if ($odReport.Count -and ($odOwnerDisabled + $odOwnerMissing)) { $spoKvIssues['OneDrives whose owner is disabled or gone'] = 'od-owner-disabled' }
if ($extShareRows.Count) { $spoKvIssues['Files shared externally in the last 30 days (users / files)'] = 'spo-external-shares' }
$spoStatus = if (-not $spoSettings) { 'Not available' } else { switch ([string]$spoSettings.sharingCapability) { 'disabled' { 'Restricted (internal only)' } 'existingExternalUserSharingOnly' { 'Restricted (existing guests only)' } 'externalUserSharingOnly' { 'Allowed (guests must sign in)' } 'externalUserAndGuestSharing' { 'Anyone links allowed' } default { [string]$spoSettings.sharingCapability } } }
$spoEvidence = if ($spoSettings) { "External sharing: $($spoKv['External sharing (SharePoint)']); domain restriction: $($spoKv['Sharing domain restriction']); guests re-share: $($spoKv['Guests can re-share items they do not own']); legacy auth: $($spoKv['Legacy authentication protocols allowed']); unmanaged sync restricted: $($spoKv['OneDrive sync restricted to managed devices']); idle sign-out: $($spoKv['Idle session sign-out']). $(if ($siteReport.Count) { "$($siteRows.Count) sites, $sitesInactive inactive, $sitesNoOwner without owner. " })$(if ($odReport.Count) { "$($odRows.Count) OneDrives, $($odOwnerDisabled + $odOwnerMissing) with disabled or missing owner. " })$(if ($extShareRows.Count) { "$($extShareRows.Count) user(s) shared $extShareTotal file(s) externally in 30 days." })" } else { [string]$spoKv['SharePoint tenant settings'] }
$reportConcealedNote = if ($teamsConcealed -or $siteConcealed -or $odConcealed) { "Microsoft 365 usage reports in this tenant conceal user, group and site names (they appear as hashes). To show real names, in the Microsoft 365 admin center go to Settings > Org settings > Services > Reports and clear 'Display concealed user, group, and site names in all reports'." } else { '' }
if ($reportConcealedNote) { $spoKv['Usage report names'] = 'Concealed by the tenant report privacy setting (OneDrive, site and team owners show as hashes)'; $spoKvIssues['Usage report names'] = 'reports-concealed' }

# ----------------------------------------------------------------------------------------------
# Sharing and Access (section 19): who can reach what, across SharePoint sites, OneDrive, Teams and Exchange.
# Looks for data open to the internet (Anyone links), to the whole organization, to outside people, and to
# insiders outside the owning team or department. Read-only: drive item metadata and permissions only.
# ----------------------------------------------------------------------------------------------
$shareRows = @(); $crossDeptRows = @(); $guestAccessRows = @(); $shareScanNote = ''; $shareSitesScanned = 0; $shareSitesDenied = 0; $shareItemsScanned = 0; $shareCapHit = $false
$shareAnyone = 0; $shareOrg = 0; $shareExternal = 0; $shareOutsideTeam = 0; $shareInternal = 0
$VerifiedDomainNames = @($org.verifiedDomains | ForEach-Object { ([string]$_.name).ToLower() })
function Test-ExternalIdentity { param([string]$Address) if (-not $Address -or $Address -notmatch '@') { return $false } ; $d = ($Address -split '@')[-1].ToLower(); if ($Address -match '#EXT#') { return $true } ; return -not ($VerifiedDomainNames -contains $d) }
function Get-DeptOf { param([string]$Upn) if (-not $Upn) { return '' } ; $u = $UserByUpn[$Upn.ToLower()]; if ($u -and $u.department) { return [string]$u.department } ; return '' }
function Add-ShareRow {
    param([string]$Source, [string]$Container, [string]$Item, [string]$Url, [string]$AccessType, [string]$Who, [string]$Rights, [string]$Scope, [string]$OutsideTeam, [string]$GrantedBy, [string]$When, [string]$Level, [string]$Issue, [string]$OwnerDept, [string]$WhoDept)
    $tags = @($Source.ToLower().Replace(' ', ''), $Scope.ToLower().Replace(' ', ''), $Level.ToLower())
    if ($OutsideTeam -eq 'Yes') { $tags += 'outsideteam' }
    if ($OwnerDept -and $WhoDept -and $OwnerDept -ne $WhoDept) { $tags += 'crossdept' }
    $script:shareRows += [ordered]@{ 'Source' = $Source; 'Site / team / mailbox' = $Container; 'Item' = $Item; 'Access type' = $AccessType; 'Who has access' = $Who; 'Rights' = $Rights; 'Scope' = $Scope; 'Outside owning team' = $OutsideTeam; 'Department (owner / who)' = $(if ($OwnerDept -or $WhoDept) { "$OwnerDept / $WhoDept" } else { '' }); 'Granted by' = $GrantedBy; 'When' = $When; 'Level' = $Level; '_tags' = ($tags -join ' '); '_issues' = $(if ($Issue) { @{ 'Access type' = $Issue } } else { $null }); '_links' = $(if ($Url) { @{ 'Item' = $Url } } else { $null }) }
}
$script:shareRows = @()
if ($SkipShareScan) { $shareScanNote = 'Share scan skipped (-SkipShareScan).' }
else {
    Write-Host "Scanning sharing and access across sites, OneDrive, Teams and guests..." -ForegroundColor Cyan
    $MaxSites = 80; $MaxItemsPerDrive = 1500; $MaxPermissionLookups = 600; $permLookups = 0
    $allSites = @(Get-GraphAllSafe 'v1.0/sites?search=*&$select=id,displayName,webUrl,createdDateTime' 'SharePoint sites' -MaxItems 400 | Where-Object { $_.webUrl -notmatch '/portals/|/sites/appcatalog|-my\.sharepoint\.com/personal/|/search$|contentTypeHub' })
    if (Get-GraphFailure 'SharePoint sites') { $shareScanNote = "Sites could not be listed: $(Get-GraphFailure 'SharePoint sites')" }
    $siteIdx = 0
    foreach ($site in ($allSites | Sort-Object displayName)) {
        $siteIdx++
        if ($siteIdx -gt $MaxSites) { $shareCapHit = $true; break }
        Write-Progress -Activity 'Scanning sharing' -Status $site.displayName -PercentComplete ([int](100 * $siteIdx / [Math]::Max(1, [Math]::Min($MaxSites, @($allSites).Count))))
        $drives = @()
        try { Confirm-GraphToken; $drives = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/sites/$($site.id)/drives?`$select=id,name,driveType,webUrl" -OutputType PSObject -ErrorAction Stop).value) }
        catch { $shareSitesDenied++; continue }
        $shareSitesScanned++
        $teamName = [string]$site.displayName
        $teamMembers = if ($TeamMembersByName.ContainsKey($teamName)) { $TeamMembersByName[$teamName] } else { $null }
        $isOneDrive = $site.webUrl -match '-my\.sharepoint\.com'
        $sourceLabel = if ($isOneDrive) { 'OneDrive' } elseif ($teamMembers) { 'Teams' } else { 'SharePoint' }
        foreach ($drv in ($drives | Where-Object { $_.driveType -in @('documentLibrary', 'business') })) {
            $items = @()
            try {
                $next = "v1.0/drives/$($drv.id)/root/search(q='')?`$select=id,name,webUrl,shared,createdBy,lastModifiedDateTime,size,folder,file,parentReference&`$top=200"
                while ($next -and $items.Count -lt $MaxItemsPerDrive) {
                    Confirm-GraphToken
                    $pg = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject -ErrorAction Stop
                    $items += @($pg.value); $next = $pg.'@odata.nextLink'
                }
                if ($next) { $shareCapHit = $true }
            } catch { Write-Verbose "Drive $($drv.name) in $($site.displayName) not readable: $($_.Exception.Message)"; continue }
            $shareItemsScanned += $items.Count
            foreach ($it in ($items | Where-Object { $_.shared })) {
                $itemPath = $(if ($it.parentReference -and $it.parentReference.path) { (([string]$it.parentReference.path) -replace '^.*root:', '') + '/' + $it.name } else { [string]$it.name })
                $itemLabel = "$(if ($it.folder) { '[folder] ' })$itemPath"
                $ownerUpn = [string]$it.createdBy.user.email; if (-not $ownerUpn) { $ownerUpn = [string]$it.shared.owner.user.email }
                $ownerDept = Get-DeptOf $ownerUpn
                if ($permLookups -ge $MaxPermissionLookups) {
                    # cap reached: record the item with the scope the drive reports, without per-grantee detail
                    $scopeWord = switch ([string]$it.shared.scope) { 'anonymous' { 'Anyone' } 'organization' { 'Whole organization' } default { 'Specific people' } }
                    Add-ShareRow -Source $sourceLabel -Container $site.displayName -Item $itemLabel -Url $it.webUrl -AccessType "Shared ($scopeWord)" -Who '(detail not fetched: permission lookup cap reached)' -Rights '' -Scope $scopeWord -OutsideTeam 'n/a' -GrantedBy ([string]$it.shared.sharedBy.user.displayName) -When (Format-DateText $it.shared.sharedDateTime) -Level $(if ($scopeWord -eq 'Anyone') { 'Urgent' } else { 'Review' }) -Issue $(if ($scopeWord -eq 'Anyone') { 'share-anonymous' } elseif ($scopeWord -eq 'Whole organization') { 'share-org' } else { $null }) -OwnerDept $ownerDept -WhoDept ''
                    $shareCapHit = $true; continue
                }
                $perms = @()
                try { Confirm-GraphToken; $permLookups++; $perms = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/drives/$($drv.id)/items/$($it.id)/permissions" -OutputType PSObject -ErrorAction Stop).value) } catch { Write-Verbose "Permissions not readable for $itemPath"; continue }
                foreach ($pm in ($perms | Where-Object { -not $_.inheritedFrom })) {
                    $rights = Join-List (@($pm.roles) | ForEach-Object { [string]$_ }) ', '
                    $granted = @($(if ($pm.grantedToIdentitiesV2) { @($pm.grantedToIdentitiesV2) } elseif ($pm.grantedToV2) { @($pm.grantedToV2) } elseif ($pm.grantedTo) { @($pm.grantedTo) } else { @() }))
                    $whoList = @(); $extAny = $false; $outside = 'n/a'; $whoDept = ''
                    foreach ($g in $granted) {
                        $gm = [string]$(if ($g.user) { $(if ($g.user.email) { $g.user.email } else { $g.user.displayName }) } elseif ($g.siteUser) { $(if ($g.siteUser.email) { $g.siteUser.email } else { $g.siteUser.displayName }) } elseif ($g.group) { "$($g.group.displayName) (group)" } elseif ($g.siteGroup) { "$($g.siteGroup.displayName) (site group)" } elseif ($g.application) { "$($g.application.displayName) (application)" } else { '' })
                        if ($gm) { $whoList += $gm }
                        if ($gm -and (Test-ExternalIdentity $gm)) { $extAny = $true }
                        if ($gm -match '@' -and -not $whoDept) { $whoDept = Get-DeptOf $gm }
                        if ($teamMembers -and $gm -match '@' -and -not (Test-ExternalIdentity $gm)) { $outside = $(if ($teamMembers -contains $gm.ToLower()) { 'No' } else { 'Yes' }) }
                    }
                    if ($pm.link) {
                        $lscope = [string]$pm.link.scope
                        $ltype = [string]$pm.link.type
                        switch ($lscope) {
                            'anonymous'    { $shareAnyone++;   Add-ShareRow -Source $sourceLabel -Container $site.displayName -Item $itemLabel -Url $it.webUrl -AccessType "Anyone link ($ltype)" -Who 'Anyone on the internet with the link' -Rights $rights -Scope 'Anyone' -OutsideTeam 'n/a' -GrantedBy ([string]$pm.grantedBy.user.displayName) -When (Format-DateText $pm.expirationDateTime) -Level 'Urgent' -Issue 'share-anonymous' -OwnerDept $ownerDept -WhoDept '' }
                            'organization' { $shareOrg++;      Add-ShareRow -Source $sourceLabel -Container $site.displayName -Item $itemLabel -Url $it.webUrl -AccessType "Whole-organization link ($ltype)" -Who 'Anyone signed in to the organization (including guests)' -Rights $rights -Scope 'Whole organization' -OutsideTeam $(if ($teamMembers) { 'Yes' } else { 'n/a' }) -GrantedBy ([string]$pm.grantedBy.user.displayName) -When '' -Level 'Review' -Issue 'share-org' -OwnerDept $ownerDept -WhoDept '' }
                            default        { $who = $(if ($whoList.Count) { Join-List $whoList '; ' } else { 'Specific people (list not returned)' }); if ($extAny) { $shareExternal++ } elseif ($outside -eq 'Yes') { $shareOutsideTeam++ } else { $shareInternal++ }
                                             Add-ShareRow -Source $sourceLabel -Container $site.displayName -Item $itemLabel -Url $it.webUrl -AccessType "Specific-people link ($ltype)" -Who $who -Rights $rights -Scope $(if ($extAny) { 'External' } else { 'Specific people' }) -OutsideTeam $outside -GrantedBy ([string]$pm.grantedBy.user.displayName) -When (Format-DateText $pm.expirationDateTime) -Level $(if ($extAny) { 'Review' } elseif ($outside -eq 'Yes') { 'Review' } else { 'Info' }) -Issue $(if ($extAny) { 'share-external' } elseif ($outside -eq 'Yes') { 'share-outside-team' } else { $null }) -OwnerDept $ownerDept -WhoDept $whoDept }
                        }
                    } elseif ($granted.Count) {
                        $who = Join-List $whoList '; '
                        if ($extAny) { $shareExternal++ } elseif ($outside -eq 'Yes') { $shareOutsideTeam++ } else { $shareInternal++ }
                        Add-ShareRow -Source $sourceLabel -Container $site.displayName -Item $itemLabel -Url $it.webUrl -AccessType 'Direct permission' -Who $who -Rights $rights -Scope $(if ($extAny) { 'External' } else { 'Specific people' }) -OutsideTeam $outside -GrantedBy ([string]$pm.grantedBy.user.displayName) -When '' -Level $(if ($extAny) { 'Review' } elseif ($outside -eq 'Yes') { 'Review' } else { 'Info' }) -Issue $(if ($extAny) { 'share-external' } elseif ($outside -eq 'Yes') { 'share-outside-team' } else { $null }) -OwnerDept $ownerDept -WhoDept $whoDept
                    }
                }
            }
        }
    }
    Write-Progress -Activity 'Scanning sharing' -Completed
    # OneDrive of each enabled user: readable only where the signed-in account has rights (usually its own); counted, not warned
    $odAccessible = 0; $odDenied = 0
    foreach ($u in @($enabledMembers | Select-Object -First 150)) {
        try {
            Confirm-GraphToken
            $odItems = @((Invoke-MgGraphRequest -Method GET -Uri "v1.0/users/$($u.id)/drive/root/search(q='')?`$select=id,name,webUrl,shared,createdBy,parentReference,folder&`$top=200" -OutputType PSObject -ErrorAction Stop).value)
            $odAccessible++
            foreach ($it in ($odItems | Where-Object { $_.shared })) {
                $scopeWord = switch ([string]$it.shared.scope) { 'anonymous' { 'Anyone' } 'organization' { 'Whole organization' } default { 'Specific people' } }
                if ($scopeWord -eq 'Anyone') { $shareAnyone++ } elseif ($scopeWord -eq 'Whole organization') { $shareOrg++ } else { $shareInternal++ }
                Add-ShareRow -Source 'OneDrive' -Container "OneDrive of $($u.displayName)" -Item "$(if ($it.folder) { '[folder] ' })$($it.name)" -Url $it.webUrl -AccessType "Shared ($scopeWord)" -Who $(switch ($scopeWord) { 'Anyone' { 'Anyone on the internet with the link' } 'Whole organization' { 'Anyone signed in to the organization' } default { 'Specific people (see item)' } }) -Rights '' -Scope $scopeWord -OutsideTeam 'n/a' -GrantedBy ([string]$it.shared.sharedBy.user.displayName) -When (Format-DateText $it.shared.sharedDateTime) -Level $(if ($scopeWord -eq 'Anyone') { 'Urgent' } elseif ($scopeWord -eq 'Whole organization') { 'Review' } else { 'Info' }) -Issue $(if ($scopeWord -eq 'Anyone') { 'share-anonymous' } elseif ($scopeWord -eq 'Whole organization') { 'share-org' } else { $null }) -OwnerDept (Get-DeptOf $u.userPrincipalName) -WhoDept ''
            }
        } catch { $odDenied++ }
    }
    if (-not $shareScanNote) { $shareScanNote = "Scanned $shareItemsScanned item(s) in $shareSitesScanned site(s)$(if ($shareSitesDenied) { "; $shareSitesDenied site(s) were not readable by $($ctx.Account)" }); OneDrive: $odAccessible readable, $odDenied not readable by this account$(if ($shareCapHit) { '; caps reached (80 sites, 1,500 items per library, 600 permission lookups), so the list is a sample' })." }
}
# Teams: guests, public teams and external users / shared channels from the activity report
foreach ($tr in $teamRows) {
    if ([int]$tr['Guests'] -gt 0) { $shareExternal += [int]$tr['Guests']; Add-ShareRow -Source 'Teams' -Container $tr['Team'] -Item 'Whole team (all channels and files)' -Url '' -AccessType 'Guest membership' -Who $(if ($tr['Guest names']) { $tr['Guest names'] } else { "$($tr['Guests']) guest(s)" }) -Rights 'member' -Scope 'External' -OutsideTeam 'n/a' -GrantedBy (Join-List @($tr['Owners'])) -When '' -Level 'Review' -Issue 'teams-guests' -OwnerDept '' -WhoDept '' }
    if ($tr['Visibility'] -eq 'Public') { $shareOrg++; Add-ShareRow -Source 'Teams' -Container $tr['Team'] -Item 'Whole team (all channels and files)' -Url '' -AccessType 'Public team' -Who 'Anyone in the organization can join and read' -Rights 'member' -Scope 'Whole organization' -OutsideTeam 'n/a' -GrantedBy (Join-List @($tr['Owners'])) -When $tr['Created'] -Level 'Review' -Issue 'teams-public' -OwnerDept '' -WhoDept '' }
}
foreach ($tid in $teamsActivity.Keys) {
    $ta = $teamsActivity[$tid]
    $extUsers = [int]$(if ($ta.'Active External Users') { $ta.'Active External Users' } else { 0 }); $sharedCh = [int]$(if ($ta.'Active Shared Channels') { $ta.'Active Shared Channels' } else { 0 })
    if ($extUsers -gt 0 -or $sharedCh -gt 0) { $shareExternal += $extUsers; Add-ShareRow -Source 'Teams' -Container ([string]$ta.'Team Name') -Item "Shared channels ($sharedCh active)" -Url '' -AccessType 'External users via shared channels (90 days)' -Who "$extUsers active external user(s) from other organizations" -Rights 'channel member' -Scope 'External' -OutsideTeam 'n/a' -GrantedBy '' -When '' -Level 'Review' -Issue 'share-external' -OwnerDept '' -WhoDept '' }
}
# Exchange: shared mailboxes and their delegates, Send As, calendar folders open to the organization or anonymous users
foreach ($fa in $fullAccessRows) { $ext = Test-ExternalIdentity ([string]$fa['Delegate (Full Access)']); $shareInternal++; Add-ShareRow -Source 'Exchange' -Container $fa['Mailbox'] -Item "Mailbox ($($fa['Mailbox type']))" -Url '' -AccessType 'Full Access delegate' -Who $fa['Delegate (Full Access)'] -Rights $fa['Rights'] -Scope $(if ($ext) { 'External' } else { 'Specific people' }) -OutsideTeam 'n/a' -GrantedBy '' -When '' -Level $(if ($fa['Mailbox type'] -eq 'SharedMailbox') { 'Info' } else { 'Review' }) -Issue $(if ($fa['Mailbox type'] -ne 'SharedMailbox') { 'ioc-delegate' } else { $null }) -OwnerDept (Get-DeptOf $fa['Mailbox']) -WhoDept (Get-DeptOf $fa['Delegate (Full Access)']) }
foreach ($sa in $sendAsRows) { $shareInternal++; Add-ShareRow -Source 'Exchange' -Container $sa['Mailbox'] -Item 'Mailbox identity' -Url '' -AccessType 'Send As' -Who $sa['Trustee (can send as)'] -Rights 'SendAs' -Scope 'Specific people' -OutsideTeam 'n/a' -GrantedBy '' -When '' -Level 'Review' -Issue 'ioc-sendas' -OwnerDept '' -WhoDept '' }
foreach ($cr in $calendarShareRows) { if ($cr['Level'] -eq 'Anonymous') { $shareAnyone++ } else { $shareOrg++ }; Add-ShareRow -Source 'Exchange' -Container $cr['Mailbox'] -Item 'Calendar' -Url '' -AccessType "Calendar folder permission ($($cr['Level']))" -Who $(if ($cr['Level'] -eq 'Anonymous') { 'Anyone (anonymous)' } else { 'Everyone in the organization (Default)' }) -Rights $cr['Access rights'] -Scope $(if ($cr['Level'] -eq 'Anonymous') { 'Anyone' } else { 'Whole organization' }) -OutsideTeam 'n/a' -GrantedBy '' -When '' -Level $(if ($cr['Level'] -eq 'Anonymous') { 'Urgent' } elseif ($cr['Access rights'] -match 'Reviewer|Editor|Owner|Author|PublishingEditor') { 'Review' } else { 'Info' }) -Issue $(if ($cr['Level'] -eq 'Anonymous') { 'calendar-anonymous' } elseif ($cr['Access rights'] -match 'Reviewer|Editor|Owner|Author|PublishingEditor') { 'calendar-default-open' } else { $null }) -OwnerDept (Get-DeptOf $cr['Mailbox']) -WhoDept '' }
$shareRows = @($script:shareRows)
# Guests: who they are and what they are members of
$guestUsers = @($users | Where-Object { $_.userType -eq 'Guest' } | Sort-Object displayName | Select-Object -First 300)
foreach ($g in $guestUsers) {
    $memberOf = @(Get-GraphAllSafe "v1.0/users/$($g.id)/memberOf?`$select=id,displayName" "groups of guest $($g.displayName)" -MaxItems 200 | ForEach-Object { [string]$_.displayName } | Where-Object { $_ })
    $gdomain = $(if ($g.mail) { ([string]$g.mail -split '@')[-1] } elseif ($g.userPrincipalName -match '#EXT#') { (([string]$g.userPrincipalName -split '#EXT#')[0] -split '_')[-1] } else { '' })
    $guestAccessRows += [ordered]@{ 'Guest' = $g.displayName; 'Email domain' = $gdomain; 'Enabled' = $g.accountEnabled; 'Invited' = (Format-DateText $g.createdDateTime); 'Last sign-in' = $(if ($g.signInActivity) { Format-DateText $g.signInActivity.lastSignInDateTime } else { '' }); 'Member of (teams and groups)' = $(if ($memberOf.Count) { Join-List ($memberOf | Select-Object -First 8) '; ' } else { 'No group memberships (direct shares only)' }); 'Groups' = $memberOf.Count; '_issues' = @{ 'Guest' = 'guest' }; '_tags' = $(if ($memberOf.Count) { 'ingroups' } else { 'nogroups' }) + $(if ($g.accountEnabled -eq $false) { ' disabled' } else { '' }) }
}
$guestAccessRows = @($guestAccessRows)
# Cross-department view: who in one department can reach another department's data (from the share rows that carry both departments)
$crossPairs = @{}
foreach ($sr in ($shareRows | Where-Object { $_['_tags'] -match '\bcrossdept\b' })) { $k = $sr['Department (owner / who)']; if (-not $crossPairs.ContainsKey($k)) { $crossPairs[$k] = @{ Count = 0; Items = @() } }; $crossPairs[$k].Count++; if ($crossPairs[$k].Items.Count -lt 3) { $crossPairs[$k].Items += "$($sr['Item']) -> $($sr['Who has access'])" } }
$crossDeptRows = @($crossPairs.GetEnumerator() | Sort-Object { $_.Value.Count } -Descending | ForEach-Object { $parts = $_.Key -split ' / '; [ordered]@{ 'Owning department' = $parts[0]; 'Department with access' = $parts[1]; 'Shares' = $_.Value.Count; 'Examples' = (Join-List $_.Value.Items ' | ') } })
$ShareFilters = @(@{ Label = 'Anyone (anonymous)'; Tag = 'anyone' }, @{ Label = 'Whole organization'; Tag = 'wholeorganization' }, @{ Label = 'External'; Tag = 'external' }, @{ Label = 'Outside owning team'; Tag = 'outsideteam' }, @{ Label = 'Cross-department'; Tag = 'crossdept' }, @{ Label = 'Specific people'; Tag = 'specificpeople' }, @{ Label = 'SharePoint'; Tag = 'sharepoint' }, @{ Label = 'OneDrive'; Tag = 'onedrive' }, @{ Label = 'Teams'; Tag = 'teams' }, @{ Label = 'Exchange'; Tag = 'exchange' }, @{ Label = 'Urgent'; Tag = 'urgent' }, @{ Label = 'Review'; Tag = 'review' })
$GuestFilters = @(@{ Label = 'In teams or groups'; Tag = 'ingroups' }, @{ Label = 'No group memberships'; Tag = 'nogroups' }, @{ Label = 'Disabled'; Tag = 'disabled' })
$shareKv = [ordered]@{
    'Scan coverage'                                        = $shareScanNote
    'Items or resources open to anyone on the internet'    = $shareAnyone
    'Items or resources open to the whole organization'    = $shareOrg
    'Grants to people outside the organization (guests, external links, shared-channel users)' = $shareExternal
    'Grants to insiders outside the owning team'           = $shareOutsideTeam
    'Cross-department grants (owner and recipient in different departments)' = @($shareRows | Where-Object { $_['_tags'] -match '\bcrossdept\b' }).Count
    'Guest accounts in the directory'                      = @($users | Where-Object { $_.userType -eq 'Guest' }).Count
    'Shared mailboxes'                                     = $(if ($exoCollected) { $sharedMailboxCount } else { 'Not collected' })
    'Calendars readable by everyone in the organization / by anonymous users' = $(if ($exoCollected -and -not $SkipInboxRules) { "$(@($calendarShareRows | Where-Object { $_['Level'] -eq 'Default' }).Count) / $(@($calendarShareRows | Where-Object { $_['Level'] -eq 'Anonymous' }).Count)" } else { 'Not scanned' })
    'Calendar sharing policy (external)'                   = $(if ($sharingPolicyText) { $sharingPolicyText } else { 'Not collected' })
}
$shareKvIssues = @{}
if ($shareAnyone) { $shareKvIssues['Items or resources open to anyone on the internet'] = 'share-anonymous' }
if ($shareOrg) { $shareKvIssues['Items or resources open to the whole organization'] = 'share-org' }
if ($shareExternal) { $shareKvIssues['Grants to people outside the organization (guests, external links, shared-channel users)'] = 'share-external' }
if ($shareOutsideTeam) { $shareKvIssues['Grants to insiders outside the owning team'] = 'share-outside-team' }
if ($sharingPolicyText -and $sharingPolicyText -match 'Anonymous|CalendarSharingFreeBusyDetail|CalendarSharingFreeBusyReviewer') { $shareKvIssues['Calendar sharing policy (external)'] = 'sharing-policy-external' }

# Roll-up: urgent items (investigate today) versus items to review, and the plain-language situation table
$iocUrgent = [ordered]@{}; $iocReview = [ordered]@{}
$iocSituationRows = @()
function Add-IocSituation {
    param([string]$Indicator, [int]$Count, [ValidateSet('Urgent', 'Review', 'Clear', 'Not checked')][string]$Status, [string]$Found, [string]$Why, [string]$Anchor, [string]$Issue)
    if ($Status -eq 'Urgent') { $script:iocUrgentMap[$Indicator] = $Count } elseif ($Status -eq 'Review') { $script:iocReviewMap[$Indicator] = $Count }
    $script:iocSituationRows += [ordered]@{ 'Indicator' = $Indicator; 'Status' = $Status; 'What we found' = $Found; 'Why it matters' = $Why; 'Details' = 'Show'; '_links' = @{ 'Details' = "#$Anchor" }; '_tags' = $Status.ToLower().Replace(' ', ''); '_issues' = $(if ($Status -in @('Urgent', 'Review') -and $Issue) { @{ 'Status' = $Issue } } else { $null }) }
}
$script:iocUrgentMap = $iocUrgent; $script:iocReviewMap = $iocReview; $script:iocSituationRows = @()
function Get-IocStatus { param([int]$Count, [bool]$Checked, [bool]$Urgent) if (-not $Checked) { 'Not checked' } elseif ($Count -eq 0) { 'Clear' } elseif ($Urgent) { 'Urgent' } else { 'Review' } }

$fwdTotal = $fwdMailboxExternal.Count + $fwdRuleExternal.Count + $fwdTransportAll.Count
Add-IocSituation 'Mail automatically forwarded outside the organization' $fwdTotal (Get-IocStatus $fwdTotal $exoCollected $true) $(if (-not $exoCollected) { 'Exchange Online was not collected.' } elseif ($fwdTotal) { "$($fwdMailboxExternal.Count) mailbox(es) forward to an external address$(if ($fwdMailboxExternal.Count) { " ($(Join-List ($fwdMailboxExternal | Select-Object -First 3 | ForEach-Object { "$($_['Mailbox']) to $($_['Forwarding SMTP address'])" })))" }); $($fwdRuleExternal.Count) inbox rule(s) forward externally$(if ($SkipInboxRules) { ' (rules not scanned)' }); $($fwdTransportAll.Count) transport rule(s) redirect, copy or Bcc mail." } else { "No mailbox, inbox rule or transport rule forwards mail outside the organization$(if ($SkipInboxRules) { ' (inbox rules were not scanned)' })." }) 'After taking over a mailbox, attackers forward a copy of everything to an outside address so they keep reading mail after the password changes. In payment fraud this is how they learn invoice cycles and intercept wire instructions.' 'ioc-forwarding' 'forwarding'
Add-IocSituation 'Inbox rules that hide, delete or redirect mail' $suspiciousRuleRows.Count (Get-IocStatus $suspiciousRuleRows.Count ($exoCollected -and -not $SkipInboxRules) $true) $(if (-not $exoCollected) { 'Exchange Online was not collected.' } elseif ($SkipInboxRules) { 'Inbox rules were not scanned (-SkipInboxRules).' } elseif ($suspiciousRuleRows.Count) { "$($suspiciousRuleRows.Count) rule(s) on $(@($suspiciousRuleRows | ForEach-Object { $_['Mailbox'] } | Select-Object -Unique).Count) mailbox(es): $(Join-List ($suspiciousRuleRows | Select-Object -First 3 | ForEach-Object { "$($_['Mailbox']) rule '$($_['Rule'])' $($_['Why it is suspicious'])" }) ' | ')" } else { 'No inbox rule deletes, hides (RSS Feeds, Archive, Deleted Items) or externally forwards mail.' }) 'The signature of business email compromise: a rule named "." or blank that moves replies to RSS Feeds or deletes them, so the victim never sees the conversation the attacker is having in their name.' 'ioc-rules' 'ioc-inbox-rule-hide'
Add-IocSituation 'Known-abused OAuth applications present' $rogueAppRows.Count (Get-IocStatus $rogueAppRows.Count (-not (Get-GraphFailure 'enterprise applications')) $true) $(if ($rogueAppRows.Count) { Join-List ($rogueAppRows | ForEach-Object { "$($_['Application']) ($($_['Typical abuse']))" }) '; ' } else { "None of the $($RogueApps.Count) applications on the known-abused list (PerfectData Software, eM Client, rclone, Newsletter Software Supermailer and others) is present." }) 'These are legitimate products that attackers consent to with a stolen session so they can export the whole mailbox or OneDrive, send bulk phishing, or keep access after a password reset. Huntress reports under 2% false positives when they appear in a tenant.' 'ioc-apps' 'ioc-rogue-app'
Add-IocSituation 'Unverified-publisher applications with mail, file or contact access' $unverifiedAppRows.Count (Get-IocStatus $unverifiedAppRows.Count (-not (Get-GraphFailure 'enterprise applications')) $false) $(if ($unverifiedAppRows.Count) { "$($unverifiedAppRows.Count) application(s): $(Join-List ($unverifiedAppRows | Select-Object -First 4 | ForEach-Object { $_['Application'] }))$(if ($unverifiedAppRows.Count -gt 4) { ' ...' })" } else { 'Every application that can read mail, files or contacts comes from a verified publisher or Microsoft.' }) 'Consent phishing works by getting one user to approve an app that then reads mail and files through the API, bypassing MFA entirely. An unverified publisher with those permissions is the pattern.' 'ioc-apps' 'ioc-unverified-app'
Add-IocSituation 'New enterprise applications and application credentials' ($iocNewAppRows.Count + $iocAppRows.Count + $iocNewCredRows.Count) (Get-IocStatus ($iocNewAppRows.Count + $iocAppRows.Count + $iocNewCredRows.Count) $iocAuditAvailable $false) "$($iocNewAppRows.Count) application(s) added, $($iocNewCredRows.Count) new secret(s) or certificate(s) on app registrations, $($iocAppRows.Count) credential or ownership audit event(s) in $AuditDays days." 'A secret added to an existing app gives the holder that app''s permissions without any sign-in or MFA; new apps may be the consent-phishing payload.' 'ioc-apps' 'ioc-app-cred'
$anyoneTotal = $anonLinkCount + $shareAnyone
$anyoneChecked = $ualCollected -or (-not $SkipShareScan -and $shareSitesScanned -gt 0)
$anyoneStatus = if (-not $anyoneChecked) { 'Urgent' } elseif ($anyoneTotal -gt 0) { 'Urgent' } else { 'Clear' }
$anyoneFound = if (-not $anyoneChecked) { "Cannot be checked: $(if ($ualEnabled -eq $false) { 'the unified audit log is turned off' } else { 'the unified audit log was not searched' }) and the share scan $(if ($SkipShareScan) { 'was skipped' } elseif ($shareSitesDenied) { "could not read any of the $shareSitesDenied site(s)" } else { 'found no readable site' }). Exposure to the internet cannot be ruled out; see the Critical finding 'Audit logging (unified audit log)'." }
               elseif ($anyoneTotal) { "$shareAnyone Anyone link(s) or anonymous calendar(s) found on the $shareSitesScanned site(s) this account can read (section 19)$(if ($ualCollected) { "; $anonLinkCount Anyone link(s) created in the last $AuditDays days per the audit log" } else { '; the audit log is off, so links on sites this account cannot read are invisible' })." }
               else { "None on the $shareSitesScanned site(s) scanned$(if ($ualCollected) { " and none created in the last $AuditDays days" } else { "; the unified audit log is $(if ($ualEnabled -eq $false) { 'turned off' } else { 'not searched' }), so links on the $shareSitesDenied site(s) this account cannot read are invisible" })." }
Add-IocSituation 'Files shared with anyone (anonymous links)' $anyoneTotal $anyoneStatus $anyoneFound 'An Anyone link opens a file to the whole internet with no sign-in and no record of who opened it; attackers use it to hand data to themselves. Not being able to check is itself a finding.' 'ioc-sharing' $(if (-not $anyoneChecked) { 'ual-off' } else { 'share-anonymous' })
Add-IocSituation 'Files shared with people outside the organization' $extShareEventCount $(if ($ualEnabled -eq $false) { 'Urgent' } else { Get-IocStatus $extShareEventCount $ualCollected $false })  $(if ($ualEnabled -eq $false) { 'Cannot be checked: the unified audit log is turned off, so this activity is not recorded. Turn it on (Critical finding "Audit logging (unified audit log)"); examiners treat the absence of logging as a finding in itself.' } elseif ($ualNote) { $ualNote } elseif ($extShareEventCount) { "$extShareEventCount external sharing event(s) in $AuditDays days$(if ($extShareRows.Count) { "; usage reports show $($extShareRows.Count) user(s) shared $extShareTotal file(s) externally in 30 days" })." } else { "No files were shared with outside people in the last $AuditDays days." }) 'Legitimate for auditors and partners, but a compromised account sharing a folder to an outside address is a common second exit for data.' 'ioc-sharing' $(if ($ualEnabled -eq $false) { 'ual-off' } else { 'ioc-ext-share' })
Add-IocSituation 'Unusual volume of file downloads or sync' ($massDownloadUrgent + $massDownloadReview) $(if ($ualEnabled -eq $false) { 'Urgent' } else { Get-IocStatus ($massDownloadUrgent + $massDownloadReview) $ualCollected ($massDownloadUrgent -gt 0) }) $(if ($ualEnabled -eq $false) { 'Cannot be checked: the unified audit log is turned off, so this activity is not recorded. Turn it on (Critical finding "Audit logging (unified audit log)"); examiners treat the absence of logging as a finding in itself.' } elseif ($ualNote) { $ualNote } elseif ($massDownloadUrgent + $massDownloadReview) { "$massDownloadUrgent user(s) downloaded $MassDownloadThreshold or more files in $downloadDays days; $massDownloadReview more are well above the median ($dlMedian). Top: $(Join-List ($downloadRows | Where-Object { $_['_tags'] -ne 'normal' } | Select-Object -First 3 | ForEach-Object { "$($_['User']) ($($_["Files downloaded ($downloadDays days)"]))" }))." } else { "$downloadTotal download events in $downloadDays days across $($dlBy.Count) user(s); nobody is far above the median ($dlMedian per user)." }) 'Mass download or a full OneDrive sync from a new device is how a departing employee or an intruder takes the data with them. Tools such as rclone show up here.' 'ioc-downloads' $(if ($ualEnabled -eq $false) { 'ual-off' } else { 'ioc-mass-download' })
Add-IocSituation 'Mailbox rule, forwarding and permission changes in the audit log' $auditRuleDanger $(if ($ualEnabled -eq $false) { 'Urgent' } else { Get-IocStatus $auditRuleDanger $ualCollected $true }) $(if ($ualEnabled -eq $false) { 'Cannot be checked: the unified audit log is turned off, so this activity is not recorded. Turn it on (Critical finding "Audit logging (unified audit log)"); examiners treat the absence of logging as a finding in itself.' } elseif ($ualNote) { $ualNote } elseif ($auditRuleDanger) { "$auditRuleDanger change(s) that forward, redirect, delete, hide or grant access, out of $($auditRuleRows.Count) mailbox configuration events in $AuditDays days." } else { "$($auditRuleRows.Count) mailbox configuration event(s) in $AuditDays days; none forwards, deletes or grants access." }) 'The audit log records who created each rule, from which IP address and client, including rules created and then deleted, which the mailbox scan above cannot see.' 'ioc-auditrules' $(if ($ualEnabled -eq $false) { 'ual-off' } else { 'ioc-audit-rule-change' })
Add-IocSituation 'Users who clicked malicious links (Safe Links)' ($urlClickedThrough + $urlBlocked) (Get-IocStatus ($urlClickedThrough + $urlBlocked) $urlClicksCollected ($urlClickedThrough -gt 0)) $(if (-not $script:HasDefenderP1) { 'Not available: Safe Links requires Defender for Office 365 (Plan 1 or 2).' } elseif (-not $urlClicksCollected) { 'Safe Links report could not be read.' } elseif ($urlClickedThrough + $urlBlocked) { "$urlBlocked malicious link(s) blocked and $urlClickedThrough clicked through the warning in 7 days by $(@($urlClickRows | ForEach-Object { $_['User'] } | Select-Object -Unique).Count) user(s)." } else { 'No malicious URL clicks in the last 7 days.' }) 'A user who clicked through a Safe Links warning may have entered their password on a phishing page; that account should be checked today.' 'ioc-urls' 'ioc-url-click'
Add-IocSituation 'Domain or federation changes' $iocDomainRows.Count (Get-IocStatus $iocDomainRows.Count $iocAuditAvailable $true) $(if ($iocDomainRows.Count) { Join-List ($iocDomainRows | Select-Object -First 3 | ForEach-Object { "$($_['Date (local)']) $($_['Activity']) by $($_['Performed by'])" }) ' | ' } else { "No domain was added, verified or switched to federated authentication in $AuditDays days." }) 'Switching a domain to federation lets an attacker sign in as anyone without a password or MFA; a hallmark of tenant-level compromise.' 'ioc-domains' 'ioc-domain-change'
Add-IocSituation 'Users flagged at risk by Entra ID Protection' $riskyUserRows.Count (Get-IocStatus $riskyUserRows.Count (-not $iocRiskNote) $true) $(if ($iocRiskNote) { $iocRiskNote } elseif ($riskyUserRows.Count) { Join-List ($riskyUserRows | Select-Object -First 4 | ForEach-Object { "$($_['UPN']) ($($_['Risk level']))" }) } else { 'No users are currently at risk; ' + "$($riskDetectionRows.Count) risk detection(s) in $AuditDays days." }) 'Microsoft correlates leaked credentials, impossible travel and malicious IP addresses; "at risk" means probably compromised.' 'ioc-risk' 'ioc-risky-user'
Add-IocSituation 'Successful legacy authentication sign-ins' $legacySignInRows.Count (Get-IocStatus $legacySignInRows.Count $iocSignInsCollected $true) $(if (-not $iocSignInsCollected) { $(if ($iocSignInNote) { $iocSignInNote } else { 'Sign-in log not collected.' }) } elseif ($legacySignInRows.Count) { Join-List ($legacySignInRows | Select-Object -First 3 | ForEach-Object { "$($_['User']) via $($_['Protocol / client']) ($($_['Successful sign-ins']))" }) '; ' } else { "No successful IMAP, POP, SMTP, ActiveSync or other legacy sign-ins in $signInDays days." }) 'Legacy protocols skip MFA; a successful legacy sign-in is a password being used with no second factor, possibly by an attacker.' 'ioc-signins' 'ioc-legacy-signin'
Add-IocSituation 'Password spray against the tenant' @($sprayRows | Where-Object { $_['Distinct accounts targeted'] -ge 5 }).Count (Get-IocStatus @($sprayRows | Where-Object { $_['Distinct accounts targeted'] -ge 5 }).Count $iocSignInsCollected $true) $(if (-not $iocSignInsCollected) { 'Sign-in log not collected.' } elseif ($sprayRows.Count) { "$($sprayRows.Count) source address(es) failed passwords against several accounts; worst: $($sprayRows[0]['Source IP address']) ($($sprayRows[0]['Country'])) hit $($sprayRows[0]['Distinct accounts targeted']) accounts $($sprayRows[0]['Failed passwords']) times." } else { "$badPwTotal wrong-password attempts against $badPwUsers account(s); no address targeted three or more accounts." }) 'One address trying passwords across many accounts is a spray; any weak or reused password falls to it, and MFA plus legacy-auth blocking is the defence.' 'ioc-signins' 'ioc-spray'
Add-IocSituation 'Sign-ins from more than one country, MFA denials, Conditional Access blocks' ($multiCountryRows.Count + $mfaDeniedRows.Count + $caBlockedRows.Count) (Get-IocStatus ($multiCountryRows.Count + $mfaDeniedRows.Count + $caBlockedRows.Count) $iocSignInsCollected $false) $(if (-not $iocSignInsCollected) { 'Sign-in log not collected.' } else { "$($multiCountryRows.Count) user(s) signed in from more than one country (home country $homeCountry); $($mfaDeniedRows.Count) user(s) denied or failed MFA; $($caBlockedRows.Count) user(s) were blocked by Conditional Access after a correct password; $($lockoutRows.Count) locked out." }) 'Each of these means a password was accepted somewhere unexpected: the user travelled, uses a VPN, or someone else has the credentials.' 'ioc-signins' 'ioc-multi-country'
Add-IocSituation 'Privileged roles, MFA methods, password resets and policy changes' ($iocRoleRows.Count + $iocMfaRows.Count + $iocPwAdminRows.Count + $iocCaRows.Count) (Get-IocStatus ($iocRoleRows.Count + $iocMfaRows.Count + $iocPwAdminRows.Count + $iocCaRows.Count) $iocAuditAvailable $false) $(if (-not $iocAuditAvailable) { "Directory audit log not available: $(Get-GraphFailure 'role assignment audit events')" } else { "$($iocRoleRows.Count) privileged role assignment(s), $($iocMfaRows.Count) MFA method registration(s) or change(s), $($iocPwAdminRows.Count) password reset(s) by another account, $($iocCaRows.Count) Conditional Access or security policy change(s) in $AuditDays days." }) 'Routine administration and attacker persistence look identical in the log; each item should match a ticket or a person who says "that was me".' 'ioc-admin' 'ioc-review'
Add-IocSituation 'Accounts created, guests invited, mailbox delegates and Send As grants' ($iocNewUserRows.Count + $fullAccessRows.Count + $sendAsRows.Count) (Get-IocStatus ($iocNewUserRows.Count + $fullAccessRows.Count + $sendAsRows.Count) ($iocAuditAvailable -or $exoCollected) $false) "$($iocNewUserRows.Count) account(s) created or invited in $AuditDays days; $($fullAccessRows.Count) non-owner Full Access delegate(s)$(if ($SkipInboxRules) { ' (not scanned)' }); $($sendAsRows.Count) Send As grant(s)." 'Attackers create their own accounts and grant themselves mailbox access so they survive the victim''s password reset.' 'ioc-accounts' 'ioc-delegate'
if ($mailboxAuditDisabled -eq $true) { Add-IocSituation 'Mailbox auditing' 1 'Urgent' 'Mailbox auditing is disabled organization-wide.' 'Without it there is no record of who read, moved or sent mail, so none of the above can be investigated.' 'ioc-auditrules' 'ioc-audit-off' }
$iocSituationRows = @($script:iocSituationRows)
$IocSituationFilters = @(@{ Label = 'Urgent'; Tag = 'urgent' }, @{ Label = 'Review'; Tag = 'review' }, @{ Label = 'Clear'; Tag = 'clear' }, @{ Label = 'Not checked'; Tag = 'notchecked' })
$iocUrgentCount = 0; foreach ($v in $iocUrgent.Values) { $iocUrgentCount += [int]$v }
$iocReviewCount = 0; foreach ($v in $iocReview.Values) { $iocReviewCount += [int]$v }
$iocCollected = $iocAuditAvailable -or $exoCollected -or $iocSignInsCollected
$iocStatus = if (-not $iocCollected) { 'Not collected' } elseif ($iocUrgentCount -gt 0) { "$iocUrgentCount urgent, $iocReviewCount to review" } elseif ($iocReviewCount -gt 0) { "$iocReviewCount to review" } else { 'None found' }
$iocEvidence = if (-not $iocCollected) { 'Directory audit log, Exchange and sign-in log were all unavailable.' } else { (Join-List (@($iocUrgent.Keys | ForEach-Object { "URGENT $_ ($($iocUrgent[$_]))" }) + @($iocReview.Keys | ForEach-Object { "$_ ($($iocReview[$_]))" })) '; ') }
if (-not $iocEvidence) { $iocEvidence = "No indicators in the last $AuditDays days across the directory audit log$(if ($exoCollected) { ', Exchange Online' })$(if ($ualCollected) { ', the unified audit log' })$(if ($iocSignInsCollected) { ', the sign-in log' })$(if ($script:HasEntraP2 -and -not $iocRiskNote) { ' and Entra ID Protection' })." }
$iocKv = [ordered]@{}
$iocKvIssues = @{}
$iocKv['Sources checked'] = Join-List (@($(if ($iocAuditAvailable) { 'Entra directory audit log' }), $(if ($iocSignInsCollected) { "Entra sign-in log ($signInDays days)" }), $(if ($script:HasEntraP2 -and -not $iocRiskNote) { 'Entra ID Protection' }), $(if ($exoCollected) { 'Exchange Online (mailboxes, rules, permissions)' }), $(if ($ualCollected) { "unified audit log (sharing and mailbox changes $AuditDays days, downloads $downloadDays days)" }), $(if ($urlClicksCollected) { 'Safe Links clicks (7 days)' }), 'enterprise applications', 'app registrations') | Where-Object { $_ }) ', '
if ($ualNote) { $iocKv['Unified audit log'] = $ualNote; if ($ualEnabled -eq $false) { $iocKvIssues['Unified audit log'] = 'ioc-audit-off' } }
if ($iocSignInsCollected) { $iocKv['Successful sign-ins analysed'] = "$($okSignIns.Count) in $signInDays days from $($countryRows.Count) countr$(if ($countryRows.Count -eq 1) { 'y' } else { 'ies' }); home country $homeCountry"; $iocKv['Wrong-password attempts'] = "$badPwTotal against $badPwUsers account(s)" }
if ($exoCollected -and $mailboxAuditDisabled -eq $false) { $iocKv['Mailbox auditing'] = 'Enabled (organization-wide)' }
if ($mailboxAuditDisabled -eq $true) { $iocKv['Mailbox auditing'] = 'Disabled'; $iocKvIssues['Mailbox auditing'] = 'ioc-audit-off' }
$IocAuditFilters = @(@{ Label = 'By another account'; Tag = 'other' }, @{ Label = 'Self-service'; Tag = 'self' })
# Defender for Office 365 / EOP overall status
$safeLinksActive  = @($safeLinksRows | Where-Object { $_['Email'] -eq $true }).Count
$safeAttachActive = @($safeAttachRows | Where-Object { $_['Enabled'] -eq $true -and $_['Action'] -ne 'Allow' }).Count
$presetOn = @($presetRows | Where-Object { $_['State'] -eq 'Enabled' }).Count
$defenderStatus = if (-not $exoCollected -or -not $defenderCollected) { 'Not collected' }
                  elseif (-not $script:HasDefenderP1) { 'EOP only (Defender not licensed)' }
                  elseif ($safeLinksActive -and $safeAttachActive) { 'Configured' }
                  else { 'Licensed, not configured' }
$defenderEvidence = if (-not $exoCollected) { $exoKv['Exchange Online'] }
    else { "Defender for Office 365: $(if ($script:HasDefenderP2) { 'Plan 2' } elseif ($script:HasDefenderP1) { 'Plan 1' } else { 'not licensed' }). Preset security policies enabled: $presetOn of $(@($presetRows).Count). Anti-spam policies: $(@($spamRows).Count); anti-phishing: $(@($phishRows).Count); Safe Links active: $safeLinksActive; Safe Attachments active: $safeAttachActive. Defender for SharePoint/OneDrive/Teams: $(if ($atpGlobal) { $atpGlobal.EnableATPForSPOTeamsODB } else { 'n/a' }); Safe Documents: $(if ($atpGlobal) { $atpGlobal.EnableSafeDocs } else { 'n/a' })." }

$script:summaryRows = @()
function Add-SummaryRow { param($control, $status, $evidence, $link) $script:summaryRows += [ordered]@{ 'Control' = $control; 'Status' = $status; 'Evidence' = $evidence; 'Section' = $link } }

Add-SummaryRow 'Security defaults' $(if ($secDefaults.isEnabled) { 'Enabled' } else { 'Disabled' }) $(if ($secDefaults.isEnabled) { 'Security defaults enforce MFA registration and block legacy authentication tenant-wide; Conditional Access cannot be used alongside them.' } else { 'Security defaults are off; MFA and legacy authentication controls depend on Conditional Access.' }) 'Authentication Methods'

$mfaStatus = if ($secDefaults.isEnabled) { 'Enforced (security defaults)' } elseif ($mfaAllUsersEnforced.Count) { 'Enforced' } elseif (@($mfaPolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not enforced' }
$mfaEvidence = if ($mfaAllUsersEnforced.Count) { "Enabled policy covering all users and all cloud apps: " + (Join-List ($mfaAllUsersEnforced | ForEach-Object { "$($_.displayName) [$(Get-CaScopeText $_)]" }) ' || ') }
               elseif ($mfaPolicies.Count) { "MFA policies exist but none that is enabled targets all users and all cloud apps: " + (Join-List ($mfaPolicies | ForEach-Object { "$($_.displayName) ($($_.state))" }) '; ') }
               elseif ($secDefaults.isEnabled) { 'Security defaults enabled.' } else { 'No Conditional Access policy requires MFA and security defaults are off.' }
Add-SummaryRow 'MFA enforced for sign-in' $mfaStatus $mfaEvidence 'Conditional Access'

Add-SummaryRow 'MFA registered by all users' $(if (-not @($regDetails).Count) { 'Not available' } elseif ($mfaRegistered -eq $enabledMembers.Count) { 'Yes' } else { 'No' }) $(if (@($regDetails).Count) { "$mfaRegistered of $($enabledMembers.Count) enabled member accounts have an MFA method registered; $(@($noMfaRows).Count) do not (listed under Users)." } else { 'Registration report requires Entra ID P1.' }) 'Users'

Add-SummaryRow 'Phishing-resistant MFA (passkeys) rollout' $passkeyStatus $passkeyEvidence 'Authentication Methods'

Add-SummaryRow 'Phishing-resistant MFA for administrators' $adminPhishStatus $adminPhishEvidence 'Privileged Users'

$legacyStatus = if ($secDefaults.isEnabled) { 'Blocked (security defaults)' } elseif ($legacyBlockedAll.Count) { 'Blocked' } elseif (@($legacyPolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not blocked by Conditional Access' }
$legacyEvidence = if ($legacyBlockedAll.Count) { "Block policy for legacy clients covering all users: " + (Join-List ($legacyBlockedAll | ForEach-Object { $_.displayName }) '; ') } elseif ($legacyPolicies.Count) { "Legacy block policies exist but are scoped or not enabled: " + (Join-List ($legacyPolicies | ForEach-Object { "$($_.displayName) ($($_.state))" }) '; ') } elseif ($secDefaults.isEnabled) { 'Security defaults block legacy authentication protocols tenant-wide (no Conditional Access policy is needed).' } else { 'No Conditional Access policy blocks Exchange ActiveSync and other (legacy) clients.' }
if ($exoCollected) { $legacyEvidence += " Exchange: SMTP AUTH disabled tenant-wide = $($exoKv['SMTP AUTH disabled tenant-wide']); POP enabled on $($exoKv['Mailboxes with POP enabled']) mailboxes; IMAP enabled on $($exoKv['Mailboxes with IMAP enabled']) mailboxes." }
Add-SummaryRow 'Legacy authentication protocols disabled' $legacyStatus $legacyEvidence 'Conditional Access'

$devStatus = if ($deviceEnforcedAll.Count) { 'Enforced' } elseif (@($devicePolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not enforced' }
$devEvidence = if ($devicePolicies.Count) { Join-List ($devicePolicies | ForEach-Object { "$($_.displayName) ($($_.state)): $(Get-CaScopeText $_)" }) ' || ' } else { 'No Conditional Access policy requires a compliant or Entra joined device; any device with valid credentials can access Microsoft 365.' }
Add-SummaryRow 'Conditional Access restricts access to Entra joined or compliant devices' $devStatus $devEvidence 'Conditional Access'

Add-SummaryRow 'Intune deployed to Windows computers' $(if (Get-GraphFailure 'Intune managed devices') { 'Not available' } elseif ($winEntra.Count -eq 0) { 'No Windows devices in Entra' } elseif ($winManaged.Count -eq $winEntra.Count) { 'All enrolled' } elseif ($winManaged.Count -gt 0) { 'Partial' } else { 'Not deployed' }) "$(if (Get-GraphFailure 'Intune managed devices') { "Intune device list could not be read. $(Get-GraphFailure 'Intune managed devices') " })$($winManaged.Count) of $($winEntra.Count) enabled Windows devices in Entra ID are enrolled in Intune; $($winEntra.Count - $winManaged.Count) are not (listed under Managed Devices). Total Intune-enrolled devices: $(@($managedDevices).Count)." 'Managed Devices'

Add-SummaryRow 'Intune compliance policies' $(if (Get-GraphFailure 'compliance policies') { 'Not available' } elseif (@($compliancePolicies).Count) { "$(@($compliancePolicies).Count) policies" } else { 'None' }) $(if (Get-GraphFailure 'compliance policies') { "Could not be read. $(Get-GraphFailure 'compliance policies')" } elseif (@($compliancePolicies).Count) { Join-List ($compliancePolicies | ForEach-Object { "$($_.displayName) [$(Get-Platform $_.'@odata.type')]" }) '; ' } else { 'No compliance policies exist, so no device can be evaluated as compliant.' }) 'MDM Compliance Policies'

$mobileCompliance = @($compliancePolicies | Where-Object { (Get-Platform ([string]$_.'@odata.type')) -match 'Android|iOS' })
Add-SummaryRow 'Intune compliance policies for mobile devices' $(if (Get-GraphFailure 'compliance policies') { 'Not available' } elseif ($mobileCompliance.Count) { "$($mobileCompliance.Count) policies" } else { 'None' }) $(if (Get-GraphFailure 'compliance policies') { "Could not be read. $(Get-GraphFailure 'compliance policies')" } elseif ($mobileCompliance.Count) { Join-List ($mobileCompliance | ForEach-Object { "$($_.displayName) ($(Get-Platform ([string]$_.'@odata.type')))" }) '; ' } else { 'No compliance policy targets iOS/iPadOS or Android. Phones and tablets that read institution mail have no enforced passcode, encryption or operating-system requirement.' }) 'MDM Compliance Policies'

Add-SummaryRow 'Intune configuration and security policies' $(if (Get-GraphFailure 'device configuration profiles') { 'Not available' } else { "$(@($profileRows).Count) policies" }) $(if (Get-GraphFailure 'device configuration profiles') { "Could not be read. $(Get-GraphFailure 'device configuration profiles')" } elseif (@($profileRows).Count) { Join-List ($profileRows | ForEach-Object { "$($_.Policy) [$($_.Kind)]" }) '; ' } else { 'No configuration profiles, settings catalog policies, or endpoint security policies exist.' }) 'Intune Policies'

Add-SummaryRow 'Temporary Access Pass (TAP) implemented' $(if (-not $tap) { 'Not available' } elseif ($tap.state -eq 'enabled') { 'Enabled' } else { 'Disabled' }) $(if ($tap) { "State: $($tap.state); enabled for: $(Resolve-AuthTargets $tap.includeTargets); default lifetime $($tap.defaultLifetimeInMinutes) minutes; one-time use: $($tap.isUsableOnce)" } else { "Could not read the authentication methods policy. $(Get-GraphFailure 'authentication methods policy')" }) 'Authentication Methods'

Add-SummaryRow 'Users restricted from joining devices to Entra ID' $(if ($devRegPolicy) { $(if ($devRegKv['Users may join devices to Entra ID'] -match 'all') { 'All users allowed' } elseif ($devRegKv['Users may join devices to Entra ID'] -match 'no') { 'Restricted (none)' } else { 'Restricted (selected)' }) } else { 'Not available' }) $(if ($devRegPolicy) { "Join allowed for: $($devRegKv['Users may join devices to Entra ID']); MFA required to join: $($devRegPolicy.multiFactorAuthConfiguration); device quota per user: $($devRegPolicy.userDeviceQuota)" } else { '' }) 'Authentication Methods'

$fwdMailboxCount = if ($exoCollected) { [int]$exoKv['Mailboxes forwarding to an external address'] } else { 0 }
$fwdRuleCount    = if ($exoCollected -and -not $SkipInboxRules) { [int]$exoKv['Inbox rules forwarding to an external address'] } else { 0 }
$fwdSpamBlocks   = $exoCollected -and ([string]$exoKv['Outbound spam policy automatic forwarding'] -match '(?i)\b(Off|Automatic)\b') -and ([string]$exoKv['Outbound spam policy automatic forwarding'] -notmatch '(?i)\bOn\b')
$fwdRemoteAllows = $exoCollected -and ($exoKv['Default remote domain (*) allows auto-forward'] -eq $true)
$fwdStatus = if (-not $exoCollected) { 'Not collected' }
             elseif ($fwdMailboxCount -gt 0 -or $fwdRuleCount -gt 0) { 'Active external forwarding found' }
             elseif ($fwdSpamBlocks -and -not $fwdRemoteAllows) { 'Blocked' }
             elseif ($fwdSpamBlocks) { 'Blocked (outbound spam policy)' }
             else { 'Allowed by policy' }
Add-SummaryRow 'Automatic email forwarding outside the tenant' $fwdStatus $(if ($exoCollected) { "Default remote domain auto-forward: $($exoKv['Default remote domain (*) allows auto-forward']); outbound spam policy: $($exoKv['Outbound spam policy automatic forwarding']) (Automatic and Off both block external auto-forwarding; the outbound spam policy overrides the remote domain setting); mailbox-level external forwards: $($exoKv['Mailboxes forwarding to an external address']); inbox-rule external forwards: $($exoKv['Inbox rules forwarding to an external address']); transport rules redirecting mail: $($exoKv['Transport rules that redirect, copy, or bcc mail'])" } else { $exoKv['Exchange Online'] }) 'Exchange Online'

Add-SummaryRow 'User consent to applications' $(if ($userConsentSetting -match 'Do not allow') { 'Restricted' } else { 'Allowed' }) "Setting: $userConsentSetting; admin consent workflow: $($appGovKv['Admin consent request workflow'])" 'Enterprise Applications'

Add-SummaryRow 'SPF record (DNS)' $spfStatus $(if ($emailAuthRows.Count) { Join-List ($emailAuthRows | ForEach-Object { "$($_['Domain']): $($_['SPF'])" }) '; ' } elseif (-not $dnsAvailable) { 'Resolve-DnsName is not available on this computer.' } else { 'No email-enabled custom domains.' }) 'Exchange Online'

Add-SummaryRow 'DKIM signing (Exchange Online and DNS)' $dkimStatus $(if ($emailAuthRows.Count) { Join-List ($emailAuthRows | ForEach-Object { "$($_['Domain']): signing $($_['DKIM signing (Exchange)']), DNS $($_['DKIM DNS (selector1 / selector2)'])" }) '; ' } elseif (-not $dnsAvailable) { 'Resolve-DnsName is not available on this computer.' } else { 'No email-enabled custom domains.' }) 'Exchange Online'

Add-SummaryRow 'DMARC policy (DNS)' $dmarcStatus $(if ($emailAuthRows.Count) { Join-List ($emailAuthRows | ForEach-Object { "$($_['Domain']): p=$($_['DMARC policy'])$(if ($_['DMARC coverage']) { ", pct $($_['DMARC coverage'])" })$(if ($_['DMARC reporting (rua)']) { ", reporting $($_['DMARC reporting (rua)'])" })" }) '; ' } elseif (-not $dnsAvailable) { 'Resolve-DnsName is not available on this computer.' } else { 'No email-enabled custom domains.' }) 'Exchange Online'

Add-SummaryRow 'Email protection (Defender for Office 365 / EOP)' $defenderStatus $defenderEvidence 'Defender'

Add-SummaryRow 'Endpoint protection (Defender for Endpoint)' $mdeStatus $mdeEvidence 'Defender'

Add-SummaryRow 'Microsoft Secure Score' $(if ($null -eq $scoreCurrent) { 'Not available' } elseif ($scoreCurrent -ge $scoreTarget) { 'Target met' } else { 'Below target' }) $(if ($null -ne $scoreCurrent) { "Current $scoreCurrent of $scoreMax points ($scorePct%); target $scoreTarget points ($TargetScorePercent%); gap $scoreGap points.$(if ($scoreCompareText) { " Comparison: $scoreCompareText." })" } else { "Could not be read. $(Get-GraphFailure 'Microsoft Secure Score')" }) 'Tenant Overview'

Add-SummaryRow 'Privileged role assignments' "$(@($roleRows).Count) assignments" $(if ($roles) { Join-List (($roles | Where-Object { @($_.members).Count -gt 0 } | Sort-Object displayName) | ForEach-Object { "$($_.displayName): $(@($_.members).Count)" }) '; ' } else { '' }) 'Privileged Users'

Add-SummaryRow 'Indications of compromise to review' $iocStatus $iocEvidence 'Indications of Compromise'

Add-SummaryRow 'Audit logging (unified audit log)' $(if (-not $exoCollected) { 'Not collected' } elseif ($null -eq $ualEnabled) { 'Not available' } elseif ($ualEnabled) { 'Enabled' } else { 'Disabled' }) $(if (-not $exoCollected) { 'Exchange Online was not collected.' } elseif ($null -eq $ualEnabled) { 'Get-AdminAuditLogConfig could not be read.' } elseif ($ualEnabled) { "Unified audit log ingestion is on; mailbox auditing is $(if ($mailboxAuditDisabled) { 'disabled' } else { 'enabled' }) organization-wide." } else { 'Unified audit log ingestion is turned off. Sharing, downloads, mailbox rules and administrator actions are not being recorded, so a compromise cannot be investigated and examiners have no activity evidence.' }) 'Indications of Compromise'

Add-SummaryRow 'Teams guest and external access' $teamsStatus $teamsEvidence 'Teams'

Add-SummaryRow 'SharePoint and OneDrive external sharing' $spoStatus $spoEvidence 'SharePoint'

# ----------------------------------------------------------------------------------------------
# Remediation plan: every failed control plus every unimplemented Secure Score action, ranked by priority
# ----------------------------------------------------------------------------------------------
$PriorityRank = @{ 'Critical' = 1; 'High' = 2; 'Medium' = 3; 'Low' = 4 }
$script:remediationItems = New-Object System.Collections.Generic.List[object]
# Findings that a financial institution should treat as this-week work whatever their Secure Score weight (matched against the finding title).
$CriticalFindingPatterns = @(
    '(?i)anonymous users.*(join|meeting)',                       # Restrict anonymous users from joining meetings
    '(?i)more than one global admin',                            # Designate more than one global admin
    '(?i)invited users.*(admitted|lobby)|bypass.*lobby',         # Only invited users should be automatically admitted to Teams meetings
    '(?i)user consent to app',                                   # Ensure user consent to apps accessing company data is not allowed
    '(?i)compliance polic.*(mobile|iOS|Android)|mobile device.*compliance'
)
function Add-Remediation {
    param([string]$Priority, [string]$Finding, [string]$State, [string]$Action, $Points, [string]$Impact, [string]$Section, [string]$Url)
    foreach ($pat in $CriticalFindingPatterns) { if ($Finding -match $pat) { $Priority = 'Critical'; break } }
    foreach ($k in $script:PriorityOverrides.Keys) { if ($Finding -match $k) { $Priority = $script:PriorityOverrides[$k] } }
    # Stable reference: same finding name -> same AH-xxxxxx in every report, so tickets can be tracked to closure
    $sha = [System.Security.Cryptography.SHA256]::Create(); $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($Finding.Trim().ToLowerInvariant())); $sha.Dispose()
    $ref = 'AH-' + (($hash[0..2] | ForEach-Object { $_.ToString('X2') }) -join '')
    $pb = Get-Playbook $Finding $Section
    $script:remediationItems.Add([pscustomobject]@{ Ref = $ref; Priority = $Priority; Rank = $PriorityRank[$Priority]; Finding = $Finding; State = $State; Action = $Action; Points = $Points; Impact = $Impact; Section = $Section; Url = $Url; Playbook = $pb; Frameworks = (Get-FrameworkText $Finding); Phase = $PhaseByPriority[$Priority] })
}
# Control -> (status pattern that counts as a finding, priority, recommended action)
# ----------------------------------------------------------------------------------------------
# Framework alignment: which NIST, CIS and FFIEC/NCUA expectations each control answers.
# Informational mapping for auditors and examiners; section numbers follow the editions named in the README.
# ----------------------------------------------------------------------------------------------
$FrameworkMap = @{
    'MFA enforced for sign-in'                     = @{ Nist = 'PR.AA-03 (users authenticated); SP 800-63B AAL2'; Cis365 = '5.2.2 Conditional Access: MFA for all users and for administrative roles'; CisV8 = '6.3, 6.4, 6.5 (MFA for external, remote and administrative access)'; Ffiec = 'Authentication Guidance (2021): MFA for users and high-risk access; CAT/ACET D3.PC.Am Access and Data Management; IT Handbook Information Security: Authentication' }
    'MFA registered by all users'                  = @{ Nist = 'PR.AA-01, PR.AA-03; SP 800-63B AAL2'; Cis365 = '5.2.3 Authentication methods: all users MFA capable'; CisV8 = '6.3, 6.5'; Ffiec = 'Authentication Guidance (2021); CAT/ACET D3.PC.Am' }
    'Phishing-resistant MFA (passkeys) rollout'    = @{ Nist = 'PR.AA-03; SP 800-63B AAL3 (phishing-resistant authenticators); CISA Phishing-Resistant MFA guidance'; Cis365 = '5.2.2 Phishing-resistant MFA strength for administrators; 5.2.3 Authentication methods'; CisV8 = '6.5 (MFA for administrative access)'; Ffiec = 'Authentication Guidance (2021): stronger controls for high-risk users; CAT/ACET D3.PC.Am' }
    'Phishing-resistant MFA for administrators'    = @{ Nist = 'PR.AA-03, PR.AA-05; SP 800-63B AAL3'; Cis365 = '5.2.2 Phishing-resistant MFA strength required for administrators'; CisV8 = '5.4, 6.5'; Ffiec = 'Authentication Guidance (2021): privileged users; CAT/ACET D3.PC.Am' }
    'Legacy authentication protocols disabled'     = @{ Nist = 'PR.AA-03, PR.PS-01; SP 800-63B (authenticators must be MFA-capable)'; Cis365 = '5.2.2 Block legacy authentication; 6.5.1 Modern authentication for Exchange Online; 6.5.4 SMTP AUTH disabled'; CisV8 = '4.1, 4.8, 6.3'; Ffiec = 'Authentication Guidance (2021); CAT/ACET D3.PC.Im Infrastructure Management' }
    'Security defaults'                            = @{ Nist = 'PR.AA-03'; Cis365 = '5.1.1.1 Security defaults (disabled when Conditional Access is used)'; CisV8 = '6.3, 6.5'; Ffiec = 'Authentication Guidance (2021)' }
    'Conditional Access restricts access to Entra joined or compliant devices' = @{ Nist = 'PR.AA-05 (access permissions incorporate least privilege), PR.PS-01'; Cis365 = '5.2.2 Managed device required for authentication'; CisV8 = '1.1, 4.1, 6.7'; Ffiec = 'CAT/ACET D3.PC.Am, D3.PC.De Device/End-Point Security' }
    'Intune deployed to Windows computers'         = @{ Nist = 'ID.AM-01 (hardware inventory), PR.PS-01 (configuration management)'; Cis365 = 'Not covered by the M365 benchmark (see CIS Intune/Windows benchmarks)'; CisV8 = '1.1, 4.1, 4.2'; Ffiec = 'CAT/ACET D3.PC.De Device/End-Point Security; IT Handbook Information Security: Endpoint controls' }
    'Intune compliance policies for mobile devices' = @{ Nist = 'PR.PS-01, PR.DS-01, PR.AA-05'; Cis365 = 'Not covered (CIS Intune benchmark); 5.2.2 managed device required'; CisV8 = '3.6 (encrypt end-user devices), 4.1, 4.3, 4.11 (remote wipe)'; Ffiec = 'CAT/ACET D3.PC.De Device/End-Point Security: mobile device management; IT Handbook Information Security: mobile devices' }
    'Intune compliance policies'                   = @{ Nist = 'PR.PS-01, PR.DS-01 (data at rest protected)'; Cis365 = 'Not covered (CIS Intune benchmark)'; CisV8 = '3.6 (encrypt end-user devices), 4.1, 4.3'; Ffiec = 'CAT/ACET D3.PC.De; D3.PC.Am' }
    'Intune configuration and security policies'   = @{ Nist = 'PR.PS-01, PR.PS-02 (software maintained), PR.PS-05'; Cis365 = 'Not covered (CIS Windows/Intune benchmarks)'; CisV8 = '4.1, 4.2, 10.1, 10.2'; Ffiec = 'CAT/ACET D3.PC.Im; D3.PC.De' }
    'Temporary Access Pass (TAP) implemented'      = @{ Nist = 'PR.AA-02 (identities proofed and bound), PR.AA-03; SP 800-63B account recovery'; Cis365 = '5.2.3 Authentication methods'; CisV8 = '6.1, 6.2 (access granting and revoking process)'; Ffiec = 'Authentication Guidance (2021): credential issuance and recovery' }
    'Users restricted from joining devices to Entra ID' = @{ Nist = 'ID.AM-01, PR.AA-05'; Cis365 = '5.1.2 Entra: device join and registration settings'; CisV8 = '1.1, 1.2'; Ffiec = 'CAT/ACET D3.PC.De' }
    'Automatic email forwarding outside the tenant' = @{ Nist = 'PR.DS-02 (data in transit), DE.CM-09; PR.PS-01'; Cis365 = '6.2.1 All forms of mail forwarding blocked'; CisV8 = '3.3, 9.x Email protections'; Ffiec = 'CAT/ACET D3.PC.Am; IT Handbook Information Security: Data loss prevention' }
    'User consent to applications'                 = @{ Nist = 'PR.AA-05, GV.SC (supply chain), PR.PS-05 (unauthorized software prevented)'; Cis365 = '5.1.5 User consent to applications; admin consent workflow'; CisV8 = '2.3, 2.5 (allowlist authorized software)'; Ffiec = 'CAT/ACET D4.RM Third-party risk; D3.PC.Am' }
    'SPF record (DNS)'                             = @{ Nist = 'PR.DS-02, PR.PS-01; NIST SP 800-177 (Trustworthy Email)'; Cis365 = '2.1.8 Ensure that SPF records are published for all Exchange domains'; CisV8 = '9.5 (implement DMARC; SPF and DKIM are prerequisites)'; Ffiec = 'CAT/ACET D3.PC.Im Infrastructure Management; IT Handbook Information Security: Email controls' }
    'DKIM signing (Exchange Online and DNS)'       = @{ Nist = 'PR.DS-02, PR.PS-01; NIST SP 800-177'; Cis365 = '2.1.9 Ensure that DKIM is enabled for all Exchange Online domains'; CisV8 = '9.5'; Ffiec = 'CAT/ACET D3.PC.Im; IT Handbook Information Security: Email controls' }
    'DMARC policy (DNS)'                           = @{ Nist = 'PR.DS-02, PR.PS-01, DE.CM-09; NIST SP 800-177'; Cis365 = '2.1.10 Ensure DMARC records for all Exchange Online domains are published'; CisV8 = '9.5 (implement DMARC)'; Ffiec = 'CAT/ACET D3.PC.Im; D3.DC.Th Threat Detection (DMARC reports)' }
    'Email protection (Defender for Office 365 / EOP)' = @{ Nist = 'PR.PS-05, DE.CM-09, PR.AT-01 (phishing)'; Cis365 = '2.1 Defender: Safe Links, Safe Attachments, anti-phishing, anti-spam, preset policies'; CisV8 = '9.6, 9.7, 10.1'; Ffiec = 'CAT/ACET D3.PC.Im, D3.DC.Th Threat Detection' }
    'Endpoint protection (Defender for Endpoint)'  = @{ Nist = 'DE.CM-09 (hardware/software monitored), PR.PS-05, RS.AN'; Cis365 = 'Not covered (Defender for Endpoint)'; CisV8 = '10.1, 10.7 (behaviour-based anti-malware), 13.2 (host-based intrusion detection)'; Ffiec = 'CAT/ACET D3.DC.Th, D3.DC.An Anomalous Activity Detection; D3.PC.De' }
    'Microsoft Secure Score'                       = @{ Nist = 'GV.OV-01 (cybersecurity risk management strategy outcomes reviewed), ID.IM-01 (improvements identified)'; Cis365 = 'Whole benchmark (Secure Score tracks many benchmark items)'; CisV8 = 'Implementation group progress'; Ffiec = 'CAT/ACET D1.G.Ov Oversight; D1.RM Risk Management: measurement and reporting' }
    'Privileged role assignments'                  = @{ Nist = 'PR.AA-05 (least privilege), GV.RR-02 (roles and responsibilities)'; Cis365 = '1.1.3 Between two and four Global Administrators; 1.1.1 Cloud-only admin accounts; 5.3 Privileged Identity Management and access reviews'; CisV8 = '5.4 (dedicated admin accounts), 6.8 (role-based access)'; Ffiec = 'CAT/ACET D3.PC.Am: privileged access; IT Handbook Information Security: Access rights administration' }
    'Global Administrators without MFA registered' = @{ Nist = 'PR.AA-03, PR.AA-05; SP 800-63B AAL2/AAL3 for privileged users'; Cis365 = '5.2.2 MFA for administrative roles'; CisV8 = '6.5'; Ffiec = 'Authentication Guidance (2021): privileged users; CAT/ACET D3.PC.Am' }
    'Too many Global Administrators'               = @{ Nist = 'PR.AA-05, GV.RR-02'; Cis365 = '1.1.3 Between two and four Global Administrators'; CisV8 = '5.4, 6.8'; Ffiec = 'CAT/ACET D3.PC.Am; IT Handbook: segregation of duties' }
    'Single Global Administrator'                  = @{ Nist = 'PR.AA-05, PR.IR-03 (resilience), RC'; Cis365 = '1.1.3 Between two and four Global Administrators'; CisV8 = '5.4'; Ffiec = 'CAT/ACET D5 Resilience; business continuity of administration' }
    'Indications of compromise to review'          = @{ Nist = 'DE.CM-01, DE.CM-03 (personnel activity monitored), DE.AE-02, DE.AE-06, RS.MA-01'; Cis365 = '6.1 Mailbox auditing; 5.2.2 sign-in and user risk policies; 3.x Purview audit'; CisV8 = '8.2, 8.5, 8.11 (collect, enrich, review audit logs), 17.x Incident response'; Ffiec = 'CAT/ACET D3.DC.An Anomalous Activity Detection, D3.DC.Ev Event Detection, D5.IR Incident Response; Authentication Guidance (2021): monitoring and logging' }
    'Teams guest and external access'              = @{ Nist = 'PR.AA-05, GV.SC-05 (third-party access), PR.DS-01'; Cis365 = '8.2 External access and unmanaged accounts; 8.5 Meeting policies (anonymous join, lobby); 5.1.6 Guest user settings; 1.2.1 Public groups'; CisV8 = '3.3 (data access control), 6.8'; Ffiec = 'CAT/ACET D3.PC.Am; D4.C Connections (third-party connectivity)' }
    'SharePoint and OneDrive external sharing'     = @{ Nist = 'PR.DS-01, PR.DS-02, PR.AA-05, PR.DS-10'; Cis365 = '7.2.3 External content sharing restricted; 7.2.5 Guests cannot re-share; 7.2.6 Sharing domain allow list; 7.2.7 Default link type; 7.3.2 OneDrive sync for unmanaged devices'; CisV8 = '3.3, 3.10, 6.8'; Ffiec = 'CAT/ACET D3.PC.Am; IT Handbook Information Security: Data classification and sharing' }
    'Audit logging (unified audit log)'            = @{ Nist = 'PR.PS-04 (log records generated and available), DE.CM-01, DE.CM-03, DE.CM-09, RS.AN-03'; Cis365 = '3.1.1 Ensure Microsoft 365 audit log search is enabled; 6.1.1 Mailbox auditing'; CisV8 = '8.2 (collect audit logs), 8.5 (detailed logs), 8.10 (retain logs)'; Ffiec = 'CAT/ACET D3.DC.Ev Event Detection, D2.MA Monitoring and Analysing; IT Handbook Information Security: Log management; Authentication Guidance (2021): logging and monitoring' }
    'Content open to anyone on the internet'       = @{ Nist = 'PR.DS-01, PR.DS-10, PR.AA-05'; Cis365 = '7.2.3 External sharing restricted; 7.2.7 Link sharing restricted'; CisV8 = '3.3 (data access control lists)'; Ffiec = 'CAT/ACET D3.PC.Am; IT Handbook Information Security: data classification; GLBA Safeguards' }
    'Content open to the whole organization'       = @{ Nist = 'PR.AA-05 (least privilege), PR.DS-01'; Cis365 = '7.2.7 Default link type; 1.2.1 Public groups'; CisV8 = '3.3, 6.8'; Ffiec = 'CAT/ACET D3.PC.Am: access rights administration' }
    'Usage report names are concealed'             = @{ Nist = 'ID.AM-01, ID.AM-03 (data flows and ownership), DE.CM-03'; Cis365 = 'n/a (reporting prerequisite)'; CisV8 = '1.1, 3.2 (data inventory)'; Ffiec = 'CAT/ACET D1.G.Ov Oversight: management information' }
    'Intune is not licensed'                       = @{ Nist = 'ID.AM-01, PR.PS-01'; Cis365 = 'n/a (licensing prerequisite)'; CisV8 = '1.1, 4.1'; Ffiec = 'CAT/ACET D3.PC.De' }
    'Entra ID P1 is not licensed'                  = @{ Nist = 'PR.AA-03, DE.CM-03'; Cis365 = 'n/a (licensing prerequisite for Conditional Access and sign-in logs)'; CisV8 = '6.3, 8.2'; Ffiec = 'Authentication Guidance (2021): risk-based controls' }
}
# Secure Score actions are mapped by keyword to the nearest framework expectations
$ScoreFrameworkRules = @(
    @{ Match = '(?i)multi-?factor|MFA|authentication strength|passwordless|FIDO|authenticator'; Nist = 'PR.AA-03; SP 800-63B'; Cis365 = '5.2.2 / 5.2.3'; CisV8 = '6.3, 6.5'; Ffiec = 'Authentication Guidance (2021); CAT/ACET D3.PC.Am' }
    @{ Match = '(?i)legacy|basic auth|SMTP|POP|IMAP';                     Nist = 'PR.AA-03, PR.PS-01'; Cis365 = '5.2.2 / 6.5'; CisV8 = '4.8, 6.3'; Ffiec = 'CAT/ACET D3.PC.Im' }
    @{ Match = '(?i)global admin|administrator|privileged|PIM|role';       Nist = 'PR.AA-05, GV.RR-02'; Cis365 = '1.1 / 5.3'; CisV8 = '5.4, 6.8'; Ffiec = 'CAT/ACET D3.PC.Am' }
    @{ Match = '(?i)forward|transport rule|mail flow';                     Nist = 'PR.DS-02, DE.CM-09'; Cis365 = '6.2'; CisV8 = '3.3'; Ffiec = 'CAT/ACET D3.PC.Am' }
    @{ Match = '(?i)SPF|DKIM|DMARC';                                       Nist = 'PR.DS-02; SP 800-177'; Cis365 = '2.1.8 / 2.1.9 / 2.1.10'; CisV8 = '9.5'; Ffiec = 'CAT/ACET D3.PC.Im' }
    @{ Match = '(?i)safe links|safe attachments|anti-?phish|spam|malware|impersonat|spoof|quarantine|zero-hour|ZAP'; Nist = 'PR.PS-05, DE.CM-09'; Cis365 = '2.1'; CisV8 = '9.6, 9.7, 10.1'; Ffiec = 'CAT/ACET D3.PC.Im, D3.DC.Th' }
    @{ Match = '(?i)consent|third[- ]party app|application permission|OAuth'; Nist = 'PR.AA-05, PR.PS-05, GV.SC'; Cis365 = '5.1.5'; CisV8 = '2.3, 2.5'; Ffiec = 'CAT/ACET D4.RM' }
    @{ Match = '(?i)audit|log';                                            Nist = 'DE.CM-03, PR.PS-04'; Cis365 = '6.1 / 3.x'; CisV8 = '8.2, 8.5'; Ffiec = 'CAT/ACET D3.DC.Ev' }
    @{ Match = '(?i)risk polic|risky|identity protection|sign-in risk|user risk'; Nist = 'DE.CM-03, DE.AE-02'; Cis365 = '5.2.2 risk policies'; CisV8 = '8.11, 13.x'; Ffiec = 'CAT/ACET D3.DC.An' }
    @{ Match = '(?i)password';                                             Nist = 'PR.AA-01; SP 800-63B password guidance'; Cis365 = '1.3.1 / 5.2.3 password protection'; CisV8 = '5.2'; Ffiec = 'Authentication Guidance (2021)' }
    @{ Match = '(?i)sharing|SharePoint|OneDrive|guest|external user|Teams|anonymous';   Nist = 'PR.DS-01, PR.AA-05'; Cis365 = '7.2 / 8.2 / 8.5 / 5.1.6'; CisV8 = '3.3, 6.8'; Ffiec = 'CAT/ACET D3.PC.Am' }
    @{ Match = '(?i)device|Intune|compliance|encrypt|BitLocker|Defender for Endpoint|endpoint|attack surface|tamper'; Nist = 'PR.PS-01, PR.DS-01, DE.CM-09'; Cis365 = 'CIS Intune / Windows benchmarks'; CisV8 = '4.1, 3.6, 10.x'; Ffiec = 'CAT/ACET D3.PC.De' }
    @{ Match = '(?i)DLP|data loss|sensitivity|label|retention';           Nist = 'PR.DS-01, PR.DS-10, GV.PO'; Cis365 = '3.x Purview'; CisV8 = '3.1, 3.2, 3.13'; Ffiec = 'CAT/ACET D3.PC.Am; GLBA Safeguards' }
)
function Get-FrameworkText { param([string]$Finding)
    $fw = $FrameworkMap[$Finding]
    if (-not $fw) { foreach ($rule in $ScoreFrameworkRules) { if ($Finding -match $rule.Match) { $fw = $rule; break } } }
    if (-not $fw) { return '' }
    return "NIST CSF 2.0: $($fw.Nist) | CIS M365 Benchmark: $($fw.Cis365) | CIS Controls v8: $($fw.CisV8) | FFIEC/NCUA: $($fw.Ffiec)"
}

# ----------------------------------------------------------------------------------------------
# Playbooks: who does it, how long it takes, the exact steps, and how the next report proves it.
# Written for the IT provider or MSP who will receive the ticket, in the order the clicks happen.
# ----------------------------------------------------------------------------------------------
$Playbooks = @{
    'MFA enforced for sign-in' = @{ Owner = 'IT provider / MSP with Global Administrator'; Effort = 'Half a day plus a one-week pilot'; Prereq = 'Entra ID P1 (included in Business Premium, E3, E5) for Conditional Access; otherwise turn on security defaults'; Steps = @(
        'Confirm every user has an MFA method registered (section Users, filter No MFA) and fix the gaps first, or the policy will lock them out.',
        'Entra admin center > Protection > Conditional Access > New policy. Name it "Require MFA for all users".',
        'Users: All users. Exclude one break-glass account (cloud-only, long random password, stored offline) and the Entra Connect sync account if present.',
        'Target resources: All cloud apps. Grant: Require multifactor authentication (or Require authentication strength: Phishing-resistant MFA for administrators).',
        'Enable policy: Report-only for one week, review the Insights workbook for users who would be blocked, then switch to On.',
        'If the tenant has no Entra ID P1: Entra admin center > Identity > Overview > Properties > Manage security defaults > Enabled.'); Verify = 'Section 3 shows "MFA enforced for sign-in: Enforced" and section 8 lists the policy as On covering All users / All cloud apps.' }
    'MFA registered by all users' = @{ Owner = 'Security officer with help desk'; Effort = 'One to two weeks of user follow-up'; Prereq = 'None'; Steps = @(
        'Export the Users table with the No MFA filter; disable or delete accounts that are no longer needed (that alone usually clears half the list).',
        'Entra admin center > Protection > Authentication methods > Registration campaign: nudge users to set up Microsoft Authenticator at sign-in; set the snooze limit to three days.',
        'For remaining users issue a Temporary Access Pass and walk them through Authenticator registration (phone call, not email, for executives).',
        'Add a Conditional Access policy "Secure security info registration" that requires a compliant device or trusted location to register methods.'); Verify = 'Section Users shows zero enabled members without MFA; section 3 row "MFA registered by all users" turns to Yes.' }
    'Phishing-resistant MFA (passkeys) rollout' = @{ Owner = 'IT provider / MSP; security officer sponsors'; Effort = 'Pilot in one month, full rollout in a quarter'; Prereq = 'Microsoft Authenticator app or FIDO2 security keys; Windows Hello for Business on managed PCs'; Steps = @(
        'Entra admin center > Protection > Authentication methods > Policies > Passkey (FIDO2): Enable, Target all users, Allow self-service set up, and allow the Microsoft Authenticator passkey AAGUIDs (or leave key restrictions off to allow any passkey).',
        'Start with administrators: each registers a passkey in Authenticator (My Security Info > Add method > Passkey) or a hardware key; keep a second key per admin in a safe.',
        'Create a Conditional Access policy "Phishing-resistant MFA for administrators": Users = directory roles (Global Administrator, Security Administrator, Exchange Administrator, SharePoint Administrator, Conditional Access Administrator, Privileged Role Administrator); Grant = Require authentication strength: Phishing-resistant MFA.',
        'Roll out to finance and wire staff next (they are the targets of payment fraud), then everyone, with Windows Hello for Business on Intune-managed PCs so the device itself is the passkey.',
        'When registration passes 90%, disable SMS and voice as MFA methods in the authentication methods policy.'); Verify = 'Section 9 "Passkey and passwordless rollout" shows the percentage climbing each month; section 3 row "Phishing-resistant MFA for administrators" reaches Complete.' }
    'Phishing-resistant MFA for administrators' = @{ Owner = 'IT provider / MSP'; Effort = 'One day'; Prereq = 'Passkey (FIDO2) method enabled'; Steps = @(
        'Each administrator registers a passkey (Microsoft Authenticator passkey, or a FIDO2 key such as YubiKey) at https://mysignins.microsoft.com/security-info.',
        'Keep a backup key per administrator and a break-glass account excluded from the policy.',
        'Conditional Access > New policy: Users = directory roles (all administrator roles); Target = All cloud apps; Grant = Require authentication strength: Phishing-resistant MFA; enable.'); Verify = 'Section 3 row shows Complete (N of N administrators) and section 13 lists each admin with a passkey or Windows Hello method.' }
    'Legacy authentication protocols disabled' = @{ Owner = 'IT provider / MSP'; Effort = 'Half a day plus a two-week report-only period'; Prereq = 'Entra ID P1 for Conditional Access; Exchange Online for SMTP AUTH'; Steps = @(
        'Entra admin center > Monitoring > Sign-in logs: filter Client app = Exchange ActiveSync, Other clients, IMAP, POP, SMTP for the last 30 days to see who still uses legacy protocols (scanners, copiers, line-of-business apps).',
        'Move those senders to modern authentication or to a connector/relay (Exchange admin center > Mail flow > Connectors) so they no longer need a mailbox password.',
        'Conditional Access > New policy "Block legacy authentication": Users = All users; Target = All cloud apps; Conditions > Client apps = Exchange ActiveSync clients and Other clients; Grant = Block. Report-only two weeks, then On.',
        'Exchange admin center > Settings > Mail flow: turn off SMTP AUTH organization-wide; enable it per mailbox only for the relay accounts that truly need it.',
        'Disable POP and IMAP on every mailbox that does not need them (Exchange admin center > Mailboxes > Manage email apps).'); Verify = 'Section 3 shows Blocked; section 14 shows SMTP AUTH disabled tenant-wide and zero POP/IMAP mailboxes; section 4 shows no successful legacy sign-ins.' }
    'Automatic email forwarding outside the tenant' = @{ Owner = 'IT provider / MSP (Exchange Administrator)'; Effort = 'One hour'; Prereq = 'Exchange Online'; Steps = @(
        'Defender portal > Email & collaboration > Policies & rules > Threat policies > Anti-spam > Anti-spam outbound policy (Default): Automatic forwarding rules = Off (or On with an approved list).',
        'Exchange admin center > Mail flow > Remote domains > Default: Allow automatic forwarding = Off.',
        'Review the mailboxes and inbox rules listed in section 14 with their owners; remove forwarding that is not approved in writing.',
        'Optionally add a transport rule that rejects auto-forwarded messages to external recipients with an explanatory NDR.'); Verify = 'Section 3 row shows Blocked and the forwarding tables in section 14 are empty.' }
    'Conditional Access restricts access to Entra joined or compliant devices' = @{ Owner = 'IT provider / MSP'; Effort = 'One to three months (depends on Intune enrolment)'; Prereq = 'Intune licensed and devices enrolled; Entra ID P1'; Steps = @(
        'Finish Intune enrolment and a compliance policy for every platform first (see the Intune findings).',
        'Conditional Access > New policy "Require compliant or Entra joined device": Users = All users (exclude break-glass); Target = Office 365; Grant = Require device to be marked as compliant OR Require Microsoft Entra hybrid joined device.',
        'Add a second policy for unmanaged devices that allows browser access with no download (Session: Use app enforced restrictions) if staff must use personal devices.',
        'Report-only for two weeks; review which users would be blocked; then On.'); Verify = 'Section 3 shows Enforced and section 7 shows devices as compliant.' }
    'Intune deployed to Windows computers' = @{ Owner = 'IT provider / MSP'; Effort = 'One to two months for a small fleet'; Prereq = 'Intune licence (Business Premium, E3/E5, or Intune Plan 1)'; Steps = @(
        'Intune admin center > Devices > Enrollment > Windows > Automatic Enrollment: MDM user scope = All.',
        'Entra admin center > Devices > Device settings: Users may join devices to Entra ID = Selected (IT staff) and require MFA to join.',
        'For domain-joined PCs configure hybrid Entra join through Entra Connect and a Group Policy that enables automatic MDM enrolment; for new PCs use Autopilot.',
        'Enrol the devices listed under section 7 with the "Not in Intune" filter, replacing any that cannot run a supported Windows version.'); Verify = 'Section 7 shows every enabled Windows device as Enrolled and section 3 row turns to All enrolled.' }
    'Intune compliance policies for mobile devices' = @{ Owner = 'IT provider / MSP'; Effort = 'Half a day plus user enrolment'; Prereq = 'Intune (Business Premium, E3/E5 or Intune Plan 1); Company Portal app on phones'; Steps = @(
        'Intune admin center > Devices > Compliance > Create policy > iOS/iPadOS: require passcode (6 digits or more, not simple), block jailbroken devices, minimum OS = current major release minus one, mark noncompliant after 1 day.',
        'Create policy > Android Enterprise (personally-owned work profile): require work-profile password, encryption (always on), block rooted devices, minimum OS version, Play Protect verified.',
        'Assign both to All users (or the licensed users group). Set the noncompliance action to notify the user and, after 7 days, retire the work data.',
        'Conditional Access > New policy "Require compliant mobile device": Users = All users; Target = Office 365; Conditions > Device platforms = iOS, Android; Grant = Require device to be marked as compliant (or Require approved client app). Report-only first, then On.',
        'Publish a short enrolment guide (Company Portal, Outlook) and give staff two weeks before enforcement.'); Verify = 'Section 6 lists an iOS/iPadOS and an Android compliance policy with assignments; section 3 row shows the count and the Critical finding disappears.' }
    'Intune compliance policies' = @{ Owner = 'IT provider / MSP'; Effort = 'One day'; Prereq = 'Intune'; Steps = @(
        'Intune admin center > Devices > Compliance > Create policy, one per platform.',
        'Windows: require BitLocker, Secure Boot, TPM, Defender Antivirus on and up to date, firewall on, minimum OS version, password/PIN; mark noncompliant after 1 day.',
        'iOS/iPadOS and Android: require passcode, minimum OS version, no jailbreak/root, Defender for Endpoint risk level at or below Medium if licensed.',
        'Assign to All devices (or All users) and set the action for noncompliance to email the user and mark noncompliant after a grace period.'); Verify = 'Section 6 lists one compliance policy per platform with assignments; section 7 shows compliance per device.' }
    'Intune configuration and security policies' = @{ Owner = 'IT provider / MSP'; Effort = 'Two to three days'; Prereq = 'Intune'; Steps = @(
        'Intune admin center > Endpoint security > Security baselines: deploy the Microsoft Security Baseline for Windows and the Defender for Endpoint baseline to a pilot group, then all devices.',
        'Endpoint security > Disk encryption: BitLocker policy (XTS-AES 256, require startup PIN or TPM, escrow keys to Entra ID).',
        'Endpoint security > Antivirus and Firewall policies; Attack surface reduction rules in block mode.',
        'Devices > Configuration: screen lock timeout, Windows Update rings, OneDrive Known Folder Move, and disable local administrator for standard users (LAPS for the local admin).'); Verify = 'Section 6 lists the policies with assignments; section 7 shows devices encrypted.' }
    'User consent to applications' = @{ Owner = 'IT provider / MSP'; Effort = 'One hour'; Prereq = 'None'; Steps = @(
        'Entra admin center > Identity > Applications > Enterprise applications > Consent and permissions > User consent settings: Do not allow user consent (or Allow for verified publishers, selected permissions).',
        'Admin consent settings: Users can request admin consent = Yes; add the security officer and IT provider as reviewers; set expiry to 30 days.',
        'Review the applications in section 10 with user-consented permissions and revoke any that are not business-justified.'); Verify = 'Section 3 row shows Restricted and section 11 shows only admin-approved consent events.' }
    'Temporary Access Pass (TAP) implemented' = @{ Owner = 'IT provider / MSP'; Effort = 'Thirty minutes'; Prereq = 'None'; Steps = @(
        'Entra admin center > Protection > Authentication methods > Policies > Temporary Access Pass: Enable, Target = All users (or help desk group).',
        'Settings: one-time use = Yes, default lifetime 60 minutes, maximum 8 hours, length 8 characters.',
        'Write the help-desk procedure: verify identity by callback to a known number before issuing a TAP; never read it out over an inbound call.'); Verify = 'Section 9 shows TAP Enabled with targets; section 3 row turns to Enabled.' }
    'Users restricted from joining devices to Entra ID' = @{ Owner = 'IT provider / MSP'; Effort = 'Fifteen minutes'; Prereq = 'None'; Steps = @(
        'Entra admin center > Identity > Devices > Overview > Device settings.',
        'Users may join devices to Microsoft Entra = Selected (IT staff / enrolment group); Users may register their devices = Selected or None if Intune enrolment uses another path; Require Multifactor Authentication to register or join = Yes.',
        'Maximum number of devices per user = 5 or fewer.'); Verify = 'Section 9 Device registration settings shows the restriction; section 3 row shows Restricted.' }
    'SPF record (DNS)' = @{ Owner = 'IT provider / MSP with access to the public DNS zone'; Effort = 'Thirty minutes per domain plus DNS propagation'; Prereq = 'Login to the DNS host (registrar, Cloudflare, GoDaddy, the core provider) for each domain'; Steps = @(
        'List every system that legitimately sends mail as the domain: Exchange Online, the core banking or statement provider, marketing tools, scanners and alerting systems (check the Exchange message trace and current SPF record).',
        'Build one TXT record at the domain root: v=spf1 include:spf.protection.outlook.com <one include or ip4 per other sender> -all. Keep the total DNS lookups under ten (flatten or remove unused includes).',
        'Publish it in the DNS zone, replacing any existing SPF record (there must be exactly one). Use -all (hard fail); ~all lets spoofed mail through to the inbox.',
        'For domains that never send mail (parked domains), publish v=spf1 -all.',
        'Check with https://mxtoolbox.com/spf.aspx or the next run of this report.'); Verify = 'Section 15 shows SPF "Hard fail (-all)" for every domain and the control reads Complete.' }
    'DKIM signing (Exchange Online and DNS)' = @{ Owner = 'IT provider / MSP (Exchange Administrator plus DNS access)'; Effort = 'Twenty minutes per domain plus DNS propagation'; Prereq = 'Exchange Online; access to the public DNS zone'; Steps = @(
        'Open the DKIM page in the Defender portal (link in the Where column): Email & collaboration > Policies & rules > Threat policies > Email authentication settings > DKIM.',
        'Select the domain. If no keys exist, click Create DKIM keys. The page shows two CNAME records: selector1._domainkey.<domain> and selector2._domainkey.<domain> with their target values (selector1-<domain-with-dashes>._domainkey.<tenant>.onmicrosoft.com).',
        'Create both CNAME records in the public DNS zone exactly as shown (host = selector1._domainkey, value = the target; same for selector2). Wait for propagation (minutes to a few hours).',
        'Back on the DKIM page, turn the "Sign messages for this domain with DKIM signatures" toggle to Enabled. If it reports the CNAMEs are missing, wait and retry.',
        'Repeat for every custom domain that sends mail. Rotate keys from the same page once a year.'); Verify = 'Section 15 shows DKIM signing Enabled and "Both selectors published" for every domain; the control reads Complete.' }
    'DMARC policy (DNS)' = @{ Owner = 'IT provider / MSP with access to the public DNS zone; security officer receives the reports'; Effort = 'Fifteen minutes to publish, then two to four weeks of monitoring before enforcement'; Prereq = 'SPF and DKIM in place for the domain; a mailbox or DMARC reporting service for aggregate reports'; Steps = @(
        'Publish a TXT record at _dmarc.<domain>: v=DMARC1; p=none; rua=mailto:dmarc@<domain>; pct=100. p=none only monitors; it changes nothing about delivery yet.',
        'Collect the aggregate (rua) reports for two to four weeks (a free or paid DMARC report service makes them readable) and confirm every legitimate sender passes SPF or DKIM.',
        'Change the record to p=quarantine; pct=100, wait another two weeks, then p=reject; pct=100. Reject is the state examiners and CIS expect.',
        'Add adkim=s and aspf=s (strict alignment) once all senders align, and publish p=reject on parked domains immediately.',
        'Keep the rua address monitored; a new sender failing DMARC is either a new vendor to add to SPF or an attacker spoofing the institution.'); Verify = 'Section 15 shows DMARC p=reject (or quarantine) at pct 100 with reporting Yes for every domain; the control reads Complete.' }
    'Email protection (Defender for Office 365 / EOP)' = @{ Owner = 'IT provider / MSP'; Effort = 'Two hours'; Prereq = 'Defender for Office 365 Plan 1 (included in Business Premium) for Safe Links / Safe Attachments'; Steps = @(
        'Defender portal > Email & collaboration > Policies & rules > Threat policies > Preset security policies: turn on Standard protection for all users (Strict for executives and finance). This configures anti-spam, anti-phishing, Safe Links and Safe Attachments consistently.',
        'Anti-phishing policy: enable impersonation protection for the CEO, CFO, controller and the institution''s domains; mailbox intelligence = On; phishing threshold = 2 (Aggressive).',
        'Anti-spam: set high-confidence spam and phishing actions to Quarantine, bulk threshold 6 or lower; keep zero-hour auto purge on.',
        'Safe Attachments for SharePoint, OneDrive and Teams and Safe Documents: On (Defender portal > Policies > Safe Attachments > Global settings).',
        'Turn on the user-reported message settings so staff can report phishing to Microsoft and to the security officer.'); Verify = 'Section 15 shows presets enabled, Quarantine actions, impersonation protection on; section 3 row shows Configured.' }
    'Endpoint protection (Defender for Endpoint)' = @{ Owner = 'IT provider / MSP'; Effort = 'One week'; Prereq = 'Defender for Business (Business Premium) or Defender for Endpoint P1/P2'; Steps = @(
        'Intune admin center > Endpoint security > Microsoft Defender for Endpoint: connect Intune to Defender (Open the Defender portal > Settings > Endpoints > Advanced features > Microsoft Intune connection = On).',
        'Endpoint security > Endpoint detection and response > Create policy: onboard via the connector, assign to all Windows devices; do the same for macOS.',
        'Enable tamper protection and automated investigation and remediation (Full) in Defender portal > Settings > Endpoints.',
        'Set Defender vulnerability management to feed Intune compliance (Compliance policy: machine risk score at or below Medium).'); Verify = 'Section 15 shows every managed device onboarded with an active sensor; section 3 row shows Onboarded.' }
    'Microsoft Secure Score' = @{ Owner = 'Security officer with IT provider / MSP'; Effort = 'Ongoing; one review per month'; Prereq = 'None'; Steps = @(
        'Work the Remediation Plan above from the top; each Secure Score action lists its points.',
        'In the Defender portal > Secure Score > Recommended actions, mark actions that are handled by a third-party product as Resolved through third party so the score reflects reality.',
        'Review the score and the trend chart in this report at each monthly review.'); Verify = 'The Secure Score KPI reaches the target percentage and the trend chart rises month over month.' }
    'Global Administrators without MFA registered' = @{ Owner = 'Security officer; the administrator concerned'; Effort = 'Today'; Prereq = 'None'; Steps = @(
        'Have each listed administrator register Microsoft Authenticator (or a passkey) now at https://mysignins.microsoft.com/security-info.',
        'If an account cannot register (service account, shared account), remove Global Administrator from it today and replace it with a dedicated, cloud-only admin account.',
        'Create the Conditional Access policy requiring MFA for administrator roles if it does not exist.'); Verify = 'Section 13 shows every Global Administrator with MFA registered; the finding disappears from the next report.' }
    'Too many Global Administrators' = @{ Owner = 'Security officer with IT provider / MSP'; Effort = 'Half a day'; Prereq = 'None (PIM needs Entra ID P2)'; Steps = @(
        'List the accounts in section 13 and decide which two to four people truly need Global Administrator; include one break-glass account.',
        'Give everyone else a least-privilege role (Exchange Administrator, User Administrator, Helpdesk Administrator, Security Reader, Global Reader).',
        'If Entra ID P2 is licensed, make the remaining assignments Eligible in Privileged Identity Management with approval and MFA on activation.'); Verify = 'Section 1 KPI shows two to four Global Administrators; the finding disappears.' }
    'Single Global Administrator' = @{ Owner = 'Security officer with IT provider / MSP'; Effort = 'One hour'; Prereq = 'None'; Steps = @(
        'Create a second, cloud-only break-glass account (no licence needed beyond the role) with a 25+ character random password stored in a sealed envelope or password vault with dual control.',
        'Register MFA for it (FIDO2 key in the safe is ideal) and exclude it from Conditional Access policies.',
        'Document the procedure for using it and test it twice a year.'); Verify = 'Section 1 KPI shows two Global Administrators.' }
    'Indications of compromise to review' = @{ Owner = 'Security officer with IT provider / MSP'; Effort = 'Same day for urgent items; one week for the rest'; Prereq = 'None'; Steps = @(
        'Take each urgent row in section 2 to the person named (mailbox owner, administrator) and ask one question: did you do this?',
        'If the answer is no or uncertain: reset the password, revoke sessions (Entra admin center > Users > user > Revoke sessions), remove the rule, credential or role, and preserve the evidence (export the audit rows).',
        'Check the affected mailbox for sent items, forwarding and other rules; check the user''s devices in Defender.',
        'Record the outcome (expected / incident) next to each item; anything confirmed as an incident goes through the incident response plan and, where required, regulatory notification.',
        'Match the "to review" rows to change tickets or help-desk records within the week.'); Verify = 'Next month''s section 4 shows only items that match change records; the urgent count is zero.' }
    'Teams guest and external access' = @{ Owner = 'IT provider / MSP (Teams Administrator); team owners'; Effort = 'Half a day plus owner follow-up'; Prereq = 'None'; Steps = @(
        'Entra admin center > External Identities > External collaboration settings: Guest invite settings = Only users assigned to specific admin roles can invite guests (or require approval through an access package); Collaboration restrictions = Allow invitations only to specified domains (your partners, auditors, core provider).',
        'Teams admin center > Users > External access: Allow only specific external domains; Teams accounts not managed by an organization = Off; Skype users = Off.',
        'Teams admin center > Meetings > Meeting policies > Global: Anonymous users can join = Off (or lobby for everyone), Who can bypass the lobby = People in my org, Dial-in users bypass = Off.',
        'Send each team owner the list of guests in their team (section 16) and have them remove anyone not currently working with the institution; set a Microsoft 365 group expiration policy.',
        'Set teams holding confidential material to Private.'); Verify = 'Section 16 shows guest invitations restricted, external access on an allow list, anonymous join off, and fewer teams with guests.' }
    'SharePoint and OneDrive external sharing' = @{ Owner = 'IT provider / MSP (SharePoint Administrator)'; Effort = 'Two hours'; Prereq = 'None'; Steps = @(
        'SharePoint admin center > Policies > Sharing: set SharePoint and OneDrive to "New and existing guests" (never "Anyone" for a financial institution; "Existing guests" if partners are onboarded through Entra).',
        'More external sharing settings: Limit external sharing by domain (allow list); Guests must sign in using the same account to which sharing invitations are sent; Allow guests to share items they do not own = Off; Guest access expires after 30 days.',
        'File and folder links: Default link type = Specific people; Default permission = View.',
        'Access control: Unmanaged devices = Allow limited, web-only access; Idle session sign-out = On (warn 55 min, sign out 60); Apps that do not use modern authentication = Block.',
        'OneDrive: Sync allowed only on domain-joined or compliant PCs; set deleted-user OneDrive retention to the minimum policy allows and assign a manager for departed staff.',
        'Review the users in section 17 who shared externally and the inactive or ownerless sites; archive or delete what is no longer needed.'); Verify = 'Section 17 shows the sharing level as guests only (no Anyone links), domain restriction on, idle sign-out on, unmanaged sync restricted.' }
    'Audit logging (unified audit log)' = @{ Owner = 'IT provider / MSP with Global Administrator or Compliance Administrator'; Effort = 'Ten minutes; recording starts within hours'; Prereq = 'None (included in every Microsoft 365 business plan)'; Steps = @(
        'Open the Microsoft Purview portal (purview.microsoft.com) > Solutions > Audit. If the banner "Start recording user and admin activity" appears, click it.',
        'Alternatively in Exchange Online PowerShell: Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true.',
        'Confirm mailbox auditing is on as well: Get-OrganizationConfig | Select AuditDisabled should be False.',
        'Set retention: Audit (Standard) keeps 180 days; with Microsoft 365 E5 or the Audit Premium add-on create a retention policy for one year. Examiners commonly ask for 12 months.',
        'Re-run this report after 24 hours; the Review These Now section will then show sharing, download and mailbox-rule activity.'); Verify = 'Section 3 row "Audit logging (unified audit log)" reads Enabled and the audit-log rows in section 2 stop saying Not checked.' }
    'Content open to anyone on the internet' = @{ Owner = 'Content owners with IT provider / MSP'; Effort = 'One hour plus owner follow-up'; Prereq = 'None'; Steps = @(
        'Section 19, filter Anyone (anonymous): open each item with its link, choose Manage access, and remove the Anyone link (or set an expiry if it must stay for a few days).',
        'For calendars listed with Anonymous permission: Set-MailboxFolderPermission -Identity "user:\Calendar" -User Anonymous -AccessRights None.',
        'SharePoint admin center > Policies > Sharing: set SharePoint and OneDrive to New and existing guests (never Anyone) and set Anyone links to expire if they must remain for a specific site.',
        'Re-run the report; the count must be zero.'); Verify = 'Section 19 shows zero items open to anyone and the finding disappears.' }
    'Content open to the whole organization' = @{ Owner = 'Content owners with IT provider / MSP'; Effort = 'Half a day'; Prereq = 'None'; Steps = @(
        'Section 19, filter Whole organization: for each item decide whether every employee and guest should read it; if not, Manage access and replace the organization link with specific people or a team.',
        'Teams admin center > Teams: set teams holding board, HR, lending or customer material to Private.',
        'SharePoint admin center > Sharing > File and folder links: Default link type = Specific people; Default permission = View.',
        'Exchange: set calendar Default permissions above free/busy back to AvailabilityOnly (Set-MailboxFolderPermission -User Default -AccessRights AvailabilityOnly).'); Verify = 'Section 19 shows only intended organization-wide items (for example the staff handbook site).' }
    'Usage report names are concealed' = @{ Owner = 'IT provider / MSP with Global Administrator'; Effort = 'Five minutes'; Prereq = 'Global Administrator'; Steps = @(
        'Microsoft 365 admin center (admin.microsoft.com) > Settings > Org settings > Services tab > Reports.',
        'Clear "Display concealed user, group, and site names in all reports" and Save.',
        'Run this report again (the usage reports refresh within 48 hours); OneDrive owners, site owners and team names then appear.'); Verify = 'Section 18 shows OneDrive owners and site owners by name and the finding disappears.' }
    'Intune is not licensed' = @{ Owner = 'IT committee (licensing decision)'; Effort = 'Decision plus purchase'; Prereq = 'Budget'; Steps = @(
        'Compare Microsoft 365 Business Premium (includes Intune, Entra ID P1, Defender for Business and Defender for Office 365 P1) against adding Intune Plan 1 and Entra ID P1 to the current licence.',
        'For a bank or credit union with managed PCs, Business Premium (up to 300 users) or E3/E5 is almost always the right answer; present the cost per user per month to the committee.',
        'After purchase, follow the Intune playbooks above.'); Verify = 'Section 1 licensed capabilities shows Intune; the licence finding disappears.' }
    'Entra ID P1 is not licensed' = @{ Owner = 'IT committee (licensing decision)'; Effort = 'Decision plus purchase'; Prereq = 'Budget'; Steps = @(
        'Entra ID P1 provides Conditional Access, the sign-in log, MFA registration reporting and risk-based controls that examiners expect; it is included in Business Premium, E3 and E5.',
        'Until it is licensed, enable security defaults so MFA is enforced for everyone.',
        'After purchase, implement the Conditional Access playbooks (MFA, legacy authentication, device access).'); Verify = 'Section 1 licensed capabilities shows Entra ID P1; the licence finding disappears.' }
}
# Secure Score actions: matched by title keywords to a playbook so Microsoft's text is backed by concrete steps
$ScorePlaybooks = @(
    @{ Match = '(?i)MFA.*admin|administrative roles.*MFA|multifactor.*admin';          Use = 'Phishing-resistant MFA for administrators' }
    @{ Match = '(?i)MFA|multi-?factor|authentication strength';                          Use = 'MFA enforced for sign-in' }
    @{ Match = '(?i)legacy|basic auth|SMTP AUTH|POP|IMAP';                               Use = 'Legacy authentication protocols disabled' }
    @{ Match = '(?i)forward';                                                            Use = 'Automatic email forwarding outside the tenant' }
    @{ Match = '(?i)\bSPF\b';                                                           Use = 'SPF record (DNS)' }
    @{ Match = '(?i)\bDKIM\b';                                                          Use = 'DKIM signing (Exchange Online and DNS)' }
    @{ Match = '(?i)\bDMARC\b';                                                         Use = 'DMARC policy (DNS)' }
    @{ Match = '(?i)safe links|safe attachments|anti-?phish|impersonat|spam|malware|quarantine|zero-hour|ZAP|spoof|mailbox intelligence|phishing'; Use = 'Email protection (Defender for Office 365 / EOP)' }
    @{ Match = '(?i)consent|third[- ]party app';                                         Use = 'User consent to applications' }
    @{ Match = '(?i)global admin';                                                       Use = 'Too many Global Administrators' }
    @{ Match = '(?i)sharing|SharePoint|OneDrive|anonymous link|external user';           Use = 'SharePoint and OneDrive external sharing' }
    @{ Match = '(?i)Teams|guest|external access|meeting';                                Use = 'Teams guest and external access' }
    @{ Match = '(?i)compliance polic|Intune|device';                                     Use = 'Intune compliance policies' }
    @{ Match = '(?i)Defender for Endpoint|EDR|tamper|attack surface|antivirus';          Use = 'Endpoint protection (Defender for Endpoint)' }
    @{ Match = '(?i)passwordless|FIDO|passkey|phishing-resistant';                       Use = 'Phishing-resistant MFA (passkeys) rollout' }
)
$GenericScorePlaybook = @{ Owner = 'IT provider / MSP'; Effort = 'See Microsoft''s estimate in the row'; Prereq = 'See the licence tier in the row'; Steps = @(
    'Open the action with the link in the Detail column; the Defender portal page shows Microsoft''s implementation steps and the exact setting.',
    'Make the change in a pilot scope first where the action allows it, then for all users.',
    'Return to Secure Score > Recommended actions and confirm the action moves to Completed (it can take 24 to 48 hours to update).'); Verify = 'The action no longer appears in the Remediation Plan and the Secure Score KPI rises by the points shown.' }
function Get-Playbook { param([string]$Finding, [string]$Section)
    if ($Playbooks.ContainsKey($Finding)) { return $Playbooks[$Finding] }
    if ($Section -eq 'Secure Score') {
        foreach ($rule in $ScorePlaybooks) { if ($Finding -match $rule.Match -and $Playbooks.ContainsKey($rule.Use)) { return $Playbooks[$rule.Use] } }
        return $GenericScorePlaybook
    }
    return $null
}
$PhaseByPriority = @{ 'Critical' = 'Now: this week'; 'High' = 'Next 30 days'; 'Medium' = 'Next 90 days'; 'Low' = 'Next 12 months and ongoing' }
$PhaseOrder = @('Now: this week', 'Next 30 days', 'Next 90 days', 'Next 12 months and ongoing')

$remediationGuides = @(
    @{ Control = 'Audit logging (unified audit log)';             Bad = '^Disabled';               Priority = 'Critical'; Action = 'Turn on auditing today: Microsoft Purview portal > Audit > "Start recording user and admin activity" (or Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $true in Exchange Online PowerShell). Recording begins within a few hours and nothing before that can be recovered; every investigation and most examiner requests depend on it.'; Url = 'https://purview.microsoft.com/audit/auditsearch' }
    @{ Control = 'Indications of compromise to review';          Bad = '^\d+ urgent';              Priority = 'Critical'; Action = 'Today: take every item marked urgent in the Indications of Compromise section to the person named and confirm it was expected. Anything unexplained is a security incident: reset the password, revoke sessions, remove the rule, credential or setting, and engage your IT provider or incident response contact.'; Url = 'https://security.microsoft.com/incidents' }
    @{ Control = 'Indications of compromise to review';          Bad = '^\d+ to review';           Priority = 'High';     Action = 'Within the week, match every item in the Indications of Compromise section (role assignments, MFA changes, password resets, policy changes, new apps, new accounts, delegates, multi-country sign-ins) to a change ticket or a confirmation from the person involved. Escalate anything nobody can account for.'; Url = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuditLogsMenuBlade/~/Audit' }
    @{ Control = 'SharePoint and OneDrive external sharing';     Bad = '^Anyone links';            Priority = 'High';     Action = 'In the SharePoint admin center set external sharing for SharePoint and OneDrive to "New and existing guests" (or stricter), expire existing Anyone links, set the default link type to "People you choose", and restrict sharing to approved partner domains.'; Url = 'https://admin.microsoft.com/sharepoint?page=sharing&modern=true' }
    @{ Control = 'SharePoint and OneDrive external sharing';     Bad = '^Allowed';                 Priority = 'Low';      Action = 'Limit external sharing to an allow list of partner domains, require guests to sign in with the invited account, turn off re-sharing by guests, set guest access to expire, and review the externally shared files list monthly.'; Url = 'https://admin.microsoft.com/sharepoint?page=sharing&modern=true' }
    @{ Control = 'Teams guest and external access';              Bad = '^(Allowed|Guests present)'; Priority = 'Critical';     Action = 'Decide who may invite guests (restrict to administrators or require approval), have each team owner confirm current guests, restrict Teams external access to approved partner domains, block Teams consumer accounts, and make anonymous meeting participants wait in the lobby.'; Url = 'https://admin.teams.microsoft.com/company-wide-settings/external-communications' }
    @{ Control = 'MFA enforced for sign-in';                      Bad = '^(Not enforced|Partial)'; Priority = 'Critical'; Action = 'Create an enabled Conditional Access policy that requires MFA (or a phishing-resistant authentication strength) for all users and all cloud apps, excluding only a monitored break-glass account. Without Entra ID P1, turn on security defaults instead.' }
    @{ Control = 'Phishing-resistant MFA (passkeys) rollout';    Bad = '^Not started';             Priority = 'Medium';   Action = 'Enable the passkey (FIDO2) method for all users, have administrators and finance staff register passkeys first (Microsoft Authenticator passkey or a hardware key), then roll out to everyone and retire SMS and voice. Phishing-resistant MFA stops the real-time phishing kits that defeat push and code-based MFA.' }
    @{ Control = 'Phishing-resistant MFA for administrators';    Bad = '^(None|Partial)';          Priority = 'High';     Action = 'Every administrator registers a passkey or FIDO2 key and a Conditional Access policy requires the Phishing-resistant MFA authentication strength for all administrator roles. Administrators are the accounts attackers phish for, and push-based MFA is routinely bypassed.' }
    @{ Control = 'MFA registered by all users';                   Bad = '^No$';                    Priority = 'High';     Action = 'Have every user listed under Users register Microsoft Authenticator (or a FIDO2 key). Use a registration campaign, then enforce MFA so unregistered accounts cannot sign in without completing it.' }
    @{ Control = 'Legacy authentication protocols disabled';      Bad = '^(Not blocked|Partial)';  Priority = 'High';     Action = 'Add a Conditional Access policy that blocks Exchange ActiveSync and "other clients" for all users. In Exchange Online disable SMTP AUTH tenant-wide and turn off POP and IMAP on every mailbox that does not need them.' }
    @{ Control = 'Automatic email forwarding outside the tenant'; Bad = '^(Active external|Allowed by policy)'; Priority = 'High'; Action = 'Set automatic forwarding to Off in the outbound spam filter policy, set the default remote domain to not allow auto-forward, then review and remove the mailbox forwards and inbox rules listed under Exchange Online.' }
    @{ Control = 'Conditional Access restricts access to Entra joined or compliant devices'; Bad = '^(Not enforced|Partial)'; Priority = 'Medium'; Action = 'Enroll company devices in Intune, then add a Conditional Access policy that requires a compliant or hybrid-joined device for Microsoft 365 apps. Start in report-only mode and review the sign-in impact before enforcing.' }
    @{ Control = 'Intune deployed to Windows computers';          Bad = '^(Not deployed|Partial)'; Priority = 'Medium';   Action = 'Enroll the Windows devices listed under Managed Devices into Intune (automatic enrollment through Entra join or group policy) so they can be measured against compliance policies.' }
    @{ Control = 'Intune compliance policies for mobile devices';  Bad = '^None';                   Priority = 'Critical'; Action = 'Create an Intune compliance policy for iOS/iPadOS and one for Android (work profile) that require a passcode, encryption, a minimum OS version and no jailbreak or root, assign them to all users, and pair them with a Conditional Access policy that requires a compliant device for Exchange Online and Teams. Phones are where institution mail is read outside the building.' }
    @{ Control = 'Intune compliance policies';                    Bad = '^None';                   Priority = 'Critical';   Action = 'Create a compliance policy per platform (Windows, iOS/iPadOS, Android, macOS) that requires encryption, a PIN or password, a minimum OS version and real-time protection, and marks devices noncompliant after a short grace period.' }
    @{ Control = 'Intune configuration and security policies';    Bad = '^0 policies';             Priority = 'Medium';   Action = 'Deploy a baseline: Microsoft security baseline, BitLocker, Defender Antivirus and firewall, screen lock, and Windows update rings. Assign to all devices and confirm the settings report as applied.' }
    @{ Control = 'User consent to applications';                  Bad = '^Allowed';                Priority = 'Critical';   Action = 'Set user consent to "Do not allow" (or limit it to verified publishers and low-impact permissions) and enable the admin consent request workflow with named reviewers, so every new application is approved by an administrator.' }
    @{ Control = 'Temporary Access Pass (TAP) implemented';       Bad = '^Disabled';               Priority = 'Low';      Action = 'Enable Temporary Access Pass for onboarding and MFA recovery: one-time use, a short lifetime, scoped to a help-desk-managed group.' }
    @{ Control = 'Users restricted from joining devices to Entra ID'; Bad = '^All users allowed';  Priority = 'Critical';      Action = 'Restrict Entra ID device join to selected users or groups and require MFA to register or join a device.' }
    @{ Control = 'SPF record (DNS)';                              Bad = '^(Partial|Missing)';      Priority = 'High';     Action = 'Publish one SPF TXT record per sending domain that lists every legitimate sender and ends in -all (hard fail). This is a DNS change at the domain registrar or DNS host, not a Microsoft 365 setting.'; Url = 'https://learn.microsoft.com/microsoft-365/security/office-365-security/email-authentication-spf-configure' }
    @{ Control = 'DKIM signing (Exchange Online and DNS)';        Bad = '^(Partial|Missing)';      Priority = 'High';     Action = 'On the Defender portal DKIM page create the keys for each domain, publish the two selector CNAME records it shows in public DNS, then enable signing. Without DKIM, DMARC cannot reach reject and forwarded mail fails authentication.'; Url = 'https://security.microsoft.com/dkimv2' }
    @{ Control = 'DMARC policy (DNS)';                            Bad = '^(Partial|Missing)';      Priority = 'High';     Action = 'Publish a _dmarc TXT record with aggregate reporting, monitor for two to four weeks, then move to p=quarantine and p=reject at pct=100. DMARC is what actually tells the world''s mail servers to refuse mail spoofing the institution. DNS change, not a Microsoft 365 setting.'; Url = 'https://learn.microsoft.com/microsoft-365/security/office-365-security/email-authentication-dmarc-configure' }
    @{ Control = 'Email protection (Defender for Office 365 / EOP)'; Bad = '^(Licensed, not configured|EOP only)'; Priority = 'High'; Action = 'If Defender for Office 365 is licensed, enable the Standard preset security policy for all users (Strict for executives and finance) so Safe Links, Safe Attachments and impersonation protection apply. If it is not licensed, license Plan 1 (included in Business Premium); phishing is the top attack path against bank staff.' }
    @{ Control = 'Endpoint protection (Defender for Endpoint)';  Bad = '^(Attention needed|No devices onboarded|Not licensed|Licensed, not deployed|Licensed, not assigned|Licensed, setup not finished|Licensed, no onboarding policy|Partly deployed)'; Priority = 'Medium'; Action = 'Onboard every managed Windows and macOS device to Defender for Endpoint / Defender for Business through Intune, resolve high-risk alerts, and re-onboard or retire devices with inactive sensors.' }
)
# Where each control is changed: the admin center page the ticket should point at (Secure Score items carry Microsoft's own link)
$ControlPortalUrl = @{
    'MFA enforced for sign-in'                                   = 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Policies'
    'MFA registered by all users'                                = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/UserRegistrationDetails'
    'Legacy authentication protocols disabled'                   = 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Policies'
    'Conditional Access restricts access to Entra joined or compliant devices' = 'https://entra.microsoft.com/#view/Microsoft_AAD_ConditionalAccess/ConditionalAccessBlade/~/Policies'
    'Intune deployed to Windows computers'                       = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/overview'
    'Intune compliance policies'                                 = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/compliance'
    'Intune compliance policies for mobile devices'              = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/compliance'
    'Intune configuration and security policies'                 = 'https://intune.microsoft.com/#view/Microsoft_Intune_DeviceSettings/DevicesMenu/~/configuration'
    'Temporary Access Pass (TAP) implemented'                    = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/AuthenticationMethodsMenuBlade/~/AdminAuthMethods'
    'Users restricted from joining devices to Entra ID'          = 'https://entra.microsoft.com/#view/Microsoft_AAD_Devices/DevicesMenuBlade/~/DeviceSettings'
    'Automatic email forwarding outside the tenant'              = 'https://security.microsoft.com/antispam'
    'User consent to applications'                               = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ConsentPoliciesMenuBlade/~/UserSettings'
    'SPF record (DNS)'                                           = 'https://learn.microsoft.com/microsoft-365/security/office-365-security/email-authentication-spf-configure'
    'DKIM signing (Exchange Online and DNS)'                     = 'https://security.microsoft.com/dkimv2'
    'DMARC policy (DNS)'                                         = 'https://learn.microsoft.com/microsoft-365/security/office-365-security/email-authentication-dmarc-configure'
    'Email protection (Defender for Office 365 / EOP)'           = 'https://security.microsoft.com/presetSecurityPolicies'
    'Endpoint protection (Defender for Endpoint)'                = 'https://security.microsoft.com/machines'
    'Microsoft Secure Score'                                     = 'https://security.microsoft.com/securescore'
    'Security defaults'                                          = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/Properties'
    'Audit logging (unified audit log)'                          = 'https://purview.microsoft.com/audit/auditsearch'
}
foreach ($g in $remediationGuides) {
    $row = $script:summaryRows | Where-Object { $_.Control -eq $g.Control } | Select-Object -First 1
    if ($row -and ([string]$row.Status -match $g.Bad)) { Add-Remediation -Priority $g.Priority -Finding $g.Control -State "$($row.Status). $($row.Evidence)" -Action $g.Action -Section $row.Section -Url $(if ($g.Url) { $g.Url } elseif ($ControlPortalUrl.ContainsKey($g.Control)) { $ControlPortalUrl[$g.Control] } else { $null }) }
}
# Licensing gaps that stop whole control families from existing
if ($script:HasIntune -eq $false)  { Add-Remediation -Priority 'High' -Finding 'Intune is not licensed' -State 'No Intune service plan in the subscribed SKUs, so device compliance, configuration policies and device-based Conditional Access cannot be implemented.' -Action 'License Microsoft Intune (included in Microsoft 365 Business Premium, E3 and E5) and then complete the device controls above.' -Section 'Tenant Overview' }
if ($shareAnyone -gt 0) { Add-Remediation -Priority 'High' -Finding 'Content open to anyone on the internet' -State "$shareAnyone Anyone link(s) or anonymous calendar permission(s) found in the share scan (section 19)." -Action 'Remove each Anyone link from the item (Manage access), set calendar Anonymous permissions to None, then set the SharePoint and OneDrive sharing level to New and existing guests so Anyone links cannot be created again.' -Section 'Sharing and Access' -Url 'https://admin.microsoft.com/sharepoint?page=sharing&modern=true' }
if ($shareOrg -gt 0) { Add-Remediation -Priority 'Medium' -Finding 'Content open to the whole organization' -State "$shareOrg whole-organization link(s), public team(s) or organization-readable calendar(s) (section 19)." -Action 'Replace People-in-your-organization links with specific people, make teams holding confidential material Private, set the default link type to Specific people, and set calendar Default permissions to free/busy.' -Section 'Sharing and Access' -Url 'https://admin.microsoft.com/sharepoint?page=sharing&modern=true' }
if ($reportConcealedNote) { Add-Remediation -Priority 'High' -Finding 'Usage report names are concealed' -State 'The Microsoft 365 usage reports return hashes instead of user, group and site names, so the OneDrive, SharePoint site and Teams inventories cannot be attributed to people.' -Action 'In the Microsoft 365 admin center clear the report privacy setting "Display concealed user, group, and site names in all reports" (Settings > Org settings > Services > Reports). One-time tenant setting; the script never changes it.' -Section 'SharePoint' -Url 'https://admin.microsoft.com/#/Settings/Services/:/Settings/L1/Reports' }
if ($script:HasEntraP1 -eq $false) { Add-Remediation -Priority 'Low' -Finding 'Entra ID P1 is not licensed' -State 'No Entra ID P1/P2 service plan, so Conditional Access, the MFA registration report and sign-in activity are unavailable; protection relies on security defaults.' -Action 'License Entra ID P1 (included in Microsoft 365 Business Premium and E3) to gain Conditional Access and the MFA registration and sign-in reports this report measures.' -Section 'Tenant Overview' }
# Privileged access
$gaRole  = $roles | Where-Object { $_.displayName -eq 'Global Administrator' } | Select-Object -First 1
$gaCount = @($gaRole.members).Count
$gaNoMfa = @($roleRows | Where-Object { $_.Role -eq 'Global Administrator' -and $_.'MFA registered' -eq 'No' })
$RolesPortalUrl = 'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/RolesManagementMenuBlade/~/AllRoles'
if ($gaNoMfa.Count) { Add-Remediation -Url $RolesPortalUrl -Priority 'Critical' -Finding 'Global Administrators without MFA registered' -State ("$($gaNoMfa.Count) account(s): " + (Join-List ($gaNoMfa | ForEach-Object { $_.Member }))) -Action 'Register Microsoft Authenticator or a FIDO2 key for these accounts today and require MFA for all administrator roles through Conditional Access.' -Section 'Privileged Users' }
if ($gaCount -gt 4)  { Add-Remediation -Url $RolesPortalUrl -Priority 'Medium' -Finding 'Too many Global Administrators' -State "$gaCount accounts hold Global Administrator (Microsoft recommends fewer than five)." -Action 'Reduce Global Administrators to two to four, move day-to-day administration to least-privilege roles (Exchange, Intune, User Administrator), and use PIM eligible assignments where Entra ID P2 is licensed.' -Section 'Privileged Users' }
if ($gaCount -eq 1)  { Add-Remediation -Url $RolesPortalUrl -Priority 'Critical' -Finding 'Single Global Administrator' -State 'Only one account holds Global Administrator; losing access to it locks the organization out of the tenant.' -Action 'Create a second, cloud-only break-glass Global Administrator with a long random password stored offline, excluded from Conditional Access, and alerted on every sign-in.' -Section 'Privileged Users' }
# Secure Score improvement actions that still have points available
$ProfileById = @{}
foreach ($p in $scoreProfiles) { $ProfileById[$p.id] = $p }
$scoreIgnored = 0
foreach ($cs in @($secureScore.controlScores)) {
    $p = $ProfileById[$cs.controlName]
    if (-not $p -or $p.deprecated) { continue }
    $lastState = (@($p.controlStateUpdates) | Sort-Object updatedDateTime | Select-Object -Last 1).state
    if ($lastState -in @('Ignored', 'ThirdParty', 'Resolved')) { $scoreIgnored++; continue }
    $avail = [math]::Round([double]$p.maxScore - [double]$cs.score, 1)
    if ($avail -le 0) { continue }
    $prio = if ($avail -ge 8) { 'High' } elseif ($avail -ge 3) { 'Medium' } else { 'Low' }
    $state = "$([math]::Round([double]$cs.score, 1)) of $($p.maxScore) points.$(if ($cs.implementationStatus) { " $(ConvertFrom-HtmlText $cs.implementationStatus)" })"
    $action = ConvertFrom-HtmlText ([string]$p.remediation)
    if ($action.Length -gt 700) { $action = $action.Substring(0, 700).TrimEnd() + '...' }
    if (-not $action) { $action = ConvertFrom-HtmlText ([string]$cs.description) }
    Add-Remediation -Priority $prio -Finding ([string]$p.title) -State $state -Action $action -Points $avail -Impact "User impact: $($p.userImpact); effort: $($p.implementationCost); tier: $($p.tier)" -Section 'Secure Score' -Url ([string]$p.actionUrl)
}
$remediationSorted = @($script:remediationItems | Sort-Object Rank, @{ Expression = { if ($null -ne $_.Points) { 1 } else { 0 } } }, @{ Expression = { if ($_.Points) { [double]$_.Points } else { 0 } }; Descending = $true }, Finding)
$remediationCounts = Join-List (@('Critical', 'High', 'Medium', 'Low') | ForEach-Object { $pr = $_; "$pr`: $(@($remediationSorted | Where-Object { $_.Priority -eq $pr }).Count)" }) '; '

# ----------------------------------------------------------------------------------------------
# Build HTML
# ----------------------------------------------------------------------------------------------
Write-Host "Building report..." -ForegroundColor Cyan
$css = @"
:root{--blue:#005D9C;--ink:#000;--line:#d9d9d9;--nav:#f3f6f9;--good:#1e7b34;--warn:#b00020}
*{box-sizing:border-box}body{margin:0;font-family:Arial,Helvetica,sans-serif;font-size:11pt;color:var(--ink);line-height:1.35}
nav{position:fixed;top:0;left:0;bottom:0;width:260px;background:var(--nav);border-right:1px solid var(--line);padding:18px 14px;overflow:auto}
nav h1{font-size:14pt;color:var(--blue);margin:0 0 4px}nav .sub{font-size:9pt;color:#555;margin-bottom:14px}
nav a{display:block;color:var(--ink);text-decoration:none;padding:6px 8px;border-radius:4px;font-size:10.5pt}nav a:hover{background:#e1e9f1}
nav a.l2{padding-left:22px;font-size:10pt;color:#333}nav img.logo{width:130px;display:block;margin:0 0 10px}nav .sub a,.privacy a,footer.brand a{color:var(--blue)}footer.brand{margin-top:48px;padding-top:14px;border-top:1px solid var(--line);font-size:9.5pt;color:#555;display:flex;gap:14px;align-items:center}footer.brand img.logo{width:70px}
main{margin-left:260px;padding:28px 40px 60px;max-width:1500px}
nav a.active{background:var(--blue);color:#fff}nav a.active:hover{background:var(--blue)}
body.paged main>section{display:none}body.paged main>section.active{display:block}
.pager{position:sticky;top:0;z-index:5;background:#fff;border-bottom:1px solid var(--line);display:flex;flex-wrap:wrap;gap:10px;align-items:center;padding:8px 0 10px;margin:-14px 0 18px}
.pager button{font:inherit;font-size:9.5pt;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:4px;padding:4px 12px;cursor:pointer}.pager button:disabled{opacity:.45;cursor:default}.pager button.primary{background:var(--blue);color:#fff}
.pager .pos{font-size:9.5pt;color:#555;margin-left:auto}.pager .title{font-size:10.5pt;color:var(--blue);font-weight:bold}
body.paged main>section.active>h2{position:sticky;top:var(--pager-h,46px);z-index:4;background:#fff;margin-top:0;padding:10px 0 8px;box-shadow:0 6px 8px -8px rgba(0,0,0,.35)}
body.paged .pager .title{display:none}
section{margin-bottom:48px;scroll-margin-top:12px}
h2{font-size:20pt;color:var(--blue);border-bottom:2px solid var(--blue);padding-bottom:4px;margin:0 0 14px}
h3{font-size:13pt;color:var(--blue);margin:26px 0 8px}h3.platform{margin-top:34px;padding:6px 10px;background:#eef3f8;border-left:4px solid var(--blue)}
h4{font-size:12pt;margin:0 0 10px;color:var(--ink)}h5{font-size:11pt;color:var(--blue);margin:14px 0 6px}
.card{border:1px solid var(--line);border-radius:6px;padding:16px 18px;margin:14px 0;page-break-inside:avoid}
table{border-collapse:collapse;width:100%;margin:6px 0 12px}
th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top;text-align:left;font-size:10pt;word-break:break-word}
table.grid thead th{background:var(--blue);color:#fff;position:sticky;top:0}
table.kv th{width:280px;background:#f3f6f9;font-weight:bold}
tbody tr:nth-child(even) td{background:#fafafa}
td.good{color:var(--good);font-weight:bold}td.warn{color:var(--warn);font-weight:bold}td.mid{color:#B8860B;font-weight:bold}
.st{display:inline-block;width:18px;height:18px;line-height:18px;text-align:center;border-radius:50%;color:#fff;font-size:10pt;font-weight:bold;margin-right:5px;vertical-align:middle}.st.good{background:var(--good)}.st.warn{background:var(--warn)}.st.mid{background:#B8860B}.st.na{background:#9a9a9a}.st.info{background:#7a8aa0}
.stt{font-weight:bold;vertical-align:middle}.stt.good{color:var(--good)}.stt.warn{color:var(--warn)}.stt.mid{color:#8a6508}.stt.na{color:#666;font-weight:normal}.stt.info{font-weight:normal}
.legend{font-size:9.5pt;color:#555;margin:6px 0 10px}
.count{font-size:10pt;color:#555;font-weight:normal}
.empty{color:#555;font-style:italic}.note{background:#fff8e1;border:1px solid #f0d47a;padding:10px 12px;border-radius:4px;font-size:10pt}
.meta{font-size:9.5pt;color:#555;margin-bottom:18px}
ul.kvlist{margin:0;padding-left:18px}ul.kvlist li{margin:2px 0}
.filters{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:10px 0 4px}
.filters button{font:inherit;font-size:9.5pt;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:14px;padding:3px 11px;cursor:pointer}
.filters button.on{background:var(--blue);color:#fff}.filters input{font:inherit;font-size:9.5pt;padding:4px 8px;border:1px solid var(--line);border-radius:4px;min-width:220px;margin-left:auto}
.filters .n{font-size:9pt;color:#555}
.ticketbar{display:flex;flex-wrap:wrap;gap:8px;align-items:center;margin:6px 0 10px}
.ticketbar button,button.tk{font:inherit;font-size:9.5pt;border:1px solid var(--blue);background:var(--blue);color:#fff;border-radius:4px;padding:4px 10px;cursor:pointer}
button.tk{font-size:8.5pt;padding:2px 7px;background:#fff;color:var(--blue);white-space:nowrap}
.ticketbar button:disabled,button.tk:disabled{opacity:.7;cursor:default}.ticketbar .hint{font-size:9pt;color:#555;flex:1 1 320px}
td.ref{font-family:Consolas,monospace;font-size:9pt;white-space:nowrap}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:12px;margin:14px 0 22px}
.kpi{border:1px solid var(--line);border-left:5px solid var(--blue);border-radius:6px;padding:12px 14px;background:#fff;page-break-inside:avoid}
.kpi .v{font-size:20pt;font-weight:bold;color:var(--blue);line-height:1.1}.kpi .l{font-size:9pt;color:#555;margin-top:4px}
.kpi.good{border-left-color:var(--good)}.kpi.good .v{color:var(--good)}.kpi.bad{border-left-color:var(--warn)}.kpi.bad .v{color:var(--warn)}.kpi.mid{border-left-color:#B8860B}.kpi.mid .v{color:#B8860B}
.exec{font-size:11.5pt;line-height:1.5;max-width:1100px}.exec ol{margin:6px 0 0 0;padding-left:22px}.exec li{margin:4px 0}
details.info{display:inline-block;position:relative;margin-left:6px;vertical-align:middle}
details.info summary{list-style:none;cursor:pointer;display:inline-flex;width:17px;height:17px;border-radius:50%;background:var(--warn);color:#fff;font:bold 11px Georgia,'Times New Roman',serif;align-items:center;justify-content:center;line-height:1}
details.info summary::-webkit-details-marker{display:none}details.info summary::marker{content:''}
details.info[open] summary{background:#7a0016}
details.info .tip{position:absolute;z-index:30;left:0;top:22px;width:380px;max-width:70vw;background:#fff;border:1px solid var(--line);border-left:4px solid var(--warn);box-shadow:0 6px 18px rgba(0,0,0,.18);padding:10px 12px;font-size:9.5pt;font-weight:normal;color:var(--ink);text-align:left;border-radius:4px}
details.info .tip p{margin:6px 0 0}
.legend{font-size:9.5pt;color:#555;margin:0 0 10px}.legend .i{display:inline-flex;width:15px;height:15px;border-radius:50%;background:var(--warn);color:#fff;font:bold 10px Georgia,serif;align-items:center;justify-content:center;vertical-align:middle}
.charts{display:flex;gap:18px;flex-wrap:wrap;margin:10px 0 20px}.chart{flex:1 1 460px;min-width:0;border:1px solid var(--line);border-radius:6px;padding:8px 10px;background:#fff}.chart svg{display:block}.chart svg text{font-family:Arial,Helvetica,sans-serif}
.roadmap{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:12px;margin:10px 0 18px}.roadmap .phase{border:1px solid var(--line);border-radius:6px;padding:10px 12px;background:#fff}.roadmap h4{margin:0 0 6px;color:var(--blue);font-size:11pt}.roadmap .n{display:inline-block;background:var(--blue);color:#fff;border-radius:10px;padding:0 8px;font-size:9pt;margin-left:8px}.roadmap ol{margin:0;padding-left:18px;font-size:10pt}.roadmap li{margin:3px 0}
.muted{color:#555;font-size:9pt}details.steps{margin-top:4px}details.steps summary{cursor:pointer;color:var(--blue);font-size:9.5pt}details.steps ol{margin:4px 0;padding-left:18px;font-size:9.5pt}details.steps p{font-size:9.5pt;margin:4px 0}
.howto{background:#f3f6f9;border:1px solid var(--line);border-radius:6px;padding:10px 14px;margin:10px 0 16px;font-size:10pt}.howto ul{margin:4px 0 0;padding-left:18px}.howto li{margin:3px 0}
.modal{position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.45);display:flex;align-items:center;justify-content:center;z-index:50}.modal[hidden]{display:none}
.modal-box{background:#fff;border-radius:8px;max-width:780px;width:92%;max-height:86vh;display:flex;flex-direction:column;box-shadow:0 10px 30px rgba(0,0,0,.3)}
.modal-head{display:flex;justify-content:space-between;align-items:center;gap:12px;padding:10px 14px;border-bottom:1px solid var(--line)}.modal-head b{font-size:11pt;color:var(--blue)}
.modal-head button,.modal-actions button{font:inherit;font-size:9.5pt;border:1px solid var(--blue);background:var(--blue);color:#fff;border-radius:4px;padding:4px 12px;cursor:pointer}.modal-head button{background:#fff;color:var(--blue)}
.modal pre{margin:0;padding:12px 14px;overflow:auto;white-space:pre-wrap;word-break:break-word;font:10pt Consolas,monospace;flex:1 1 auto;background:#fafafa}.modal-actions{display:flex;gap:12px;align-items:center;padding:10px 14px;border-top:1px solid var(--line)}.modal-actions .hint{font-size:9pt;color:#555}body.noscroll{overflow:hidden}
.perms{margin:0;padding-left:16px;font-size:9.5pt}.perms li{margin:1px 0}button.more{font:inherit;font-size:9pt;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:4px;padding:1px 8px;cursor:pointer;margin-top:4px}
.modal-body{padding:10px 14px;overflow:auto;font-size:10pt;flex:1 1 auto}.modal-body h5{margin:10px 0 4px;color:var(--blue);font-size:10pt}.modal-body ul{margin:0;padding-left:18px}.modal-body li{margin:2px 0}
.audience{font-size:10.5pt;color:#555;max-width:1000px;margin:0 0 10px}
.verdict{border-radius:10px;padding:18px 22px;margin:8px 0 18px;border:1px solid var(--line);border-left:10px solid var(--blue);background:#fff}
.verdict.bad{border-left-color:var(--warn);background:#fff5f5}.verdict.mid{border-left-color:#B8860B;background:#fffaf0}.verdict.good{border-left-color:var(--good);background:#f3faf4}
.verdict .vlabel{font-size:9.5pt;text-transform:uppercase;letter-spacing:.08em;color:#555}.verdict .vword{font-size:26pt;font-weight:bold;line-height:1.1;margin:2px 0 6px;color:var(--blue)}.verdict.bad .vword{color:var(--warn)}.verdict.mid .vword{color:#8a6508}.verdict.good .vword{color:var(--good)}
.verdict p{margin:4px 0 10px;font-size:11.5pt;max-width:1000px}.chips{display:flex;flex-wrap:wrap;gap:8px}.chip{display:inline-block;border-radius:14px;padding:3px 12px;font-size:9.5pt;border:1px solid var(--line);background:#fff;color:#333}.chip.good{border-color:var(--good);color:var(--good)}.chip.bad{border-color:var(--warn);color:var(--warn)}.chip.mid{border-color:#B8860B;color:#8a6508}
.eos{display:grid;grid-template-columns:repeat(auto-fit,minmax(380px,1fr));gap:24px;margin:4px 0 10px}.eos h3{margin-top:6px}
ul.headlines{list-style:none;padding:0;margin:6px 0}ul.headlines li{display:flex;gap:10px;align-items:flex-start;padding:7px 0;border-bottom:1px solid var(--line);font-size:11pt}ul.headlines .dot{flex:0 0 12px;height:12px;border-radius:50%;margin-top:5px;background:var(--good)}ul.headlines li.bad .dot{background:var(--warn)}
ol.prio{list-style:none;padding:0;margin:6px 0 14px}ol.prio li{display:flex;gap:12px;align-items:flex-start;padding:7px 0;border-bottom:1px solid var(--line)}ol.prio .badge{flex:0 0 78px;text-align:center;border-radius:4px;padding:3px 0;font-size:9pt;font-weight:bold;color:#fff;background:var(--blue)}ol.prio .badge.Critical,ol.prio .badge.High{background:var(--warn)}ol.prio .badge.Medium{background:#B8860B}ol.prio .badge.Low{background:#777}ol.prio .who{color:#555;font-size:9.5pt}
details.more{border:1px solid var(--line);border-radius:6px;padding:8px 12px;margin:10px 0;background:#fff}details.more summary{cursor:pointer;color:var(--blue);font-weight:bold;font-size:10.5pt}details.more[open] summary{margin-bottom:8px}hr.divider{border:0;border-top:2px solid var(--line);margin:26px 0 14px}
.kpi.bad{border-left-color:var(--warn)}.kpi.bad .v{color:var(--warn)}.kpi.good{border-left-color:var(--good)}.kpi.good .v{color:var(--good)}.kpi.mid{border-left-color:#B8860B}.kpi.mid .v{color:#8a6508}
.privacy{background:#eef6ee;border:1px solid #bfe0c2;border-left:5px solid var(--good);padding:10px 14px;border-radius:4px;font-size:10pt;margin:16px 0}
@media print{nav,.pager{display:none}body.paged main>section{display:block}.ticketbar,button.tk,button.more,.modal{display:none}.permfull{display:block!important}main{margin:0;padding:0}section{page-break-before:always}section:first-of-type{page-break-before:avoid}table.grid thead th{position:static}.filters{display:none}details.info .tip{position:static;display:block;width:auto;box-shadow:none;margin-top:4px}details.info summary{display:none}details.info:not([open]) .tip{display:none}}
"@

$sections = @(
    @{ Id = 'overview';     Title = '1. Executive Summary' },
    @{ Id = 'ioc';          Title = '2. Review These Now: Indications of Compromise' },
    @{ Id = 'summary';      Title = '3. Security Controls Summary' },
    @{ Id = 'remediation';  Title = '4. Remediation Plan' },
    @{ Id = 'frameworks';   Title = '5. Framework Alignment (NIST, CIS, FFIEC)' },
    @{ Id = 'mdm';          Title = '6. MDM Compliance Policies'; Sub = @($policyGroups | ForEach-Object { @{ Id = ('mdm-' + ($_.Name -replace '[^A-Za-z0-9]', '')); Title = $_.Name } }) },
    @{ Id = 'profiles';     Title = '7. Intune Policies' },
    @{ Id = 'devices';      Title = '8. Managed Devices' },
    @{ Id = 'ca';           Title = '9. Conditional Access' },
    @{ Id = 'authmethods';  Title = '10. Authentication Methods' },
    @{ Id = 'apps';         Title = '11. Enterprise Applications' },
    @{ Id = 'consent';      Title = '12. Application Consent Audit Events' },
    @{ Id = 'users';        Title = '13. Users' },
    @{ Id = 'roles';        Title = '14. Privileged Users' },
    @{ Id = 'exchange';     Title = '15. Exchange Online' },
    @{ Id = 'defender';     Title = '16. Defender and Email Protection' },
    @{ Id = 'teams';        Title = '17. Microsoft Teams' },
    @{ Id = 'sharepoint';   Title = '18. SharePoint and OneDrive' },
    @{ Id = 'sharing';      Title = '19. Sharing and Access' }
)
$SectionAnchors = @{ 'Authentication Methods' = 'authmethods'; 'Conditional Access' = 'ca'; 'Users' = 'users'; 'Managed Devices' = 'devices'; 'MDM Compliance Policies' = 'mdm'; 'Intune Policies' = 'profiles'; 'Exchange Online' = 'exchange'; 'Enterprise Applications' = 'apps'; 'Privileged Users' = 'roles' }
$SectionAnchors['Tenant Overview'] = 'overview'; $SectionAnchors['Executive Summary'] = 'overview'; $SectionAnchors['Remediation Plan'] = 'remediation'; $SectionAnchors['Secure Score'] = 'overview'; $SectionAnchors['Privileged Users'] = 'roles'; $SectionAnchors['Exchange Online'] = 'exchange'; $SectionAnchors['Defender'] = 'defender'; $SectionAnchors['Enterprise Applications'] = 'apps'
$SectionAnchors['Indications of Compromise'] = 'ioc'; $SectionAnchors['Framework Alignment'] = 'frameworks'; $SectionAnchors['Sharing and Access'] = 'sharing'; $SectionAnchors['Teams'] = 'teams'; $SectionAnchors['SharePoint'] = 'sharepoint'
# ---- Directory synchronization (Entra Connect) for the overview
$syncEnabled = [bool]$org.onPremisesSyncEnabled
$syncLast    = $org.onPremisesLastSyncDateTime
$adDomains   = @($users | ForEach-Object { $_.onPremisesDomainName } | Where-Object { $_ } | Sort-Object -Unique)
$syncServers = @($users | Where-Object { $_.userPrincipalName -match '^Sync_(.+)_[0-9a-fA-F]{8,}@' } | ForEach-Object { $Matches[1] } | Sort-Object -Unique)
$syncedUserCount = @($users | Where-Object { $_.onPremisesSyncEnabled }).Count
$tenantKv['Directory synchronization (Entra Connect)'] = $(if ($syncEnabled) { "Enabled. Last sync $(Format-DateText $syncLast). $syncedUserCount synced user account(s). The on-premises directory itself is covered by the companion AuditHat AD Governance Report ($AdRepoUrl)." } else { 'Not configured: all accounts are cloud-only.' })
if ($syncEnabled) {
    $tenantKv['On-premises AD domain(s)'] = $(if ($adDomains.Count) { $adDomains } else { 'Not reported on synced users' })
    $tenantKv['Entra Connect sync server'] = $(if ($syncServers.Count) { "$($syncServers -join '; ') (from the Sync_<server>_ service account name)" } else { 'Not determinable from Graph (no Sync_<server>_ service account found)' })
}
$tenantIssues = @{}
if ($syncEnabled -and $syncLast -and ([datetime]$syncLast) -lt (Get-Date).AddHours(-3)) { $tenantIssues['Directory synchronization (Entra Connect)'] = @('sync-stale', "Last sync $(Format-DateText $syncLast).") }
if ($script:HasEntraP1 -eq $false) { $tenantIssues['Licensed capabilities'] = 'license-p1' } elseif ($script:HasIntune -eq $false) { $tenantIssues['Licensed capabilities'] = 'license-intune' }
if ($null -ne $scoreCurrent -and $scoreCurrent -lt $scoreTarget) { $tenantIssues['Microsoft Secure Score'] = 'secure-score' }

# ---- Flag configuration issues on rows and key/value tables (rendered as clickable (i) icons)
$SummaryIssueMap = @{
    'Phishing-resistant MFA (passkeys) rollout' = 'passkey-not-started'; 'Phishing-resistant MFA for administrators' = 'passkey-admins'
    'Indications of compromise to review' = 'ioc-review'; 'Audit logging (unified audit log)' = 'ual-off'; 'Teams guest and external access' = 'groups-guests-open'; 'SharePoint and OneDrive external sharing' = 'spo-external-sharing'
    'MFA enforced for sign-in' = 'mfa-not-enforced'; 'MFA registered by all users' = 'mfa-not-registered'; 'Legacy authentication protocols disabled' = 'legacy-auth'
    'Conditional Access restricts access to Entra joined or compliant devices' = 'device-ca'; 'Intune deployed to Windows computers' = 'intune-coverage'
    'Intune compliance policies' = 'compliance-none'; 'Intune compliance policies for mobile devices' = 'compliance-none'; 'Intune configuration and security policies' = 'config-none'; 'Temporary Access Pass (TAP) implemented' = 'tap-disabled'
    'Users restricted from joining devices to Entra ID' = 'device-join-open'; 'Automatic email forwarding outside the tenant' = 'forwarding'; 'User consent to applications' = 'consent-allowed'
    'Microsoft Secure Score' = 'secure-score'; 'Security defaults' = 'security-defaults-off'; 'SPF record (DNS)' = 'spf-weak'; 'DKIM signing (Exchange Online and DNS)' = 'dkim-missing'; 'DMARC policy (DNS)' = 'dmarc-missing'; 'Email protection (Defender for Office 365 / EOP)' = 'defender-not-licensed'; 'Endpoint protection (Defender for Endpoint)' = 'mde-not-licensed'
}
foreach ($r in $userRows) {
    $tg = " $($r['_tags']) "; $is = @{}
    if ($tg -like '* nomfa *') { $is['MFA registered'] = $(if ($tg -like '* admin *') { 'admin-no-mfa' } else { 'user-no-mfa' }) }
    if ($tg -like '* stale *') { $is['Last sign-in'] = 'user-stale' }
    if ($tg -like '* disabled *' -and $tg -like '* licensed *') { $is['Licenses'] = 'user-disabled-licensed' }
    if ($tg -like '* guest *') { $is['User principal name'] = 'guest' }
    if ($is.Count) { $r['_issues'] = $is }
}
foreach ($r in $allDeviceRows) {
    $tg = " $($r['_tags']) "; $is = @{}
    if ($tg -like '* stale *') { $is['Last activity'] = 'device-stale' }
    if ($tg -like '* notintune *' -and $r['Enabled'] -ne $false) { $is['Intune'] = 'device-not-intune' }
    if ($tg -like '* noncompliant *') { $is['Compliance'] = 'device-noncompliant' }
    if ($tg -like '* notencrypted *') { $is['Encrypted'] = 'device-not-encrypted' }
    if ($is.Count) { $r['_issues'] = $is }
}
$HighPrivPerm = 'ReadWrite\.All|Directory\.ReadWrite|RoleManagement\.ReadWrite|AppRoleAssignment\.ReadWrite|Mail\.(Read|ReadWrite|Send)|Files\.ReadWrite\.All|Sites\.FullControl|full_access_as_app|Exchange\.ManageAsApp|User\.ReadWrite\.All|Group\.ReadWrite\.All|Application\.ReadWrite'
$permIdx = 0
foreach ($r in $appRows) {
    $permIdx++
    $is = @{}; $tags = @()
    $isMicrosoft = ([string]$r['Publisher']) -eq 'Microsoft'
    $unverified = ([string]$r['Publisher']) -match 'unverified|Unknown'
    $appPermText = [string]$r['Application permissions']; $adminPermText = [string]$r['Admin-consented delegated permissions']; $userPermText = ([string]$r['User-consented delegated permissions']) -replace '^\d+ grant\(s\): ', ''
    $isRogue = ($r['App ID'] -and $RogueApps.ContainsKey([string]$r['App ID'])) -or (([string]$r['Application']) -match $RogueAppNamePattern)
    $allPermText = "$appPermText $adminPermText $userPermText"
    $risk = if ($isRogue) { 'High' }
            elseif ($appPermText -ne 'None' -and $appPermText -match $HighPrivPerm) { 'High' }
            elseif ($unverified -and $allPermText -match $RiskyPermPattern) { 'High' }
            elseif ($userPermText -ne 'None' -or ($adminPermText -ne 'None' -and $adminPermText -match $HighPrivPerm)) { 'Medium' }
            elseif ($unverified -and -not $isMicrosoft) { 'Medium' }
            else { 'Low' }
    if ($isRogue) { $is['Risk'] = 'ioc-rogue-app' }
    elseif ($risk -eq 'High') { $is['Risk'] = $(if ($unverified -and $appPermText -eq 'None') { 'ioc-unverified-app' } else { 'app-high-privilege' }) }
    elseif ($userPermText -ne 'None') { $is['Risk'] = 'app-user-consent' }
    elseif ($unverified) { $is['Risk'] = 'app-unverified' }
    $tags += $risk.ToLower(); if ($userPermText -ne 'None') { $tags += 'userconsent' }; if ($unverified) { $tags += 'unverified' }; $tags += $(if ($isMicrosoft) { 'microsoft' } else { 'thirdparty' })
    # one column describing the application
    $appHtml = "<b>$(HtmlEnc $r['Application'])</b><br><span class='muted'>$(HtmlEnc $r['Publisher']) &middot; $(HtmlEnc $r['Type']) &middot; $(if ($r['Enabled for sign-in'] -eq $true) { 'enabled for sign-in' } else { 'sign-in disabled' }) &middot; $(if ($r['Assignment required'] -eq $true) { 'assignment required' } else { 'open to all users' }) &middot; $($r['Users/groups assigned']) user(s)/group(s) assigned</span>"
    # one column for every permission, first four inline, the rest in a dialog
    $groups = @()
    if ($appPermText -ne 'None') { $groups += @{ Label = 'Application permissions (run without a signed-in user)'; Items = @($appPermText -split ';\s*' | Where-Object { $_ }) } }
    if ($adminPermText -ne 'None') { $groups += @{ Label = 'Delegated permissions, admin-consented (act as any signed-in user)'; Items = @($adminPermText -split ';\s*' | Where-Object { $_ }) } }
    if ($userPermText -ne 'None') { $groups += @{ Label = 'Delegated permissions, user-consented (act as the users who approved it)'; Items = @($userPermText -split ';\s*' | Where-Object { $_ }) } }
    $flat = @(); foreach ($g in $groups) { foreach ($it in $g.Items) { $flat += @{ Group = $g.Label; Item = $it } } }
    if ($flat.Count -eq 0) { $permHtml = "<span class='muted'>None</span>" }
    else {
        $permHtml = "<ul class='perms'>" + ((@($flat | Select-Object -First 4) | ForEach-Object { "<li>$(HtmlEnc $_.Item)$(if ($_.Group -like 'Application*') { " <span class='muted'>(application)</span>" } elseif ($_.Group -like '*user-consented*') { " <span class='muted'>(user consent)</span>" })</li>" }) -join '') + '</ul>'
        $full = "<p class='muted'>$(HtmlEnc $r['Publisher']) &middot; approved: $(HtmlEnc $r['Approved by / when'])</p>" + (($groups | ForEach-Object { "<h5>$(HtmlEnc $_.Label) ($(@($_.Items).Count))</h5><ul>" + ((@($_.Items) | ForEach-Object { "<li>$(HtmlEnc $_)</li>" }) -join '') + '</ul>' }) -join '')
        $permHtml += "<button type='button' class='more' data-target='perm-$permIdx' data-title='$(HtmlEnc $r['Application']): $($flat.Count) permission(s)'>$(if ($flat.Count -gt 4) { "Show all $($flat.Count)" } else { 'Details' })</button><div id='perm-$permIdx' class='permfull' hidden>$full</div>"
    }
    $r['Risk'] = $risk
    $r['Permissions'] = "$($flat.Count) permission(s)"
    $r['_html'] = @{ 'Application' = $appHtml; 'Permissions' = $permHtml }
    $r['_tags'] = ($tags -join ' ')
    if ($is.Count) { $r['_issues'] = $is }
}
$AppFilters = @(@{ Label = 'High risk'; Tag = 'high' }, @{ Label = 'Medium risk'; Tag = 'medium' }, @{ Label = 'User-consented'; Tag = 'userconsent' }, @{ Label = 'Unverified publisher'; Tag = 'unverified' }, @{ Label = 'Third-party'; Tag = 'thirdparty' }, @{ Label = 'Microsoft'; Tag = 'microsoft' })
$appRowsSorted = @($appRows | Sort-Object @{ Expression = { @{ 'High' = 0; 'Medium' = 1; 'Low' = 2 }[[string]$_['Risk']] } }, @{ Expression = { [string]$_['Application'] } })
foreach ($r in $caSummaryRows) { if ($r['State'] -ne 'On') { $r['_issues'] = @{ 'State' = 'ca-not-enabled' } } }
foreach ($r in $authMethodRows) { if ($r['State'] -eq 'enabled' -and $r['Method'] -in @('Sms', 'Voice', 'Email')) { $r['_issues'] = @{ 'State' = 'auth-weak-method' } } }
foreach ($r in $profileRows) { if (-not $r['Included'] -or $r['Included'] -eq 'None') { $r['_issues'] = @{ 'Included' = 'policy-unassigned' } } }
foreach ($r in @($fwdMailboxRows)) { if ($r -and $r['External'] -eq 'Yes') { $r['_issues'] = @{ 'External' = 'forwarding' } } }
foreach ($r in @($fwdRuleRows))    { if ($r -and $r['External'] -eq 'Yes') { $r['_issues'] = @{ 'External' = 'forwarding' } } }
foreach ($r in @($casRows)) {
    if (-not $r) { continue }; $is = @{}
    if ($r['POP enabled'] -eq $true) { $is['POP enabled'] = 'pop-imap' }; if ($r['IMAP enabled'] -eq $true) { $is['IMAP enabled'] = 'pop-imap' }
    if ([string]$r['SMTP AUTH'] -like 'Enabled*') { $is['SMTP AUTH'] = 'smtp-auth' }
    if ($is.Count) { $r['_issues'] = $is }
}
$exoIssues = @{}
if ($exoCollected) {
    if ($exoKv['Default remote domain (*) allows auto-forward'] -eq $true) { $exoIssues['Default remote domain (*) allows auto-forward'] = 'remote-domain-forward' }
    if ([string]$exoKv['Outbound spam policy automatic forwarding'] -match '(?i)\bOn\b') { $exoIssues['Outbound spam policy automatic forwarding'] = 'spam-forward-on' }
    if ([int]$exoKv['Mailboxes forwarding to an external address'] -gt 0) { $exoIssues['Mailboxes forwarding to an external address'] = 'forwarding' }
    if (-not $SkipInboxRules -and [int]$exoKv['Inbox rules forwarding to an external address'] -gt 0) { $exoIssues['Inbox rules forwarding to an external address'] = 'forwarding' }
    if ($exoKv['Modern authentication (OAuth2) enabled'] -eq $false) { $exoIssues['Modern authentication (OAuth2) enabled'] = 'modern-auth-off' }
    if ($exoKv['SMTP AUTH disabled tenant-wide'] -eq $false) { $exoIssues['SMTP AUTH disabled tenant-wide'] = 'smtp-auth' }
    if ([int]$exoKv['Mailboxes with POP enabled'] -gt 0) { $exoIssues['Mailboxes with POP enabled'] = 'pop-imap' }
    if ([int]$exoKv['Mailboxes with IMAP enabled'] -gt 0) { $exoIssues['Mailboxes with IMAP enabled'] = 'pop-imap' }
}
$tapIssues    = @{}; if ($tap -and $tap.state -ne 'enabled') { $tapIssues['Temporary Access Pass state'] = 'tap-disabled' }
$devRegIssues = @{}; if ($devRegPolicy -and $devRegKv['Users may join devices to Entra ID'] -match 'all') { $devRegIssues['Users may join devices to Entra ID'] = 'device-join-open' }
$secDefIssues = @{}; if ($null -ne $secDefaults.isEnabled -and -not $secDefaults.isEnabled -and -not $mfaAllUsersEnforced.Count) { $secDefIssues['Security defaults'] = 'security-defaults-off' }
$appGovIssues = @{}
if ($userConsentSetting -notmatch '^Do not allow') { $appGovIssues['User consent setting'] = 'consent-allowed' }
if ($appGovKv['Admin consent request workflow'] -eq 'Disabled') { $appGovIssues['Admin consent request workflow'] = 'consent-workflow-off' }
if ($appGovKv['Users can register applications'] -eq $true) { $appGovIssues['Users can register applications'] = 'app-registration-open' }
$defenderKvIssues = @{}
if (-not $script:HasDefenderP1) { $defenderKvIssues['Defender for Office 365 licence'] = 'defender-not-licensed' }
if (-not $script:HasMDE) { $defenderKvIssues['Defender for Endpoint licence'] = 'mde-not-licensed' }
if ($defenderCollected -and $presetOn -eq 0) { $defenderKvIssues['Preset security policies enabled'] = 'preset-off' }
if ($atpGlobal -and ($atpGlobal.EnableATPForSPOTeamsODB -ne $true -or $atpGlobal.EnableSafeDocs -ne $true) -and $script:HasDefenderP1) { $defenderKvIssues['Defender for SharePoint, OneDrive and Teams'] = 'safeattach-spo-off' }
if ($script:HasDefenderP1 -and $defenderCollected -and -not $safeLinksActive) { $defenderKvIssues['Defender for Office 365 licence'] = 'safelinks-none' }
$userStatIssues = @{}
if ([string]$userStats['Enabled members with no MFA registered'] -match '^\d+$' -and [int]$userStats['Enabled members with no MFA registered'] -gt 0) { $userStatIssues['Enabled members with no MFA registered'] = 'mfa-not-registered' }
if ([string]$userStats['Enabled members with no sign-in in 90 days'] -match '^\d+$' -and [int]$userStats['Enabled members with no sign-in in 90 days'] -gt 0) { $userStatIssues['Enabled members with no sign-in in 90 days'] = 'user-stale' }
$deviceStatIssues = @{}
if ([int]$deviceStats['Entra Windows devices NOT enrolled in Intune'] -gt 0) { $deviceStatIssues['Entra Windows devices NOT enrolled in Intune'] = 'intune-coverage' }
if ($staleDeviceCount -gt 0) { $deviceStatIssues["Stale devices (no activity in $StaleDays days)"] = 'device-stale' }

# ---- Executive summary figures
$controlsRated = @($script:summaryRows | Where-Object { [string]$_.Status -notmatch '^(Not available|Not collected)' })
$controlsMet   = @($controlsRated | Where-Object { [string]$_.Status -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|None found)' })
$findCritHigh  = @($remediationSorted | Where-Object { $_.Priority -in @('Critical', 'High') }).Count
$findMedium    = @($remediationSorted | Where-Object { $_.Priority -eq 'Medium' }).Count
$enabledMemberCount = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' }).Count
$noMfaCount    = if (@($regDetails).Count) { @($userRows | Where-Object { " $($_['_tags']) " -like '* nomfa *' }).Count } else { $null }
$staleUserCount = if ($signInAvailable) { @($userRows | Where-Object { " $($_['_tags']) " -like '* stale *' }).Count } else { $null }
$notInIntuneCount = @($allDeviceRows | Where-Object { " $($_['_tags']) " -like '* notintune *' -and $_['Enabled'] -ne $false }).Count
$kpiHtml = New-Object System.Text.StringBuilder; $kpiMoreHtml = New-Object System.Text.StringBuilder
function Add-Kpi { param($Value, [string]$Label, [string]$Class = '', [switch]$More) $sb = $(if ($More) { $script:kpiMoreHtml } else { $script:kpiHtml }); [void]$sb.Append("<div class='kpi $Class'><div class='v'>$(HtmlEnc $Value)</div><div class='l'>$(HtmlEnc $Label)</div></div>") }
if ($null -ne $scoreCurrent) { Add-Kpi "$scorePct%" "Secure Score (goal $TargetScorePercent%)" $(if ($scoreCurrent -ge $scoreTarget) { 'good' } else { 'bad' }) } else { Add-Kpi 'n/a' 'Secure Score not available' 'mid' }
Add-Kpi "$($controlsMet.Count) of $($controlsRated.Count)" 'Controls in place' $(if ($controlsMet.Count -eq $controlsRated.Count) { 'good' } elseif ($controlsMet.Count * 2 -ge $controlsRated.Count) { 'mid' } else { 'bad' })
Add-Kpi $findCritHigh 'High-priority items due within 30 days' $(if ($findCritHigh -eq 0) { 'good' } else { 'bad' })
Add-Kpi -More $findMedium 'Medium findings' $(if ($findMedium -eq 0) { 'good' } else { 'mid' })
Add-Kpi -More $(if ($null -ne $passkeyPct) { "$passkeyPct%" } else { 'n/a' }) 'Users with phishing-resistant MFA (passkeys)' $(if ($null -eq $passkeyPct) { 'mid' } elseif ($passkeyPct -ge 90) { 'good' } elseif ($passkeyPct -gt 0) { 'mid' } else { 'bad' })
Add-Kpi $(if ($iocCollected) { "$iocUrgentCount / $iocReviewCount" } else { 'n/a' }) 'Warning signs of an intrusion: urgent today / review this week' $(if (-not $iocCollected) { 'mid' } elseif ($iocUrgentCount) { 'bad' } elseif ($iocReviewCount) { 'mid' } else { 'good' })
Add-Kpi -More $enabledMemberCount 'Enabled user accounts' ''
Add-Kpi -More $(if ($null -ne $noMfaCount) { $noMfaCount } else { 'n/a' }) 'Users without MFA registered' $(if ($null -eq $noMfaCount) { 'mid' } elseif ($noMfaCount -eq 0) { 'good' } else { 'bad' })
Add-Kpi -More $(if ($null -ne $staleUserCount) { $staleUserCount } else { 'n/a' }) "Stale users (no sign-in in $StaleDays days)" $(if ($null -eq $staleUserCount) { 'mid' } elseif ($staleUserCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi -More $gaCount 'Global Administrators' $(if ($gaCount -ge 2 -and $gaCount -le 4) { 'good' } else { 'mid' })
Add-Kpi -More $allDeviceRows.Count 'Devices (Entra ID and Intune)' ''
Add-Kpi -More $notInIntuneCount 'Enabled devices not in Intune' $(if ($notInIntuneCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi -More $staleDeviceCount "Stale devices ($StaleDays days)" $(if ($staleDeviceCount -eq 0) { 'good' } else { 'mid' })
$mfaProductShort = if ($mfaProduct) { ($mfaProduct -split ' \(')[0] } else { 'Microsoft Entra multifactor authentication' }
$execNarrative = "This report was produced on $($ReportDate.ToString('MMMM d, yyyy')) from the live configuration of the $(HtmlEnc $ClientLabel) Microsoft 365 tenant, read-only, by $(HtmlEnc $ctx.Account). " +
    "It measures $($controlsRated.Count) security controls that regulators and examiners ask about; <b>$($controlsMet.Count) are met</b>. " +
    $(if ($null -ne $scoreCurrent) { "Microsoft Secure Score is <b>$scoreCurrent of $scoreMax points ($scorePct%)</b> against a target of $scoreTarget points ($TargetScorePercent%)$(if ($scoreCompareText) { "; $scoreCompareText" }). " } else { 'Microsoft Secure Score could not be read. ' }) +
    "The remediation plan contains <b>$($remediationSorted.Count) findings</b> ($remediationCounts). " +
    $(if ($iocCollected) { "Possible indications of compromise in the last $AuditDays days: <b>$iocUrgentCount urgent</b> and $iocReviewCount to review (section 2). " } else { '' }) +
    $(if ($null -ne $passkeyPct) { "MFA is provided by $(HtmlEnc $mfaProductShort); <b>$passkeyPct%</b> of enabled users and $adminPhishRes of $($adminUsers.Count) administrators have phishing-resistant MFA (passkeys). " } else { '' }) +
    $(if ($enabledMemberCount) { "There are $enabledMemberCount enabled user accounts$(if ($null -ne $noMfaCount) { ", $noMfaCount without MFA registered" })$(if ($null -ne $staleUserCount) { ", $staleUserCount with no sign-in in $StaleDays days" }), $gaCount Global Administrator account(s), and $($allDeviceRows.Count) devices, of which $notInIntuneCount enabled device(s) are not managed by Intune." } else { '' })

# ----------------------------------------------------------------------------------------------
# Run history: a small JSON file per tenant in the output folder feeds the progress charts and the change log.
# It holds only the headline figures and control statuses, never user or device names.
# ----------------------------------------------------------------------------------------------
$historyName = ($ClientLabel -replace '[^A-Za-z0-9 ._()-]', '').Trim(); if (-not $historyName) { $historyName = 'Client' }
$historyPath = Join-Path $OutputFolder "$historyName - history.json"
$history = @()
if (-not $NoHistory -and (Test-Path $historyPath)) {
    try { $history = @(Get-Content $historyPath -Raw -ErrorAction Stop | ConvertFrom-Json) } catch { Write-Warning "Could not read the history file $historyPath : $($_.Exception.Message)" }
}
$controlMap = [ordered]@{}; foreach ($row in $script:summaryRows) { $controlMap[[string]$row.Control] = [string]$row.Status }
$snapshot = [ordered]@{
    Date = $ReportDate.ToString('yyyy-MM-dd HH:mm'); Version = $ScriptVersion
    SecureScore = $scoreCurrent; SecureScoreMax = $scoreMax; SecureScorePct = $scorePct; TargetPct = $TargetScorePercent
    Findings = [ordered]@{ Critical = @($remediationSorted | Where-Object { $_.Priority -eq 'Critical' }).Count; High = @($remediationSorted | Where-Object { $_.Priority -eq 'High' }).Count; Medium = $findMedium; Low = @($remediationSorted | Where-Object { $_.Priority -eq 'Low' }).Count; Total = $remediationSorted.Count }
    FindingNames = @($remediationSorted | ForEach-Object { [string]$_.Finding })
    Controls = $controlMap
    Users = [ordered]@{ Enabled = $enabledMemberCount; NoMfa = $noMfaCount; Stale = $staleUserCount; GlobalAdmins = $gaCount; Passkey = $(if ($null -ne $passkeyPct) { $passkeyUsers } else { $null }); PasskeyPct = $passkeyPct; AdminsPhishingResistant = $(if ($regEnabled.Count) { $adminPhishRes } else { $null }) }
    Devices = [ordered]@{ Total = $allDeviceRows.Count; NotInIntune = $notInIntuneCount; Stale = $staleDeviceCount }
    Ioc = [ordered]@{ Urgent = $(if ($iocCollected) { $iocUrgentCount } else { $null }); Review = $(if ($iocCollected) { $iocReviewCount } else { $null }) }
    Sharing = [ordered]@{ TeamsWithGuests = $(if (Get-GraphFailure 'teams') { $null } else { $teamsWithGuests }); ExternalShares = $(if ($spoActivity.Count -or $odActivity.Count) { $extShareTotal } else { $null }); Guests = @($users | Where-Object { $_.userType -eq 'Guest' }).Count; Anyone = $(if ($SkipShareScan -and -not $ualCollected) { $null } else { $shareAnyone + $anonLinkCount }) }
}
$previous = if ($history.Count) { $history[-1] } else { $null }
function Get-StatusRank { param([string]$s)
    if ($s -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded|None found)') { return 2 }
    if ($s -match '^(Not available|Not collected)') { return -1 }
    if ($s -match '^(Partial|Allowed|All users allowed|No Windows|EOP only|No email domains|No devices onboarded|Not licensed|Guests present|In progress|\d+ to review|\d+ (assignments|policies))') { return 1 }
    return 0
}
function Get-JsonProp { param($obj, [string]$name) if ($null -eq $obj) { return $null } ; $p = $obj.PSObject.Properties[$name]; if ($p) { return $p.Value } ; return $null }
$changeRows = @(); $improvedCount = 0; $regressedCount = 0; $resolvedCount = 0; $newCount = 0
if ($previous) {
    $prevPct = Get-JsonProp $previous 'SecureScorePct'
    if ($null -ne $prevPct -and $null -ne $scorePct -and [int]$prevPct -ne [int]$scorePct) {
        $dir = if ($scorePct -gt $prevPct) { 'Improved' } else { 'Regressed' }
        $changeRows += [ordered]@{ 'Area' = 'Microsoft Secure Score'; 'Change' = "Score moved from $prevPct% to $scorePct%"; 'Before' = "$(Get-JsonProp $previous 'SecureScore') of $(Get-JsonProp $previous 'SecureScoreMax')"; 'After' = "$scoreCurrent of $scoreMax"; 'Direction' = $dir }
        if ($dir -eq 'Improved') { $improvedCount++ } else { $regressedCount++ }
    }
    $prevControls = Get-JsonProp $previous 'Controls'
    foreach ($c in $controlMap.Keys) {
        $before = [string](Get-JsonProp $prevControls $c)
        if ($before -and $before -ne $controlMap[$c]) {
            $rb = Get-StatusRank $before; $ra = Get-StatusRank $controlMap[$c]
            $dir = if ($ra -gt $rb) { 'Improved' } elseif ($ra -lt $rb) { 'Regressed' } else { 'Changed' }
            if ($dir -eq 'Improved') { $improvedCount++ } elseif ($dir -eq 'Regressed') { $regressedCount++ }
            $changeRows += [ordered]@{ 'Area' = 'Control'; 'Change' = $c; 'Before' = $before; 'After' = $controlMap[$c]; 'Direction' = $dir }
        }
    }
    $prevNames = @(Get-JsonProp $previous 'FindingNames'); $curNames = @($snapshot.FindingNames)
    foreach ($fn in ($prevNames | Where-Object { $_ -and $curNames -notcontains $_ } | Select-Object -First 40)) { $resolvedCount++; $changeRows += [ordered]@{ 'Area' = 'Finding'; 'Change' = $fn; 'Before' = 'Open'; 'After' = 'Resolved'; 'Direction' = 'Resolved' } }
    foreach ($fn in ($curNames | Where-Object { $_ -and $prevNames -notcontains $_ } | Select-Object -First 40)) { $newCount++; $changeRows += [ordered]@{ 'Area' = 'Finding'; 'Change' = $fn; 'Before' = 'Not present'; 'After' = 'Open'; 'Direction' = 'New' } }
    $prevUsers = Get-JsonProp $previous 'Users'; $prevDevices = Get-JsonProp $previous 'Devices'
    $countChecks = @(
        @{ Area = 'Users'; Label = 'Users without MFA registered'; Before = (Get-JsonProp $prevUsers 'NoMfa'); After = $noMfaCount },
        @{ Area = 'Users'; Label = "Stale users ($StaleDays days)"; Before = (Get-JsonProp $prevUsers 'Stale'); After = $staleUserCount },
        @{ Area = 'Users'; Label = 'Global Administrators'; Before = (Get-JsonProp $prevUsers 'GlobalAdmins'); After = $gaCount },
        @{ Area = 'Users'; Label = 'Users with phishing-resistant MFA (passkeys)'; Before = (Get-JsonProp $prevUsers 'Passkey'); After = $(if ($null -ne $passkeyPct) { $passkeyUsers } else { $null }); MoreIsBetter = $true },
        @{ Area = 'Users'; Label = 'Administrators with phishing-resistant MFA'; Before = (Get-JsonProp $prevUsers 'AdminsPhishingResistant'); After = $(if ($regEnabled.Count) { $adminPhishRes } else { $null }); MoreIsBetter = $true },
        @{ Area = 'Devices'; Label = 'Enabled devices not in Intune'; Before = (Get-JsonProp $prevDevices 'NotInIntune'); After = $notInIntuneCount },
        @{ Area = 'Devices'; Label = "Stale devices ($StaleDays days)"; Before = (Get-JsonProp $prevDevices 'Stale'); After = $staleDeviceCount },
        @{ Area = 'Compromise indicators'; Label = 'Urgent indicators'; Before = (Get-JsonProp (Get-JsonProp $previous 'Ioc') 'Urgent'); After = $(if ($iocCollected) { $iocUrgentCount } else { $null }) },
        @{ Area = 'Compromise indicators'; Label = 'Indicators to review'; Before = (Get-JsonProp (Get-JsonProp $previous 'Ioc') 'Review'); After = $(if ($iocCollected) { $iocReviewCount } else { $null }) },
        @{ Area = 'Sharing'; Label = 'Teams with guests'; Before = (Get-JsonProp (Get-JsonProp $previous 'Sharing') 'TeamsWithGuests'); After = $(if (Get-GraphFailure 'teams') { $null } else { $teamsWithGuests }) },
        @{ Area = 'Sharing'; Label = 'Files shared externally (30 days)'; Before = (Get-JsonProp (Get-JsonProp $previous 'Sharing') 'ExternalShares'); After = $(if ($spoActivity.Count -or $odActivity.Count) { $extShareTotal } else { $null }) },
        @{ Area = 'Sharing'; Label = 'Guest accounts'; Before = (Get-JsonProp (Get-JsonProp $previous 'Sharing') 'Guests'); After = @($users | Where-Object { $_.userType -eq 'Guest' }).Count }
    )
    foreach ($k in $countChecks) {
        if ($null -ne $k.Before -and $null -ne $k.After -and [int]$k.Before -ne [int]$k.After) {
            $dir = if ([int]$k.After -lt [int]$k.Before) { 'Improved' } else { 'Regressed' }
            if ($k.MoreIsBetter) { $dir = if ($dir -eq 'Improved') { 'Regressed' } else { 'Improved' } }
            if ($k.Label -eq 'Global Administrators') { $dir = 'Changed' } elseif ($dir -eq 'Improved') { $improvedCount++ } else { $regressedCount++ }
            $changeRows += [ordered]@{ 'Area' = $k.Area; 'Change' = $k.Label; 'Before' = $k.Before; 'After' = $k.After; 'Direction' = $dir }
        }
    }
}
$history = @(@($history) + @([pscustomobject]$snapshot) | Select-Object -Last 60)
if (-not $NoHistory) {
    try { $history | ConvertTo-Json -Depth 8 | Set-Content -Path $historyPath -Encoding UTF8; Write-Host "History: $historyPath ($($history.Count) run(s))" -ForegroundColor DarkGray }
    catch { Write-Warning "Could not write the history file $historyPath : $($_.Exception.Message)" }
}

# ---- Inline SVG charts (no external libraries; one measure per chart)
function Get-ChartXLabels { param([int]$Count) if ($Count -le 6) { return 0..($Count - 1) } ; $step = [math]::Ceiling($Count / 6); return @(0..($Count - 1) | Where-Object { $_ % $step -eq 0 -or $_ -eq $Count - 1 }) }
function New-SvgTrendChart {
    # One series over the run dates, with an optional dashed target line. $Points: array of @{ Label; Value }.
    param([string]$Title, $Points, [double]$Max = 100, $Target, [string]$Unit = '%', [string]$Color = '#005D9C')
    $w = 600; $h = 240; $l = 46; $r = 40; $top = 30; $b = 48; $pw = $w - $l - $r; $ph = $h - $top - $b
    $pts = @($Points | Where-Object { $null -ne $_.Value }); $cnt = $pts.Count
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $w $h' width='100%' role='img' aria-label='$(HtmlEnc $Title)'><title>$(HtmlEnc $Title)</title>")
    [void]$sb.Append("<text x='$l' y='16' font-size='12' font-weight='bold' fill='#000'>$(HtmlEnc $Title)</text>")
    foreach ($g in 0, 25, 50, 75, 100) { $y = [math]::Round($top + $ph - ($g / 100) * $ph, 1); $v = [math]::Round($Max * $g / 100); [void]$sb.Append("<line x1='$l' y1='$y' x2='$($w - $r)' y2='$y' stroke='#e6e6e6' stroke-width='1'/><text x='$($l - 6)' y='$($y + 4)' font-size='10' text-anchor='end' fill='#555'>$v$Unit</text>") }
    if ($cnt -eq 0) { [void]$sb.Append("<text x='$($l + $pw / 2)' y='$($top + $ph / 2)' font-size='11' text-anchor='middle' fill='#555'>No data yet</text></svg>"); return $sb.ToString() }
    if ($null -ne $Target -and $Max -gt 0) { $ty = [math]::Round($top + $ph - ([double]$Target / $Max) * $ph, 1); [void]$sb.Append("<line x1='$l' y1='$ty' x2='$($w - $r)' y2='$ty' stroke='#1e7b34' stroke-width='1.5' stroke-dasharray='5,4'/><text x='$($w - $r)' y='$($ty - 4)' font-size='10' text-anchor='end' fill='#1e7b34'>Target $Target$Unit</text>") }
    $coords = @(); for ($i = 0; $i -lt $cnt; $i++) { $x = if ($cnt -eq 1) { $l + $pw / 2 } else { $l + $i * ($pw / ($cnt - 1)) }; $y = $top + $ph - ([math]::Min([double]$pts[$i].Value, $Max) / $Max) * $ph; $coords += ,@([math]::Round($x, 1), [math]::Round($y, 1)) }
    if ($cnt -gt 1) { [void]$sb.Append("<polyline fill='none' stroke='$Color' stroke-width='2' stroke-linejoin='round' points='$(($coords | ForEach-Object { "$($_[0]),$($_[1])" }) -join ' ')'/>") }
    $labelIdx = Get-ChartXLabels $cnt
    for ($i = 0; $i -lt $cnt; $i++) {
        [void]$sb.Append("<circle cx='$($coords[$i][0])' cy='$($coords[$i][1])' r='4' fill='$Color' stroke='#fff' stroke-width='2'><title>$(HtmlEnc $pts[$i].Label): $($pts[$i].Value)$Unit</title></circle>")
        if ($labelIdx -contains $i) { $anchor = if ($cnt -gt 1 -and $i -eq 0) { 'start' } elseif ($cnt -gt 1 -and $i -eq $cnt - 1) { 'end' } else { 'middle' }; [void]$sb.Append("<text x='$($coords[$i][0])' y='$($h - $b + 16)' font-size='10' text-anchor='$anchor' fill='#555'>$(HtmlEnc ([string]$pts[$i].Label).Substring(0, [math]::Min(10, ([string]$pts[$i].Label).Length)))</text>") }
    }
    $last = $coords[$cnt - 1]; [void]$sb.Append("<text x='$($last[0] + 8)' y='$($last[1] + 4)' font-size='11' font-weight='bold' text-anchor='start' fill='#000'>$($pts[$cnt - 1].Value)$Unit</text>")
    [void]$sb.Append('</svg>'); return $sb.ToString()
}
function New-SvgGroupedBars {
    # Grouped bars per run, up to three series; one count axis. $Runs: array of @{ Label; Values = @(a, b, c) }.
    param([string]$Title, $Runs, [string[]]$Series, [string[]]$Colors)
    $w = 600; $h = 240; $l = 40; $r = 18; $top = 30; $b = 48; $pw = $w - $l - $r; $ph = $h - $top - $b
    $runs = @($Runs); $cnt = $runs.Count; $ns = $Series.Count
    $max = 1; foreach ($run in $runs) { foreach ($v in @($run.Values)) { if ([double]$v -gt $max) { $max = [double]$v } } }
    $step = [math]::Max(1, [math]::Ceiling($max / 4)); $max = $step * 4
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append("<svg viewBox='0 0 $w $h' width='100%' role='img' aria-label='$(HtmlEnc $Title)'><title>$(HtmlEnc $Title)</title>")
    [void]$sb.Append("<text x='$l' y='16' font-size='12' font-weight='bold' fill='#000'>$(HtmlEnc $Title)</text>")
    $lx = $w - $r; for ($s = $ns - 1; $s -ge 0; $s--) { $lx -= (7 * $Series[$s].Length + 26); [void]$sb.Append("<rect x='$lx' y='7' width='10' height='10' rx='2' fill='$($Colors[$s])'/><text x='$($lx + 14)' y='16' font-size='10' fill='#333'>$(HtmlEnc $Series[$s])</text>") }
    for ($g = 0; $g -le 4; $g++) { $y = [math]::Round($top + $ph - ($g / 4) * $ph, 1); [void]$sb.Append("<line x1='$l' y1='$y' x2='$($w - $r)' y2='$y' stroke='#e6e6e6' stroke-width='1'/><text x='$($l - 6)' y='$($y + 4)' font-size='10' text-anchor='end' fill='#555'>$($g * $step)</text>") }
    if ($cnt -eq 0) { [void]$sb.Append("<text x='$($l + $pw / 2)' y='$($top + $ph / 2)' font-size='11' text-anchor='middle' fill='#555'>No data yet</text></svg>"); return $sb.ToString() }
    $gw = $pw / $cnt; $bw = [math]::Max(4, [math]::Min(28, ($gw * 0.72) / $ns - 2)); $labelIdx = Get-ChartXLabels $cnt
    for ($i = 0; $i -lt $cnt; $i++) {
        $gx = $l + $i * $gw + ($gw - ($bw + 2) * $ns) / 2
        for ($s = 0; $s -lt $ns; $s++) {
            $v = [double](@($runs[$i].Values)[$s]); $bh = [math]::Round(($v / $max) * $ph, 1); $x = [math]::Round($gx + $s * ($bw + 2), 1); $y = [math]::Round($top + $ph - $bh, 1)
            if ($bh -gt 0) { [void]$sb.Append("<rect x='$x' y='$y' width='$bw' height='$bh' rx='2' fill='$($Colors[$s])'><title>$(HtmlEnc $runs[$i].Label), $(HtmlEnc $Series[$s]): $v</title></rect>") }
            if ($i -eq $cnt - 1) { [void]$sb.Append("<text x='$($x + $bw / 2)' y='$($y - 4)' font-size='10' font-weight='bold' text-anchor='middle' fill='#000'>$v</text>") }
        }
        if ($labelIdx -contains $i) { [void]$sb.Append("<text x='$([math]::Round($l + $i * $gw + $gw / 2, 1))' y='$($h - $b + 16)' font-size='10' text-anchor='middle' fill='#555'>$(HtmlEnc ([string]$runs[$i].Label).Substring(0, [math]::Min(10, ([string]$runs[$i].Label).Length)))</text>") }
    }
    [void]$sb.Append('</svg>'); return $sb.ToString()
}
$scoreTrend = New-SvgTrendChart -Title 'Microsoft Secure Score (% of maximum) by report' -Points @($history | ForEach-Object { @{ Label = ([string](Get-JsonProp $_ 'Date')).Substring(0, 10); Value = (Get-JsonProp $_ 'SecureScorePct') } }) -Max 100 -Target $TargetScorePercent -Unit '%'
$findingTrend = New-SvgGroupedBars -Title 'Open findings by priority, by report' -Runs @($history | ForEach-Object { $fd = Get-JsonProp $_ 'Findings'; @{ Label = ([string](Get-JsonProp $_ 'Date')).Substring(0, 10); Values = @(([int](Get-JsonProp $fd 'Critical') + [int](Get-JsonProp $fd 'High')), [int](Get-JsonProp $fd 'Medium'), [int](Get-JsonProp $fd 'Low')) } }) -Series @('Critical + High', 'Medium', 'Low') -Colors @('#b00020', '#B8860B', '#005D9C')
$progressHtml = if ($previous) {
    "<p class='meta'>Compared with the previous report of $(HtmlEnc (Get-JsonProp $previous 'Date')) (script $(HtmlEnc (Get-JsonProp $previous 'Version'))): <b>$improvedCount improvement(s)</b>, $regressedCount regression(s), $resolvedCount finding(s) resolved, $newCount new finding(s). $($history.Count) reports in the history.</p>" +
    $(if ($changeRows.Count) { New-HtmlTable $changeRows -Columns 'Area','Change','Before','After','Direction' } else { "<p class='empty'>No measured control, finding or headline count changed since the previous report.</p>" })
} else { "<p class='meta'>This is the first report kept for this tenant. From the second run onward this section lists every control, finding and headline count that changed, and the charts show the trend. The history (headline figures only, no names) is stored next to the report in <b>$(HtmlEnc $historyName) - history.json</b>; keep it with the reports.</p>" }

function Get-PortalLabel {
    # Name the admin center a link opens, so the reader knows where the setting lives (not everything is in the Defender portal).
    param([string]$Url)
    switch -Regex ($Url) {
        'security\.microsoft\.com'                       { return 'Open in Defender portal' }
        'entra\.microsoft\.com|aad\.portal|portal\.azure' { return 'Open in Entra admin center' }
        'intune\.microsoft\.com|endpoint\.microsoft\.com' { return 'Open in Intune admin center' }
        'admin\.exchange\.microsoft\.com'                { return 'Open in Exchange admin center' }
        'admin\.teams\.microsoft\.com'                   { return 'Open in Teams admin center' }
        'sharepoint'                                       { return 'Open in SharePoint admin center' }
        'compliance\.microsoft\.com|purview'              { return 'Open in Purview' }
        'learn\.microsoft\.com'                           { return 'Microsoft guidance (DNS change)' }
        'admin\.microsoft\.com'                           { return 'Open in Microsoft 365 admin center' }
        default                                            { return 'Open setting' }
    }
}
$roadmapHtml = New-Object System.Text.StringBuilder
[void]$roadmapHtml.Append("<div class='roadmap'>")
foreach ($ph in $PhaseOrder) {
    $phaseItems = @($remediationSorted | Where-Object { $_.Phase -eq $ph })
    [void]$roadmapHtml.Append("<div class='phase'><h4>$(HtmlEnc $ph)<span class='n'>$($phaseItems.Count)</span></h4>")
    if ($phaseItems.Count) {
        [void]$roadmapHtml.Append('<ol>' + ((@($phaseItems | Select-Object -First 12) | ForEach-Object { "<li><a href='#rem-$($_.Ref)'>$(HtmlEnc $_.Finding)</a><span class='muted'> $(HtmlEnc $_.Ref)$(if ($_.Playbook) { " &middot; $(HtmlEnc $_.Playbook.Owner) &middot; $(HtmlEnc $_.Playbook.Effort)" })$(if ($_.Points) { " &middot; $($_.Points) pts" })</span></li>" }) -join '') + '</ol>')
        if ($phaseItems.Count -gt 12) { [void]$roadmapHtml.Append("<p class='muted'>and $($phaseItems.Count - 12) more in the table below.</p>") }
    } else { [void]$roadmapHtml.Append("<p class='empty'>Nothing scheduled.</p>") }
    [void]$roadmapHtml.Append('</div>')
}
[void]$roadmapHtml.Append('</div>')
$remHtml = New-Object System.Text.StringBuilder
[void]$remHtml.Append("<table class='grid' id='remTable'><thead><tr><th style='width:2%'>#</th><th style='width:5%'>Ref</th><th style='width:5%'>Priority</th><th style='width:7%'>Phase</th><th style='width:13%'>Finding</th><th style='width:14%'>Current state</th><th>What to do, step by step</th><th style='width:11%'>Owner and effort</th><th style='width:4%'>Points</th><th style='width:7%'>Standards</th><th style='width:7%'>Where</th><th style='width:6%'>Ticket</th></tr></thead><tbody>")
$remIdx = 0
foreach ($item in $remediationSorted) {
    $remIdx++
    $cls = switch ($item.Priority) { 'Critical' { 'warn' } 'High' { 'warn' } 'Medium' { 'mid' } default { '' } }
    $detail = if ($item.Url) { "<a href='$(HtmlEnc $item.Url)' target='_blank' rel='noopener'>$(Get-PortalLabel $item.Url)</a>" } elseif ($SectionAnchors[$item.Section]) { "<a href='#$($SectionAnchors[$item.Section])'>$(HtmlEnc $item.Section)</a>" } else { HtmlEnc $item.Section }
    if ($item.Url -and $SectionAnchors[$item.Section] -and $item.Section -ne 'Secure Score') { $detail += "<br><a href='#$($SectionAnchors[$item.Section])'>Evidence: $(HtmlEnc $item.Section)</a>" }
    $pb = $item.Playbook
    $how = "<div>$(HtmlEnc $item.Action)</div>"
    if ($pb) { $how += "<details class='steps'><summary>Step by step ($(@($pb.Steps).Count) steps)</summary><ol>" + ((@($pb.Steps) | ForEach-Object { "<li>$(HtmlEnc $_)</li>" }) -join '') + "</ol><p><b>Prerequisite:</b> $(HtmlEnc $pb.Prereq)<br><b>Done when:</b> $(HtmlEnc $pb.Verify)</p></details>" }
    $ownerCell = if ($pb) { "$(HtmlEnc $pb.Owner)<br><span class='muted'>$(HtmlEnc $pb.Effort)</span>$(if ($item.Impact) { "<br><span class='muted'>$(HtmlEnc $item.Impact)</span>" })" } else { HtmlEnc $item.Impact }
    $fwCell = if ($item.Frameworks) { "<details class='steps'><summary>Show</summary><p>$((HtmlEnc $item.Frameworks) -replace ' \| ', '<br>')</p></details>" } else { '' }
    [void]$remHtml.Append("<tr id='rem-$($item.Ref)' data-tags='$($item.Priority.ToLower())'><td>$remIdx</td><td class='ref'>$(HtmlEnc $item.Ref)</td><td class='$cls'>$(HtmlEnc $item.Priority)</td><td>$(HtmlEnc $item.Phase)</td><td><b>$(HtmlEnc $item.Finding)</b></td><td>$(HtmlEnc $item.State)</td><td>$how</td><td>$ownerCell</td><td>$(if ($null -ne $item.Points) { HtmlEnc $item.Points })</td><td>$fwCell</td><td>$detail</td><td><button type='button' class='tk' data-i='$($remIdx - 1)' title='Preview the ticket text for this finding, then copy it'>View ticket</button></td></tr>")
}
[void]$remHtml.Append('</tbody></table>')
if ($remIdx -eq 0) { $remHtml = New-Object System.Text.StringBuilder; [void]$remHtml.Append("<p class='empty'>No findings. Every measured control passed and every Secure Score action is implemented.</p>") }

# ---- Tickets: the same findings as plain rows for Copy ticket / Copy all / CSV download (in the report) and a CSV file next to the report
$safeClient = ($ClientLabel -replace '[^A-Za-z0-9 ._()-]', '').Trim()
if (-not $safeClient) { $safeClient = 'Client' }
$ticketCsvName = "$safeClient - Remediation tickets $($ReportDate.ToString('yyyy-MM-dd')).csv"
$ticketRows = @($remediationSorted | ForEach-Object {
    [pscustomobject]([ordered]@{
        'Ref' = $_.Ref; 'Priority' = $_.Priority; 'Finding' = $_.Finding; 'Current state' = $_.State; 'Recommended action' = $_.Action
        'Why it matters' = $_.Impact; 'Secure Score points' = $(if ($null -ne $_.Points) { [string]$_.Points } else { '' })
        'Report section' = $_.Section; 'Link' = $(if ($_.Url) { [string]$_.Url } else { '' })
        'Phase' = $_.Phase; 'Owner' = $(if ($_.Playbook) { [string]$_.Playbook.Owner } else { '' }); 'Effort' = $(if ($_.Playbook) { [string]$_.Playbook.Effort } else { '' })
        'Prerequisite' = $(if ($_.Playbook) { [string]$_.Playbook.Prereq } else { '' })
        'Steps' = $(if ($_.Playbook) { $i = 0; (@($_.Playbook.Steps) | ForEach-Object { $i++; "$i. $_" }) -join "`n" } else { '' })
        'Done when' = $(if ($_.Playbook) { [string]$_.Playbook.Verify } else { '' }); 'Standards' = $_.Frameworks
    })
})
$ticketJson = if ($ticketRows.Count) { (ConvertTo-Json -InputObject $ticketRows -Compress -Depth 3) -replace '</', '<\/' } else { '[]' }
$ticketMeta = ConvertTo-Json -InputObject ([pscustomobject]@{ Client = $ClientLabel; Date = $ReportDate.ToString('yyyy-MM-dd'); Version = $ScriptVersion; File = $ticketCsvName }) -Compress
$ticketBar = if ($remIdx) {
    "<div class='filters' data-table='remTable'><button class='on' data-tag='all'>All</button><button data-tag='critical'>Critical</button><button data-tag='high'>High</button><button data-tag='medium'>Medium</button><button data-tag='low'>Low</button><input type='search' placeholder='Search findings'><span class='n'></span></div>" +
    "<div class='ticketbar'><button type='button' id='tkCsv'>Download tickets (CSV)</button><button type='button' id='tkAll'>Copy all tickets</button><span class='hint'>Each finding has a stable reference (AH-xxxxxx) that stays the same from report to report. Use <b>View ticket</b> on a row to preview the ticket text (steps, owner, standards) and copy it, <b>Copy all tickets</b>, or the CSV to hand findings to your IT provider or MSP as work items; the next report's <a href='#progress'>Progress since last report</a> shows which were resolved. The same CSV is saved next to this report: <a href='$(HtmlEnc ([uri]::EscapeDataString($ticketCsvName)))'>$(HtmlEnc $ticketCsvName)</a>.</span></div>" +
    "<div id='tkModal' class='modal' hidden><div class='modal-box' role='dialog' aria-modal='true' aria-labelledby='tkTitle'><div class='modal-head'><b id='tkTitle'>Ticket</b><button type='button' id='tkClose'>Close</button></div><pre id='tkText'></pre><div class='modal-actions'><button type='button' id='tkCopy'>Copy ticket</button><span class='hint'>Paste into your ticketing system or an email to your IT provider or MSP. Esc closes.</span></div></div></div>"
} else { '' }
$summaryTable = New-Object System.Text.StringBuilder
[void]$summaryTable.Append("<table class='grid' id='tbl-controls'><thead><tr><th style='width:24%'>Control</th><th style='width:11%'>Result</th><th style='width:15%'>Status</th><th>Evidence</th><th style='width:11%'>Detail</th></tr></thead><tbody>")
foreach ($row in $script:summaryRows) {
    $cls = Get-StatusKind $row.Status
    if ($cls -eq 'info') { $cls = '' }
    $anchor = $SectionAnchors[$row.Section]
    $sumIcon = ''
    if ($SummaryIssueMap.ContainsKey($row.Control)) {
        $st = [string]$row.Status
        if ($st -match '^(Not available|Not collected)') {
            if ($row.Control -match 'Intune' -and $script:HasIntune -eq $false) { $sumIcon = New-InfoIcon 'license-intune' }
            elseif ($row.Control -match 'MFA registered' -and $script:HasEntraP1 -eq $false) { $sumIcon = New-InfoIcon 'license-p1' }
        } elseif ($cls -in @('warn', 'mid')) {
            $iid = $SummaryIssueMap[$row.Control]
            if ($row.Control -like 'Email protection*' -and $st -like 'Licensed*') { $iid = 'preset-off' }
            if ($row.Control -like 'Endpoint protection*' -and $st -eq 'Attention needed') { $iid = 'mde-high-risk' }
            if ($row.Control -like 'Indications of compromise*' -and $st -match 'urgent') { $iid = 'ioc-urgent' }
            if ($row.Control -like 'SharePoint and OneDrive*' -and $st -like 'Anyone links*') { $iid = 'spo-anyone-links' }
            if ($row.Control -like 'Teams guest*' -and $st -like 'Guests present*') { $iid = 'teams-guests' }
            if ($row.Control -like 'Endpoint protection*' -and $st -eq 'No devices onboarded') { $iid = 'mde-not-onboarded' }
            $sumIcon = New-InfoIcon $iid
        }
    }
    [void]$summaryTable.Append("<tr data-tags='$(Get-StatusKind $row.Status)'><td><b>$(HtmlEnc $row.Control)</b></td><td>$(New-StatusBadge $row.Status -WordOnly)</td><td class='$cls'>$(HtmlEnc $row.Status)$sumIcon</td><td>$(HtmlEnc $row.Evidence)</td><td><a href='#$anchor'>$(HtmlEnc $row.Section)</a></td></tr>")
}
[void]$summaryTable.Append('</tbody></table>')
# Audit Hat logo: embedded as a data URI when AuditHat-Logo.png sits next to the script, so the report stays one file.
$logoPath = Join-Path $PSScriptRoot 'AuditHat-Logo.png'
$logoImg = ''
if (Test-Path $logoPath) {
    try { $logoImg = "<img class='logo' src='data:image/png;base64,$([Convert]::ToBase64String([IO.File]::ReadAllBytes($logoPath)))' alt='Audit Hat'>" }
    catch { Write-Verbose "Logo could not be embedded: $($_.Exception.Message)" }
}
$brandLink = "<a href='https://audithat.com' target='_blank' rel='noopener'>AuditHat.com</a>"
$navHtml = New-Object System.Text.StringBuilder
[void]$navHtml.Append("$(if ($logoImg) { "<a href='https://audithat.com' target='_blank' rel='noopener'>$logoImg</a>" })<h1>$(HtmlEnc $ClientName)</h1><div class='sub'>$(if ($DefaultDomain) { "$(HtmlEnc $DefaultDomain)<br>" })Microsoft 365 Audit Evidence<br>$($ReportDate.ToString('MMMM d, yyyy'))<br>Provided by $brandLink<br>Script version $ScriptVersion &middot; <a href='$RepoUrl' target='_blank' rel='noopener'>GitHub</a><br>Companion: <a href='$AdRepoUrl' target='_blank' rel='noopener'>AD Governance Report</a></div>")
foreach ($s in $sections) {
    [void]$navHtml.Append("<a href='#$($s.Id)'>$(HtmlEnc $s.Title)</a>")
    foreach ($sub in @($s.Sub | Where-Object { $_ })) { [void]$navHtml.Append("<a class='l2' href='#$($sub.Id)'>$(HtmlEnc $sub.Title)</a>") }
}

# add anchors to the platform headings in the MDM section
$complianceHtmlText = $complianceHtml.ToString()
foreach ($grp in $policyGroups) {
    $anchor = 'mdm-' + ($grp.Name -replace '[^A-Za-z0-9]', '')
    $complianceHtmlText = $complianceHtmlText.Replace("<h3 class='platform'>$(HtmlEnc $grp.Name)</h3>", "<h3 class='platform' id='$anchor'>$(HtmlEnc $grp.Name)</h3>")
}

$mdmSection = @"
<section id='mdm'><h2>6. MDM Compliance Policies</h2>
<p class='meta'>Source: Microsoft Graph deviceManagement/deviceCompliancePolicies, exported $($ReportDate.ToString('yyyy-MM-dd HH:mm')). Values are read directly from the tenant configuration.</p>
$complianceHtmlText
</section>
"@

$caSection = @"
<section id='ca'><h2>9. Conditional Access</h2>
<p class='meta'>Source: Microsoft Graph identity/conditionalAccess/policies, exported $($ReportDate.ToString('yyyy-MM-dd HH:mm')). Enabled policies: $($enabledCa.Count); report-only: $($reportOnlyCa.Count); disabled: $(@($caPolicies).Count - $enabledCa.Count - $reportOnlyCa.Count).</p>
<h3>Policy summary</h3>
$(New-HtmlTable $caSummaryRows -Columns 'Policy','State','Created','Modified' -EmptyText 'No Conditional Access policies exist in this tenant.')
<h3>Policies that require MFA</h3>
$(New-HtmlTable $mfaRows -Columns 'Policy','State','Scope','Apps','Control' -EmptyText 'No Conditional Access policy requires multi-factor authentication.')
<h3>Policies that block legacy authentication</h3>
$(New-HtmlTable $legacyRows -Columns 'Policy','State','Scope','Client app types','Control' -EmptyText 'No Conditional Access policy blocks legacy authentication clients (Exchange ActiveSync, other clients).')
<h3>Policies that require an Entra joined or compliant device</h3>
$(New-HtmlTable $deviceCaRows -Columns 'Policy','State','Scope','Apps','Platforms','Device filter','Controls' -EmptyText 'No Conditional Access policy requires a compliant or Entra joined device.')
<h3>Named locations</h3>
$(New-HtmlTable $namedLocRows -EmptyText 'No named locations are defined.')
<h3>Policy detail</h3>
$($caHtml.ToString())
</section>
"@

$exoDetailHtml = ''
if ($exoCollected) {
    $exoDetailHtml = @"
<h3>Remote domains</h3>
$(New-HtmlTable $remoteDomainRows -Columns 'Remote domain','Domain','Auto-forward allowed','Auto-reply allowed')
<h3>Mailboxes with forwarding configured</h3>
$(New-HtmlTable $fwdMailboxRows -Columns 'Mailbox','Forwarding SMTP address','Forwarding address (internal)','External','Keep copy in mailbox' -EmptyText 'No mailbox has mailbox-level forwarding configured.')
<h3>Inbox rules that forward or redirect mail</h3>
$(if ($SkipInboxRules) { "<p class='empty'>Inbox rules were not scanned (-SkipInboxRules).</p>" } else { New-HtmlTable $fwdRuleRows -Columns 'Mailbox','Rule','Enabled','Forwards to','External' -EmptyText 'No inbox rules forward or redirect mail.' })
<h3>Transport rules that redirect, copy, or bcc mail</h3>
$(New-HtmlTable $fwdTransportRows -Columns 'Rule','State','Redirect to','Bcc to','Copy / add to','Mode' -EmptyText 'No transport rules redirect, copy, or bcc mail.')
<h3>Mailboxes with POP, IMAP, or SMTP AUTH enabled</h3>
$(New-HtmlTable $casRows -Columns 'Mailbox','POP enabled','IMAP enabled','SMTP AUTH','ActiveSync','MAPI' -EmptyText 'No mailbox has POP, IMAP, or a per-mailbox SMTP AUTH override enabled.')
"@
}

$iocSectionHtml = @"
<section id='ioc'><h2>2. Review These Now: Possible Indications of Compromise (last $AuditDays days)</h2>
<p class='meta'><b>Read this section first.</b> Nothing in this section is a confirmed incident. Each item is a trace that attackers leave behind in almost every Microsoft 365 compromise, and that legitimate administration also leaves behind. The point is to ask: an executive or security officer takes each line to the person named and gets a yes-that-was-me, or escalates it the same day. Items marked <b>urgent</b> are the patterns most strongly associated with an active intrusion (federation changes, mailbox rules that hide mail, password sprays, legacy sign-ins, users Microsoft flags as at risk); items to <b>review</b> are normal administration that should match a ticket or a known reason. Sources: the Entra directory audit log, the Entra sign-in log (Entra ID P1), Entra ID Protection (P2), application registrations and Exchange Online.</p>
<h3>The situation</h3>
$(New-HtmlTable $iocSituationRows -Id 'tbl-ioc-situation' -Filters $IocSituationFilters -Columns 'Indicator','Status','What we found','Why it matters','Details' -EmptyText 'No indicators were evaluated.')
<p class='meta'>Urgent: investigate today with the person named. Review: confirm against a ticket or a known reason this week. Clear: looked for and not found. Not checked: the source was unavailable (licence, permission or switch). The tables below hold the detail behind each row.</p>
$(New-KvTable $iocKv -Issues $iocKvIssues)
$(if ($iocSignInNote) { "<p class='note'>$(HtmlEnc $iocSignInNote)</p>" })
$(if ($iocRiskNote) { "<p class='note'>$(HtmlEnc $iocRiskNote)</p>" })
$(if (-not $iocAuditAvailable) { "<p class='note'>Directory audit events could not be read: $(HtmlEnc (Get-GraphFailure 'role assignment audit events'))</p>" })
<h3 id='ioc-forwarding'>Mail forwarded outside the organization</h3>
$(if (-not $exoCollected) { "<p class='empty'>Exchange Online was not collected.</p>" } else { (New-HtmlTable $fwdMailboxExternal -Columns 'Mailbox','Forwarding SMTP address','Forwarding address (internal)','Keep copy in mailbox' -EmptyText 'No mailbox forwards to an external address.') + $(if ($SkipInboxRules) { "<p class='empty'>Inbox rules were not scanned (-SkipInboxRules).</p>" } else { New-HtmlTable $fwdRuleExternal -Columns 'Mailbox','Rule','Enabled','Forwards to' -EmptyText 'No inbox rule forwards to an external address.' }) + (New-HtmlTable $fwdTransportAll -Columns 'Rule','State','Redirect to','Bcc to','Copy / add to','Mode' -EmptyText 'No transport rule redirects, copies or Bcc''s mail.') })
<h3 id='ioc-apps'>Known-abused and unverified applications</h3>
$(New-HtmlTable $rogueAppRows -Columns 'Application','Known as','Typical abuse','Added','Enabled','Permissions granted','Who consented' -EmptyText 'No application on the known-abused list is present in this tenant.')
<h4>Unverified-publisher applications with mail, file, contact or directory permissions</h4>
$(New-HtmlTable $unverifiedAppRows -Columns 'Application','Publisher','Added to tenant','Delegated permissions','Application permissions','Who consented' -EmptyText 'Every application with those permissions comes from a verified publisher or Microsoft.')
<h3 id='ioc-sharing'>File sharing: Anyone links and outside people (unified audit log)</h3>
$(if (-not $ualCollected) { "<p class='empty'>$(HtmlEnc $ualNote)</p>" } else { New-HtmlTable $sharingRows -Id 'tbl-ioc-sharing' -Filters @(@{ Label = 'Anyone links'; Tag = 'anonymous' }, @{ Label = 'Outside people'; Tag = 'external' }, @{ Label = 'Internal'; Tag = 'internal' }) -Columns 'Date (local)','User','Action','Shared with','Item','Site','IP address' -EmptyText "No Anyone links or external sharing events in the last $AuditDays days." })
<h3 id='ioc-downloads'>File downloads and sync by user (unified audit log, last $downloadDays days)</h3>
$(if (-not $ualCollected) { "<p class='empty'>$(HtmlEnc $ualNote)</p>" } else { New-HtmlTable $downloadRows -Id 'tbl-ioc-downloads' -Filters @(@{ Label = 'Urgent'; Tag = 'urgent' }, @{ Label = 'Review'; Tag = 'review' }) -Columns 'User',"Files downloaded ($downloadDays days)",'Of which full sync downloads','Distinct IP addresses','Sites','Clients','Last' -EmptyText "No file download events in the last $downloadDays days." })
<h3 id='ioc-auditrules'>Mailbox rule, forwarding and permission changes (unified audit log)</h3>
$(if (-not $ualCollected) { "<p class='empty'>$(HtmlEnc $ualNote)</p>" } else { New-HtmlTable $auditRuleRows -Id 'tbl-ioc-auditrules' -Filters @(@{ Label = 'Forward / delete / grant'; Tag = 'danger' }, @{ Label = 'Other'; Tag = 'other' }) -Columns 'Date (local)','Performed by','Operation','Mailbox / object','Parameters','IP address','Client' -EmptyText "No inbox rule, forwarding or mailbox permission changes were recorded in the last $AuditDays days." })
<h3 id='ioc-urls'>Malicious link clicks (Safe Links, last 7 days)</h3>
$(if (-not $script:HasDefenderP1) { "<p class='empty'>Not available: Safe Links requires Defender for Office 365 (Plan 1 or 2).</p>" } elseif (-not $urlClicksCollected) { "<p class='empty'>Safe Links report could not be read: $(HtmlEnc (Get-GraphFailure 'Safe Links detail report'))</p>" } else { New-HtmlTable $urlClickRows -Id 'tbl-ioc-urls' -Filters @(@{ Label = 'Clicked through'; Tag = 'clicked' }, @{ Label = 'Blocked'; Tag = 'blocked' }) -Columns 'Date (local)','User','Result','URL','Where','Source' -EmptyText 'No malicious URL clicks in the last 7 days.' })
<h3 id='ioc-domains'>Domain and federation changes</h3>
$(New-HtmlTable $iocDomainRows -Columns 'Date (local)','Activity','Performed by','Target','Detail','Result' -EmptyText "No domain or federation settings changed in the last $AuditDays days.")
<h3 id='ioc-rules'>Inbox rules that hide, delete or forward mail</h3>
$(if (-not $exoCollected) { "<p class='empty'>Exchange Online was not collected.</p>" } elseif ($SkipInboxRules) { "<p class='empty'>Inbox rules were not scanned (-SkipInboxRules). Run without the switch to include this check.</p>" } else { New-HtmlTable $suspiciousRuleRows -Id 'tbl-ioc-rules' -Filters @(@{ Label = 'Enabled'; Tag = 'enabled' }, @{ Label = 'Disabled'; Tag = 'disabled' }) -Columns 'Mailbox','Rule','Enabled','Why it is suspicious','Conditions' -EmptyText 'No inbox rule deletes, hides or externally forwards mail.' })
<h3 id='ioc-risk'>Users flagged at risk (Entra ID Protection)</h3>
$(if ($iocRiskNote) { "<p class='empty'>$(HtmlEnc $iocRiskNote)</p>" } else { (New-HtmlTable $riskyUserRows -Columns 'User','UPN','Risk level','Risk state','Detail','Last updated' -EmptyText 'No users are currently flagged at risk.') + '<h4>Risk detections</h4>' + (New-HtmlTable $riskDetectionRows -Columns 'Detected (local)','User','Detection','Level','State','IP address','Location' -EmptyText "No risk detections in the last $AuditDays days.") })
<h3 id='ioc-signins'>Sign-in analysis (Entra ID P1)</h3>
$(if (-not $iocSignInsCollected) { "<p class='empty'>$(HtmlEnc $(if ($iocSignInNote) { $iocSignInNote } else { 'Sign-in log not collected.' }))</p>" } else { @"
<h4>Successful sign-ins by country</h4>
$(New-HtmlTable $countryRows -Columns 'Country','Successful sign-ins','Distinct users','Share' -EmptyText 'No successful sign-ins in the window.')
<h4>Users signing in from more than one country</h4>
$(New-HtmlTable $multiCountryRows -Id 'tbl-ioc-country' -Filters @(@{ Label = 'Admins'; Tag = 'admin' }) -Columns 'User','UPN','Countries','Cities','Last sign-in outside home country','Admin' -EmptyText 'Every user signed in from a single country.')
<h4>Successful legacy authentication sign-ins</h4>
$(New-HtmlTable $legacySignInRows -Columns 'User','Protocol / client','Successful sign-ins','Applications','Countries','Last' -EmptyText 'No successful legacy authentication sign-ins.')
<h4>MFA denied or failed</h4>
$(New-HtmlTable $mfaDeniedRows -Columns 'User','Attempts','Distinct IP addresses','Countries','Applications','Last','Admin' -EmptyText 'No MFA denials or failures.')
<h4>Blocked by Conditional Access (password was accepted)</h4>
$(New-HtmlTable $caBlockedRows -Columns 'User','Attempts','Distinct IP addresses','Countries','Applications','Last','Admin' -EmptyText 'No sign-ins were blocked by Conditional Access.')
<h4>Possible password spray sources</h4>
$(New-HtmlTable $sprayRows -Columns 'Source IP address','Country','Failed passwords','Distinct accounts targeted','Accounts (sample)','Last attempt' -EmptyText 'No source address failed passwords against three or more accounts.')
<h4>Accounts locked out</h4>
$(New-HtmlTable $lockoutRows -Columns 'User','Attempts','Distinct IP addresses','Countries','Last','Admin' -EmptyText 'No account lockouts.')
<h4>Risky sign-ins</h4>
$(New-HtmlTable $riskySignInRows -Columns 'Date (local)','User','Risk during sign-in','Aggregated risk','Risk state','Application','IP address','Location' -EmptyText 'No risky successful sign-ins (risk levels need Entra ID P2 to be shown).')
"@ })
<h3 id='ioc-admin'>Privileged role assignments added</h3>
$(New-HtmlTable $iocRoleRows -Columns 'Date (local)','Activity','Performed by','Target','Detail','Result' -EmptyText "No privileged role assignments were added in the last $AuditDays days.")
<h3>Security info (MFA methods) registered or changed</h3>
$(New-HtmlTable $iocMfaRows -Id 'tbl-ioc-mfa' -Filters $IocAuditFilters -Columns 'Date (local)','Activity','Performed by','Target','Detail','Result' -EmptyText "No MFA methods were registered or changed in the last $AuditDays days.")
<h3>Password resets and changes</h3>
$(New-HtmlTable $iocPwRows -Id 'tbl-ioc-pw' -Filters $IocAuditFilters -Columns 'Date (local)','Activity','Performed by','Target','Result' -EmptyText "No password resets or changes in the last $AuditDays days.")
<h3>Conditional Access and security policy changes</h3>
$(New-HtmlTable $iocCaRows -Columns 'Date (local)','Activity','Performed by','Target','Detail','Result' -EmptyText "No Conditional Access, named location or security policy changes in the last $AuditDays days.")
<h3>Application credentials, owners and new app registrations</h3>
$(New-HtmlTable $iocAppRows -Columns 'Date (local)','Activity','Performed by','Target','Detail','Result' -EmptyText "No application credential or ownership changes in the audit log for the last $AuditDays days.")
<h4>Secrets and certificates added to app registrations</h4>
$(New-HtmlTable $iocNewCredRows -Columns 'Application','Credential','Name','Added by','Valid from','Expires' -EmptyText "No app registration received a new secret or certificate in the last $AuditDays days.")
<h4>Enterprise applications added</h4>
$(New-HtmlTable $iocNewAppRows -Columns 'Application','Publisher','Added','Added / consented by','Delegated permissions','Application permissions' -EmptyText "No enterprise applications were added in the last $AuditDays days.")
<h3 id='ioc-accounts'>Accounts created, enabled, deleted or invited</h3>
$(New-HtmlTable $iocNewUserRows -Id 'tbl-ioc-users' -Filters @(@{ Label = 'Guests'; Tag = 'guest' }, @{ Label = 'Members'; Tag = 'member' }) -Columns 'User','UPN','Type','Enabled','Created','Licensed','Admin' -EmptyText "No accounts were created in the last $AuditDays days.")
$(if ($iocUserRows.Count) { '<h4>Account lifecycle audit events</h4>' + (New-HtmlTable $iocUserRows -Columns 'Date (local)','Activity','Performed by','Target','Result') })
<h3>Mailbox delegation and mail flow (Exchange Online)</h3>
$(if (-not $exoCollected) { "<p class='empty'>Exchange Online was not collected.</p>" } else { @"
<h4>Full Access delegates</h4>
$(if ($SkipInboxRules) { "<p class='empty'>Mailbox permissions were not scanned (-SkipInboxRules).</p>" } else { New-HtmlTable $fullAccessRows -Id 'tbl-ioc-delegates' -Filters @(@{ Label = 'User mailboxes'; Tag = 'user' }, @{ Label = 'Shared mailboxes'; Tag = 'shared' }) -Columns 'Mailbox','Mailbox type','Delegate (Full Access)','Rights' -EmptyText 'No non-owner Full Access delegates.' })
<h4>Send As grants</h4>
$(New-HtmlTable $sendAsRows -Columns 'Mailbox','Trustee (can send as)' -EmptyText 'No Send As grants to other accounts.')
<h4>Mail flow rules changed in the last $AuditDays days</h4>
$(New-HtmlTable $recentTransportRows -Columns 'Rule','State','Changed','Created','Actions' -EmptyText "No transport rules were created or changed in the last $AuditDays days.")
"@ })
</section>
"@

$teamsSectionHtml = @"
<section id='teams'><h2>17. Microsoft Teams</h2>
<p class='meta'>Every team in the tenant with its visibility, owners, members and guests, the activity reported by Microsoft for the last 90 days, and the tenant-wide settings that decide who outside the institution can reach staff through Teams. Teams are Microsoft 365 groups: each one also owns a SharePoint site (section 18) that inherits the team's membership and guests.</p>
$(New-KvTable $teamsKv -Issues $teamsKvIssues)
<h3>Tenant-wide Teams policies</h3>
$(New-KvTable $teamsPolicyKv -Issues $teamsPolicyIssues)
<h3>Teams</h3>
$(if ($reportConcealedNote -and $teamsConcealed) { "<p class='note'>$(HtmlEnc $reportConcealedNote)</p>" })
$(if (@($teamGroups).Count -gt $teamLimit) { "<p class='note'>Only the first $teamLimit teams (alphabetically) were expanded for owners, members and guests.</p>" })
$(New-HtmlTable $teamRows -Id 'tbl-teams' -Filters $TeamFilters -Columns 'Team','Visibility','Archived','Owners','Members','Guests','Guest names','Last activity','Active users (90 days)','Members may add apps / connectors','Guests may create channels','Created','Classification' -EmptyText $(if (Get-GraphFailure 'teams') { "Teams could not be read: $(Get-GraphFailure 'teams')" } else { 'No teams exist in this tenant.' }))
</section>
"@

$spoSectionHtml = @"
<section id='sharepoint'><h2>18. SharePoint and OneDrive</h2>
<p class='meta'>Tenant-wide sharing and access settings for SharePoint and OneDrive, every site with its owner, activity and storage, every OneDrive with the state of its owner's account, and the users who shared files outside the organization in the last 30 days. Site and OneDrive inventories come from the Microsoft 365 usage reports (90-day period), which may lag by up to 48 hours.</p>
$(New-KvTable $spoKv -Issues $spoKvIssues)
$(if ($reportConcealedNote) { "<p class='note'>$(HtmlEnc $reportConcealedNote)</p>" })
<h3>Users who shared files externally (last 30 days)</h3>
$(New-HtmlTable $extShareRows -Columns 'User','UPN','Files shared externally from SharePoint (30 days)','Files shared externally from OneDrive (30 days)','Files shared internally (30 days)' -EmptyText $(if ($spoActivity.Count -or $odActivity.Count) { 'No user shared a file externally in the last 30 days.' } else { "Activity reports not available$(if (Get-GraphFailure 'SharePoint user activity report') { ": $(Get-GraphFailure 'SharePoint user activity report')" })." }))
<h3>SharePoint sites</h3>
$(New-HtmlTable $siteRows -Id 'tbl-sites' -Filters $SiteFilters -Columns 'Site','Type','Owner','Last activity','Files','Active files (90 days)','Storage used (GB)','Page views (90 days)','Deleted' -EmptyText $(if ($siteReport.Count) { 'No sites returned.' } else { "Site usage report not available$(if (Get-GraphFailure 'SharePoint site usage report') { ": $(Get-GraphFailure 'SharePoint site usage report')" })." }))
<h3>OneDrive accounts</h3>
$(New-HtmlTable $odRows -Id 'tbl-onedrive' -Filters $OdFilters -Columns 'Owner','Account','Last activity','Files','Active files (90 days)','Storage used (GB)','Deleted' -EmptyText $(if ($odReport.Count) { 'No OneDrive accounts returned.' } else { "OneDrive usage report not available$(if (Get-GraphFailure 'OneDrive usage report') { ": $(Get-GraphFailure 'OneDrive usage report')" })." }))
</section>
"@


# ---- Executive page: verdict, headlines, scorecard, rocks (90-day objectives), issues list, plan, board-report mapping
$critCount = @($remediationSorted | Where-Object { $_.Priority -eq 'Critical' }).Count
$highCount = @($remediationSorted | Where-Object { $_.Priority -eq 'High' }).Count
$remIndexByRef = @{}; $ri = 0; foreach ($x in $remediationSorted) { $remIndexByRef[[string]$x.Ref] = $ri; $ri++ }
function Get-FindingLinkHtml { param($Finding, [switch]$Ticket) "<a href='#rem-$($Finding.Ref)'>$(HtmlEnc $Finding.Finding)</a>$(if ($Ticket) { " <button type='button' class='tk' data-i='$($remIndexByRef[[string]$Finding.Ref])' title='Preview and copy the ticket for this item'>View ticket</button>" })" }
$controlsPct = if ($controlsRated.Count) { [math]::Round(100 * $controlsMet.Count / $controlsRated.Count) } else { 0 }
$maturity = if ($controlsPct -lt 40) { 'Baseline' } elseif ($controlsPct -lt 60) { 'Evolving' } elseif ($controlsPct -lt 80) { 'Intermediate' } elseif ($controlsPct -lt 95) { 'Advanced' } else { 'Leading' }
$verdict = if ($iocUrgentCount -gt 0 -or $critCount -gt 0 -or $ualEnabled -eq $false) {
    @{ Label = 'Act now'; Class = 'bad'; Text = (Join-List (@($(if ($iocUrgentCount) { "$iocUrgentCount warning sign(s) of an intrusion need a same-day check" }), $(if ($critCount) { "$critCount Critical item(s) are due this week" }), $(if ($ualEnabled -eq $false) { 'activity logging is turned off, so sharing, downloads and mailbox changes are not being recorded' })) | Where-Object { $_ }) '; ') + '.' }
} elseif ($highCount -gt 0 -or $controlsPct -lt 70 -or ($null -ne $scoreCurrent -and $scoreCurrent -lt $scoreTarget)) {
    @{ Label = 'Needs attention'; Class = 'mid'; Text = "No warning signs of an intrusion and no Critical items. $(if ($highCount) { "$highCount High item(s) are due within 30 days. " })$controlsPct% of measured controls are in place$(if ($null -ne $scoreCurrent -and $scoreCurrent -lt $scoreTarget) { " and Secure Score is below the $TargetScorePercent% goal" })." }
} else {
    @{ Label = 'In good shape'; Class = 'good'; Text = "No warning signs of an intrusion, no Critical or High items, $controlsPct% of measured controls in place$(if ($null -ne $scoreCurrent) { ' and Secure Score at goal' }). Keep the monthly cadence and work the remaining Medium and Low items." }
}
$prevUsersSnap = Get-JsonProp $previous 'Users'; $prevDevSnap = Get-JsonProp $previous 'Devices'; $prevIocSnap = Get-JsonProp $previous 'Ioc'; $prevFindSnap = Get-JsonProp $previous 'Findings'; $prevShareSnap = Get-JsonProp $previous 'Sharing'
$scoreDelta = if ($null -ne $scoreCurrent -and $previous -and $null -ne (Get-JsonProp $previous 'SecureScore')) { [math]::Round([double]$scoreCurrent - [double](Get-JsonProp $previous 'SecureScore'), 1) } else { $null }
$trend = if (-not $previous) { @{ Label = 'Baseline'; Class = 'neutral'; Text = 'First report kept for this tenant: the baseline that next month is measured against.' } }
         else {
             $pos = $improvedCount + $resolvedCount; $neg = $regressedCount + $newCount
             $tl = if ($pos -gt $neg) { 'Improving' } elseif ($neg -gt $pos) { 'Slipping' } else { 'Holding' }
             @{ Label = $tl; Class = $(switch ($tl) { 'Improving' { 'good' } 'Slipping' { 'bad' } default { 'mid' } }); Text = "Since $(Get-JsonProp $previous 'Date'): $resolvedCount item(s) resolved and $improvedCount control(s) improved; $newCount new item(s) and $regressedCount control(s) regressed$(if ($null -ne $scoreDelta) { $(if ($scoreDelta -gt 0) { "; Secure Score up $scoreDelta points" } elseif ($scoreDelta -lt 0) { "; Secure Score down $([math]::Abs($scoreDelta)) points" } else { '; Secure Score unchanged' }) })." }
         }
# Scorecard: measurable, goal, actual, on/off track, movement since the last report
function Get-Movement { param($Now, $Before, [switch]$HigherIsBetter)
    if ($null -eq $Now -or $null -eq $Before -or -not $previous) { return $(if ($previous) { '' } else { 'baseline' }) }
    $n = [double]$Now; $b = [double]$Before
    if ($n -eq $b) { return 'no change' }
    $better = if ($HigherIsBetter) { $n -gt $b } else { $n -lt $b }
    return $(if ($better) { 'better (was ' } else { 'worse (was ' }) + "$b)"
}
$scorecardRows = @()
function Add-Score { param([string]$Measurable, [string]$Goal, $Actual, [bool]$OnTrack, [string]$Movement, [string]$Issue, [bool]$Known = $true)
    $script:scorecardRows += [ordered]@{ 'Measurable' = $Measurable; 'Goal' = $Goal; 'Actual' = $(if ($Known) { [string]$Actual } else { 'Not measured' }); 'Status' = $(if (-not $Known) { 'Not measured' } elseif ($OnTrack) { 'On track' } else { 'Off track' }); 'Since last report' = $Movement; '_issues' = $(if ($Known -and -not $OnTrack -and $Issue) { @{ 'Status' = $Issue } } else { $null }); '_tags' = $(if (-not $Known) { 'notmeasured' } elseif ($OnTrack) { 'ontrack' } else { 'offtrack' }) }
}
$script:scorecardRows = @()
$prevControlsMet = $(if ($previous) { @((Get-JsonProp $previous 'Controls').PSObject.Properties | Where-Object { [string]$_.Value -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded|None found)' }).Count } else { $null })
Add-Score 'Security controls in place' '100%' "$controlsPct% ($($controlsMet.Count) of $($controlsRated.Count))" ($controlsPct -ge 90) (Get-Movement $controlsMet.Count $prevControlsMet -HigherIsBetter) 'secure-score'
Add-Score 'Microsoft Secure Score' "$TargetScorePercent%" "$scorePct%" ($null -ne $scoreCurrent -and $scoreCurrent -ge $scoreTarget) (Get-Movement $scorePct (Get-JsonProp $previous 'SecureScorePct') -HigherIsBetter) 'secure-score' ($null -ne $scoreCurrent)
Add-Score 'Warning signs of an intrusion needing a same-day check' '0' $iocUrgentCount ($iocUrgentCount -eq 0) (Get-Movement $iocUrgentCount (Get-JsonProp $prevIocSnap 'Urgent')) 'ioc-urgent' $iocCollected
Add-Score 'High-priority items due within 30 days (Critical and High)' '0' $findCritHigh ($findCritHigh -eq 0) (Get-Movement $findCritHigh $(if ($prevFindSnap) { [int](Get-JsonProp $prevFindSnap 'Critical') + [int](Get-JsonProp $prevFindSnap 'High') } else { $null })) 'ioc-review'
Add-Score 'Activity recording (audit log) turned on' 'Yes' $(if ($ualEnabled -eq $true) { 'Yes' } elseif ($ualEnabled -eq $false) { 'No' } else { 'Unknown' }) ($ualEnabled -eq $true) '' 'ual-off' ($null -ne $ualEnabled)
Add-Score 'Staff accounts without a second sign-in factor (MFA)' '0' $noMfaCount ($noMfaCount -eq 0) (Get-Movement $noMfaCount (Get-JsonProp $prevUsersSnap 'NoMfa')) 'mfa-not-registered' ($null -ne $noMfaCount)
Add-Score 'Staff using phishing-proof sign-in (passkeys)' '90%' "$passkeyPct%" ($null -ne $passkeyPct -and $passkeyPct -ge 90) (Get-Movement $passkeyPct (Get-JsonProp $prevUsersSnap 'PasskeyPct') -HigherIsBetter) 'passkey-not-started' ($null -ne $passkeyPct)
Add-Score 'Accounts with full administrative control (Global Administrators)' '2 to 4' $gaCount ($gaCount -ge 2 -and $gaCount -le 4) (Get-Movement $gaCount (Get-JsonProp $prevUsersSnap 'GlobalAdmins')) 'ga-count'
Add-Score "Unused staff accounts still enabled (no sign-in in $StaleDays days)" '0' $staleUserCount ($staleUserCount -eq 0) (Get-Movement $staleUserCount (Get-JsonProp $prevUsersSnap 'Stale')) 'user-stale' ($null -ne $staleUserCount)
Add-Score 'Company computers outside central management (Intune)' '0' $notInIntuneCount ($notInIntuneCount -eq 0) (Get-Movement $notInIntuneCount (Get-JsonProp $prevDevSnap 'NotInIntune')) 'intune-coverage' (-not (Get-GraphFailure 'Intune managed devices'))
Add-Score 'Files or calendars open to anyone on the internet' '0' ($shareAnyone + $anonLinkCount) (($shareAnyone + $anonLinkCount) -eq 0) (Get-Movement ($shareAnyone + $anonLinkCount) (Get-JsonProp $prevShareSnap 'Anyone')) 'share-anonymous' (-not ($SkipShareScan -and -not $ualCollected))
Add-Score 'Email anti-spoofing protections (SPF, DKIM, DMARC)' 'All domains complete' (Join-List (@("SPF $spfStatus", "DKIM $dkimStatus", "DMARC $dmarcStatus")) ', ') ($spfStatus -eq 'Complete' -and $dkimStatus -eq 'Complete' -and $dmarcStatus -eq 'Complete') '' 'dmarc-missing' ($dnsAvailable -and $emailAuthRows.Count -gt 0)
$scorecardRows = @($script:scorecardRows)
$onTrack = @($scorecardRows | Where-Object { $_['Status'] -eq 'On track' }).Count; $measured = @($scorecardRows | Where-Object { $_['Status'] -ne 'Not measured' }).Count
# Headlines: the news a board member would repeat, bad first
$headlines = @()
if ($iocUrgentCount) { $headlines += @{ Kind = 'bad'; Text = "$iocUrgentCount warning sign(s) of an intrusion need a same-day check (section 2)." } }
if ($ualEnabled -eq $false) { $headlines += @{ Kind = 'bad'; Text = 'Activity recording is turned off: nobody can see who shared, downloaded or changed what, and an intrusion could not be investigated.' } }
if ($critCount) { $headlines += @{ Kind = 'bad'; Text = "$critCount Critical item(s) are due this week: $(Join-List (@($remediationSorted | Where-Object { $_.Priority -eq 'Critical' } | Select-Object -First 3 | ForEach-Object { $_.Finding })) ', ')$(if ($critCount -gt 3) { ' and more' })." } }
if ($gaNoMfa.Count) { $headlines += @{ Kind = 'bad'; Text = "$($gaNoMfa.Count) account(s) with full administrative control have no second sign-in factor (MFA)." } }
if (($shareAnyone + $anonLinkCount) -gt 0) { $headlines += @{ Kind = 'bad'; Text = "$($shareAnyone + $anonLinkCount) item(s) are open to anyone on the internet." } }
if ($previous -and $resolvedCount) { $headlines += @{ Kind = 'good'; Text = "$resolvedCount item(s) fixed since the last report." } }
if ($null -ne $scoreDelta -and $scoreDelta -gt 0) { $headlines += @{ Kind = 'good'; Text = "Secure Score up $scoreDelta points to $scorePct%." } }
if ($iocCollected -and $iocUrgentCount -eq 0) { $headlines += @{ Kind = 'good'; Text = "No urgent warning sign of an intrusion in the last $AuditDays days." } }
if ($mfaStatus -like 'Enforced*') { $headlines += @{ Kind = 'good'; Text = 'Every sign-in requires a second factor (MFA).' } }
if ($legacyStatus -like 'Blocked*') { $headlines += @{ Kind = 'good'; Text = 'Older, less secure sign-in methods are blocked.' } }
if ($dmarcStatus -eq 'Complete') { $headlines += @{ Kind = 'good'; Text = 'Every email domain is protected against impersonation (DMARC).' } }
if ($null -ne $passkeyPct -and $passkeyPct -ge 90) { $headlines += @{ Kind = 'good'; Text = "$passkeyPct% of staff use phishing-proof sign-in (passkeys)." } }
if ($previous -and $regressedCount) { $headlines += @{ Kind = 'bad'; Text = "$regressedCount control(s) regressed since the last report." } }
$headlines = @(@($headlines | Where-Object { $_.Kind -eq 'bad' } | Select-Object -First 4) + @($headlines | Where-Object { $_.Kind -eq 'good' } | Select-Object -First 3))
# Repeat findings: open in this and at least two earlier reports
$priorSnaps = @($history | Where-Object { [string](Get-JsonProp $_ 'Date') -ne $snapshot.Date })
$repeatRows = @()
if ($priorSnaps.Count -ge 2) {
    foreach ($fnd in $remediationSorted) {
        $seen = @($priorSnaps | Where-Object { @(Get-JsonProp $_ 'FindingNames') -contains $fnd.Finding })
        if ($seen.Count -ge 2) { $repeatRows += [ordered]@{ 'Finding' = $fnd.Finding; 'Priority' = $fnd.Priority; 'Open in reports' = $seen.Count + 1; 'First seen' = ([string](Get-JsonProp $seen[0] 'Date')).Substring(0, 10); 'Owner' = $(if ($fnd.Playbook) { [string]$fnd.Playbook.Owner } else { 'IT provider / MSP' }); 'Ref' = $fnd.Ref; '_issues' = @{ 'Open in reports' = 'repeat-finding' }; '_html' = @{ 'Finding' = (Get-FindingLinkHtml $fnd -Ticket) } } }
    }
}
$repeatRows = @($repeatRows | Sort-Object { -[int]$_['Open in reports'] } | Select-Object -First 15)
if ($repeatRows.Count) { $headlines = @(@(@{ Kind = 'bad'; Text = "$($repeatRows.Count) finding(s) have been open for three or more reports." }) + $headlines) }
# Rocks: 90-day objectives that group the findings into business outcomes
$RockThemes = @(
    @{ Rock = 'Licensing in place';             Objective = 'The licences needed for the controls below are approved and assigned';                                  Match = '(?i)not licensed' }
    @{ Rock = 'Every sign-in protected';        Objective = 'MFA enforced for all users, legacy authentication blocked, administrators on phishing-resistant MFA, no stale accounts'; Match = '(?i)MFA|multifactor|multi-factor|legacy|authenticat|passkey|passwordless|security defaults|Temporary Access Pass|password|stale|sign-in|self-service|SSPR' }
    @{ Rock = 'Administration and apps locked down'; Objective = 'Two to four Global Administrators, user consent restricted, no unknown or risky applications';      Match = '(?i)global admin|administrator|consent|application|app registration|apps|role|privileged|PIM' }
    @{ Rock = 'Email protected';                Objective = 'SPF, DKIM and DMARC enforced, forwarding blocked, Defender for Office 365 presets on';                  Match = '(?i)SPF|DKIM|DMARC|forward|phish|safe links|safe attachments|spam|malware|impersonat|mail|Defender for Office|quarantine|spoof' }
    @{ Rock = 'Every device managed';           Objective = 'All computers and phones enrolled in Intune, compliant, encrypted and protected by Defender for Endpoint'; Match = '(?i)Intune|device|compliance|endpoint|BitLocker|encrypt|Defender for Endpoint|join' }
    @{ Rock = 'Data stays where it belongs';    Objective = 'No Anyone links, guests and external sharing limited to approved partners, Teams and SharePoint access governed'; Match = '(?i)shar|guest|Teams|SharePoint|OneDrive|external|anonymous|anyone|open to|link' }
    @{ Rock = 'See and respond';                Objective = 'Audit logging on, indications of compromise reviewed every month, risky sign-ins and alerts handled';    Match = '(?i)audit|logging|compromise|risk|conceal|report|alert|Secure Score' }
)
function Get-RockFor { param([string]$Finding) foreach ($rt in $RockThemes) { if ($Finding -match $rt.Match) { return $rt.Rock } } ; return 'See and respond' }
$prevNames = @(if ($previous) { Get-JsonProp $previous 'FindingNames' } else { @() })
$rockRows = @()
foreach ($rt in $RockThemes) {
    $open = @($remediationSorted | Where-Object { (Get-RockFor $_.Finding) -eq $rt.Rock })
    $prevOpen = @($prevNames | Where-Object { $_ -and (Get-RockFor $_) -eq $rt.Rock }).Count
    if ($open.Count -eq 0 -and $prevOpen -eq 0) { continue }
    $owners = @($open | ForEach-Object { if ($_.Playbook) { ([string]$_.Playbook.Owner -split ';')[0] } else { 'IT provider / MSP' } } | Group-Object | Sort-Object Count -Descending | Select-Object -First 1 | ForEach-Object { $_.Name })
    $urgentIn = @($open | Where-Object { $_.Priority -in @('Critical', 'High') }).Count
    $status = if ($open.Count -eq 0) { 'Done' } elseif (-not $previous) { 'Not started' } elseif ($open.Count -lt $prevOpen) { 'On track' } elseif ($open.Count -gt $prevOpen) { 'Off track' } else { 'Not started' }
    $due = if ($open.Count -eq 0) { 'Complete' } elseif ($urgentIn) { $ReportDate.AddDays(30).ToString('MMMM d, yyyy') } elseif (@($open | Where-Object { $_.Priority -eq 'Medium' }).Count) { $ReportDate.AddDays(90).ToString('MMMM d, yyyy') } else { $ReportDate.AddDays(365).ToString('MMMM d, yyyy') }
    $openHtml = if ($open.Count) { "<b>$($open.Count)</b>$(if ($urgentIn) { " <span class='muted'>($urgentIn Critical/High)</span>" })<ul class='perms'>" + ((@($open | Select-Object -First 8) | ForEach-Object { "<li><span class='muted'>$(HtmlEnc $_.Priority):</span> $(Get-FindingLinkHtml $_)</li>" }) -join '') + "</ul>$(if ($open.Count -gt 8) { "<a href='#remediation' class='muted'>and $($open.Count - 8) more in the plan</a>" })" } else { "<span class='muted'>0 (complete)</span>" }
    $rockRows += [ordered]@{ 'Rock (90-day objective)' = $rt.Rock; 'Outcome' = $rt.Objective; 'Owner' = $(if ($owners) { $owners[0] } else { '' }); 'Open items' = "$($open.Count)$(if ($urgentIn) { " ($urgentIn Critical/High)" })"; 'Due' = $due; 'Status' = $status; '_tags' = $status.ToLower().Replace(' ', ''); '_html' = @{ 'Open items' = $openHtml } }
}
$rockRows = @($rockRows)
# Issues list: what the committee must decide or unblock (EOS: identify, discuss, solve)
$PolicyDecisionControls = @('Teams guest and external access', 'SharePoint and OneDrive external sharing', 'User consent to applications', 'Users restricted from joining devices to Entra ID', 'Automatic email forwarding outside the tenant', 'Phishing-resistant MFA (passkeys) rollout')
$issueRows = @()
foreach ($fnd in $remediationSorted) {
    $isLic = ($fnd.Finding -match '(?i)not licensed') -or ($fnd.Playbook -and ([string]$fnd.Playbook.Owner) -match '(?i)committee|licen|budget')
    $isPol = $PolicyDecisionControls -contains $fnd.Finding
    if ($isLic -or $isPol) { $issueRows += [ordered]@{ 'Issue' = $fnd.Finding; 'Type' = $(if ($isLic) { 'Licence / budget decision' } else { 'Policy decision' }); 'What the committee is asked to decide' = ([string]$fnd.Action -split '(?<=\.)\s+')[0]; 'Who' = "IT committee decides; $(if ($fnd.Playbook) { ([string]$fnd.Playbook.Owner -split ';')[0] } else { 'IT provider / MSP' }) implements"; 'Phase' = $fnd.Phase; '_html' = @{ 'Issue' = (Get-FindingLinkHtml $fnd -Ticket) } } }
}
foreach ($rr in $repeatRows) { $issueRows += [ordered]@{ 'Issue' = $rr['Finding']; 'Type' = 'Repeat finding'; 'What the committee is asked to decide' = "Open for $($rr['Open in reports']) reports since $($rr['First seen']): name the owner and date, and remove what is blocking it."; 'Who' = $rr['Owner']; 'Phase' = ''; '_html' = $rr['_html'] } }
$ualFinding = $remediationSorted | Where-Object { $_.Finding -eq 'Audit logging (unified audit log)' } | Select-Object -First 1
$concealFinding = $remediationSorted | Where-Object { $_.Finding -eq 'Usage report names are concealed' } | Select-Object -First 1
if ($ualEnabled -eq $false) { $issueRows += [ordered]@{ 'Issue' = 'Activity logging is off'; 'Type' = 'Coverage gap'; 'What the committee is asked to decide' = 'Direct the IT provider to turn on the unified audit log this week; until then intrusion and data-leak checks cannot be completed.'; 'Who' = 'IT committee directs; IT provider / MSP implements'; 'Phase' = 'Now: this week'; '_html' = @{ 'Issue' = $(if ($ualFinding) { "<a href='#rem-$($ualFinding.Ref)'>Activity logging is off</a> <button type='button' class='tk' data-i='$($remIndexByRef[[string]$ualFinding.Ref])'>View ticket</button>" } else { "<a href='#ioc'>Activity logging is off</a>" }) } } }
if ($reportConcealedNote) { $issueRows += [ordered]@{ 'Issue' = 'Usage report names are concealed'; 'Type' = 'Coverage gap'; 'What the committee is asked to decide' = 'Approve clearing the report privacy setting so OneDrive, site and team owners are named.'; 'Who' = 'IT committee approves; IT provider / MSP implements'; 'Phase' = 'Now: this week'; '_html' = @{ 'Issue' = $(if ($concealFinding) { "<a href='#rem-$($concealFinding.Ref)'>Usage report names are concealed</a> <button type='button' class='tk' data-i='$($remIndexByRef[[string]$concealFinding.Ref])'>View ticket</button>" } else { "<a href='#sharepoint'>Usage report names are concealed</a>" }) } } }
if ($shareSitesDenied -gt 0) { $issueRows += [ordered]@{ 'Issue' = "$shareSitesDenied site(s) could not be read by the account that ran the report"; 'Type' = 'Coverage gap'; 'What the committee is asked to decide' = 'Run the report as a SharePoint administrator with site access, or accept that sharing on those sites is unverified.'; 'Who' = 'Security officer'; 'Phase' = 'Next 30 days'; '_links' = @{ 'Issue' = '#sharing' } } }
if ($script:HasMDE -and -not $mdeCollected -and $mdeCheck -in @('consent', 'role')) { $issueRows += [ordered]@{ 'Issue' = 'Device protection (Defender for Endpoint) could not be verified device by device'; 'Type' = 'Coverage gap'; 'What the committee is asked to decide' = $(if ($mdeCheck -eq 'consent') { 'Ask a Global Administrator to approve the read-only Defender permissions (link in section 16), so the next report shows whether every device is protected.' } else { 'Give the reporting account read access to Defender (Security Reader), so the next report shows whether every device is protected.' }); 'Who' = 'Security officer; IT provider / MSP implements'; 'Phase' = 'Next 30 days'; '_links' = @{ 'Issue' = '#mde' } } }
if ($script:HasEntraP1 -eq $false) { $issueRows += [ordered]@{ 'Issue' = 'Entra ID P1 not licensed'; 'Type' = 'Coverage gap'; 'What the committee is asked to decide' = 'Without P1 there is no Conditional Access, sign-in log or risk policy; decide whether to license it (included in Business Premium).'; 'Who' = 'IT committee'; 'Phase' = 'Next 90 days'; '_links' = @{ 'Issue' = '#remediation' } } }
$issueRows = @($issueRows)
# Plan by phase and the GLBA board-report mapping
$PhaseDays = @{ 'Now: this week' = 7; 'Next 30 days' = 30; 'Next 90 days' = 90; 'Next 12 months and ongoing' = 365 }
$planRows = @($PhaseOrder | ForEach-Object { $ph = $_; $items = @($remediationSorted | Where-Object { $_.Phase -eq $ph }); [ordered]@{ 'Phase' = $ph; 'Open items' = $items.Count; 'Target date' = $ReportDate.AddDays($PhaseDays[$ph]).ToString('MMMM d, yyyy'); 'Owners' = (Join-List (@($items | ForEach-Object { if ($_.Playbook) { ([string]$_.Playbook.Owner -split ';')[0] } else { 'IT provider / MSP' } }) | Select-Object -Unique -First 3) '; '); 'Closed when' = 'the next monthly report lists the item as resolved' } })
$glbaRows = @(
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Overall status of the information security program and compliance'; 'Where it is in this report' = "This page: the verdict, scorecard, rocks and trend (Microsoft 365 scope)" },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Risk assessment'; 'Where it is in this report' = 'Section 2 (possible indications of compromise), section 3 (control gaps) and section 19 (who can reach what)' },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Risk management and control decisions'; 'Where it is in this report' = 'Issues list on this page; section 4 roadmap with owners, effort and target dates' },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Service provider arrangements'; 'Where it is in this report' = 'Section 11 (applications with access to data), section 19 (external access); tickets go to the IT provider / MSP with a "done when" test' },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Results of testing'; 'Where it is in this report' = 'Sections 3 and 5: every control measured monthly against the live tenant and mapped to NIST, CIS and FFIEC / NCUA' },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Security breaches or violations and management responses'; 'Where it is in this report' = 'Section 2, with the outcome of each urgent item recorded by the security officer' },
    [ordered]@{ 'Board report element (GLBA Interagency Guidelines)' = 'Recommendations for changes'; 'Where it is in this report' = 'Section 4 remediation plan; rocks and issues on this page' }
)

$legendHtml = "<p class='legend'><span class='st good'>&#10004;</span> Pass &nbsp; <span class='st mid'>&#9684;</span> Partial: review &nbsp; <span class='st warn'>&#10006;</span> Action needed &nbsp; <span class='st na'>&mdash;</span> Not measured (licence, permission or switch) &nbsp; <span class='i'>i</span> click for why it matters and what to do</p>"
# ---- Framework alignment table and governance figures for the executive summary
$frameworkRows = @(); $fwMet = 0; $fwRated = 0
foreach ($row in $script:summaryRows) {
    $fw = $FrameworkMap[[string]$row.Control]
    if (-not $fw) { continue }
    $st = [string]$row.Status
    $rated = $st -notmatch '^(Not available|Not collected)'
    $met = $st -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded|None found)'
    if ($rated) { $fwRated++ }; if ($met) { $fwMet++ }
    $frameworkRows += [ordered]@{ 'Control' = $row.Control; 'Result' = (Get-ResultWord $st); 'Status' = $st; 'NIST CSF 2.0 / SP 800-63B' = $fw.Nist; 'CIS Microsoft 365 Foundations Benchmark' = $fw.Cis365; 'CIS Controls v8' = $fw.CisV8; 'FFIEC / NCUA' = $fw.Ffiec; '_tags' = $(if ($met) { 'met' } elseif ($rated) { 'gap' } else { 'na' }) + ' ' + (Get-StatusKind $st); '_issues' = $(if (-not $met -and $rated -and $SummaryIssueMap.ContainsKey($row.Control)) { @{ 'Status' = $SummaryIssueMap[$row.Control] } } else { $null }) }
}
$FrameworkFilters = @(@{ Label = 'Action needed'; Tag = 'warn' }, @{ Label = 'Partial'; Tag = 'mid' }, @{ Label = 'Pass'; Tag = 'good' }, @{ Label = 'Not measured'; Tag = 'na' })
$frameworkSectionHtml = @"
<section id='frameworks'><h2>5. Framework Alignment (NIST, CIS, FFIEC / NCUA)</h2>
<p class='meta'>Each control this report measures, with the expectation it answers in the frameworks examiners and auditors use: <b>NIST Cybersecurity Framework 2.0</b> subcategories and <b>NIST SP 800-63B</b> authenticator assurance levels; the <b>CIS Microsoft 365 Foundations Benchmark</b> (section numbering follows v4.0; check the edition your auditor uses) and <b>CIS Critical Security Controls v8</b> safeguards; and <b>FFIEC</b> references: the IT Examination Handbook Information Security booklet, the 2021 <i>Authentication and Access to Financial Institution Services and Systems</i> guidance, and the Cybersecurity Assessment Tool / NCUA ACET domain codes (D3.PC.Am = Domain 3, Preventative Controls, Access and Data Management) that many institutions still use for the inherent risk and maturity assessment. The mapping is informational and does not replace the auditor's own judgement. <b>$fwMet of $fwRated</b> mapped controls are met.</p>
$legendHtml
$(New-HtmlTable $frameworkRows -Id 'tbl-frameworks' -Filters $FrameworkFilters -Columns 'Control','Result','Status','NIST CSF 2.0 / SP 800-63B','CIS Microsoft 365 Foundations Benchmark','CIS Controls v8','FFIEC / NCUA' -EmptyText 'No controls were measured.')
<p class='meta'>Pass means the control is in place as measured; Partial: review means it is partly in place or needs a judgement; Action needed means it is missing and appears in the Remediation Plan; Not measured means the licence, permission or switch did not allow the check.</p>
<h3>What this report does not test</h3>
<p class='meta'>So that nobody mistakes this report for a complete assessment: it does not test backup and recovery of Microsoft 365 data, the incident response plan, security awareness training, vendor and core-provider management, data classification and retention policies in Purview, physical security, network perimeter controls, or on-premises Active Directory (see the companion AuditHat AD Governance Report). Those remain part of the institution's wider information security program and the GLBA Safeguards Rule risk assessment.</p>
</section>
"@
$phaseCounts = Join-List ($PhaseOrder | ForEach-Object { $ph = $_; "$ph`: $(@($remediationSorted | Where-Object { $_.Phase -eq $ph }).Count)" }) '; '
$governanceKv = [ordered]@{
    'Report number for this tenant'  = $(if ($NoHistory) { 'History disabled (-NoHistory)' } else { "$($history.Count) (kept in $(Split-Path $historyPath -Leaf))" })
    'Previous report'                = $(if ($previous) { [string](Get-JsonProp $previous 'Date') } else { 'None: this is the baseline' })
    'Since the previous report'      = $(if ($previous) { "$improvedCount improved, $regressedCount regressed, $resolvedCount finding(s) resolved, $newCount new" } else { 'n/a' })
    'Open findings by phase'         = $phaseCounts
    'Possible indications of compromise' = $(if ($iocCollected) { "$iocUrgentCount urgent, $iocReviewCount to review (section 2)" } else { 'Not collected' })
    'Review cadence'                 = 'Run monthly. The security officer and IT provider review each report and raise tickets the same week; the IT committee reviews quarterly; the board receives this summary at least annually and after any incident.'
    'Next report due'                = $ReportDate.AddDays(30).ToString('MMMM d, yyyy')
    'Standards referenced'           = 'NIST CSF 2.0, NIST SP 800-63B, CIS Microsoft 365 Foundations Benchmark, CIS Controls v8, FFIEC IT Examination Handbook, FFIEC Authentication Guidance (2021), FFIEC CAT / NCUA ACET domains (section 5)'
}

$shareSectionHtml = @"
<section id='sharing'><h2>19. Sharing and Access</h2>
<p class='meta'>Who can reach what. Every share and access grant the signed-in account can see across SharePoint sites, OneDrive, Teams and Exchange, classified by how far it reaches: <b>Anyone</b> (open to the internet, no sign-in), <b>Whole organization</b> (every account in the tenant, including signed-in guests), <b>External</b> (named people outside the institution: guests, external links, shared-channel users), <b>Specific people</b> inside the organization, with a flag when the recipient is <b>outside the owning team</b> or in a <b>different department</b> from the owner. Use it to find data leaving the bank, data open wider than intended, and insiders with access they should not have. The scan reads only what $(HtmlEnc $ctx.Account) is allowed to read; sites and OneDrives it cannot open are counted, not listed, so a SharePoint Administrator or site owner may see more.</p>
$(New-KvTable $shareKv -Issues $shareKvIssues)
<h3>All shares and access grants</h3>
$(New-HtmlTable @($shareRows | Sort-Object @{ Expression = { @{ 'Urgent' = 0; 'Review' = 1; 'Info' = 2 }[[string]$_['Level']] } }, @{ Expression = { [string]$_['Source'] } } | Select-Object -First 2000) -Id 'tbl-shares' -Filters $ShareFilters -Columns 'Source','Site / team / mailbox','Item','Access type','Who has access','Rights','Scope','Outside owning team','Department (owner / who)','Granted by','When','Level' -EmptyText $(if ($SkipShareScan) { 'Share scan skipped (-SkipShareScan); Teams, Exchange and guest access are still listed when present.' } else { 'No shares or access grants were found in the sources this account can read.' }))
$(if ($shareRows.Count -gt 2000) { "<p class='note'>Showing the first 2,000 of $($shareRows.Count) rows (Urgent and Review first). The CSV export of tickets and the usage reports hold the rest.</p>" })
<h3>Cross-department access</h3>
<p class='meta'>Grants where the owner of the content and the person given access are in different departments (from the Department field in Entra ID). Expected between, for example, Lending and Compliance; worth a question when, for example, a branch user can read Finance or HR files.</p>
$(New-HtmlTable $crossDeptRows -Columns 'Owning department','Department with access','Shares','Examples' -EmptyText 'No cross-department grants were found (or departments are not populated in Entra ID).')
<h3>Guests and what they can reach</h3>
$(New-HtmlTable $guestAccessRows -Id 'tbl-guests' -Filters $GuestFilters -Columns 'Guest','Email domain','Enabled','Invited','Last sign-in','Member of (teams and groups)','Groups' -EmptyText 'No guest accounts exist in the directory.')
<h3>Calendars open to the organization or to anonymous users</h3>
$(if (-not $exoCollected) { "<p class='empty'>Exchange Online was not collected.</p>" } elseif ($SkipInboxRules) { "<p class='empty'>Calendar permissions were not scanned (-SkipInboxRules).</p>" } else { New-HtmlTable $calendarShareRows -Columns 'Mailbox','Level','Access rights' -EmptyText 'No calendar is readable beyond free/busy by everyone in the organization or by anonymous users.' })
</section>
"@

$html = @"
<!DOCTYPE html>
<html lang='en'><head><meta charset='utf-8'><title>$(HtmlEnc $ClientLabel) - Microsoft 365 Audit Evidence</title><style>$css</style></head>
<body>
<nav>$($navHtml.ToString())</nav>
<main>
<div class='pager'><button type='button' id='prevSec' title='Previous section'>&larr; Previous</button><button type='button' id='nextSec' title='Next section'>Next &rarr;</button><span class='title' id='pagerTitle'></span><span class='pos' id='pagerPos'></span><button type='button' id='viewToggle' title='Switch between one section per page and the whole report on one page'>Show all sections</button></div>
<section id='overview'><h2>1. Executive Summary</h2>
<p class='audience'>Prepared for the board and IT committee of $(HtmlEnc $ClientLabel) from a read-only review of the Microsoft 365 environment on $($ReportDate.ToString('MMMM d, yyyy')). One page answers three questions: are we secure, are we improving, what needs a decision. The detail behind every number follows for the security officer, the IT provider, auditors and examiners.</p>
<div class='verdict $($verdict.Class)'><div class='vlabel'>Overall</div><div class='vword'>$($verdict.Label)</div><p>$($verdict.Text)</p>
<div class='chips'><span class='chip $($trend.Class)'>Trend: $($trend.Label)</span><span class='chip $(if (-not $iocCollected) { 'mid' } elseif ($iocUrgentCount) { 'bad' } else { 'good' })'>Intrusion check: $(if (-not $iocCollected) { 'not collected' } elseif ($iocUrgentCount) { "$iocUrgentCount urgent" } else { 'nothing urgent' })</span><span class='chip $(if ($measured -and $onTrack -eq $measured) { 'good' } elseif ($measured -and $onTrack * 2 -ge $measured) { 'mid' } else { 'bad' })'>Scorecard: $onTrack of $measured on track</span><span class='chip neutral'>Maturity (indicative): $maturity</span><span class='chip neutral'>Next report: $($ReportDate.AddDays(30).ToString('MMMM d, yyyy'))</span></div></div>
<div class='kpis'>$($kpiHtml.ToString())</div>
<div class='eos'>
<div class='eos-col'>
<h3>Headlines</h3>
<ul class='headlines'>$(if ($headlines.Count) { (@($headlines | ForEach-Object { "<li class='$($_.Kind)'><span class='dot'></span>$(HtmlEnc $_.Text)</li>" }) -join '') } else { "<li class='good'><span class='dot'></span>Nothing to report beyond the scorecard.</li>" })</ul>
</div>
<div class='eos-col'>
<h3>Top priorities (next 30 days)</h3>
$(if ($remediationSorted.Count) { "<ol class='prio'>" + ((@($remediationSorted | Select-Object -First 5) | ForEach-Object { "<li><span class='badge $($_.Priority)'>$(HtmlEnc $_.Priority)</span><div>$(Get-FindingLinkHtml $_ -Ticket)<br><span class='who'>$(if ($_.Playbook) { "$(HtmlEnc (([string]$_.Playbook.Owner -split ';')[0])) &middot; $(HtmlEnc $_.Playbook.Effort)" } else { 'IT provider / MSP' }) &middot; $(HtmlEnc $_.Phase)</span></div></li>" }) -join '') + '</ol>' } else { "<p class='empty'>Nothing is due. Every measured control is in place.</p>" })
</div>
</div>
<h3>Scorecard</h3>
<p class='meta'>The measurables the committee tracks every month, each with its goal. Off track means work is needed; the (i) icon says what and why.</p>
$(New-HtmlTable $scorecardRows -Id 'tbl-scorecard' -Filters @(@{ Label = 'Off track'; Tag = 'offtrack' }, @{ Label = 'On track'; Tag = 'ontrack' }, @{ Label = 'Not measured'; Tag = 'notmeasured' }) -Columns 'Measurable','Goal','Actual','Status','Since last report' -EmptyText 'No measurables.')
<h3>Rocks: 90-day objectives</h3>
<p class='meta'>The findings grouped into business outcomes, each with an owner, a due date and a status. Done means no open item remains; On track means fewer open items than last month; Off track means more.</p>
$(New-HtmlTable $rockRows -Columns 'Rock (90-day objective)','Outcome','Owner','Open items','Due','Status' -EmptyText 'Every objective is complete.')
<h3>Issues list</h3>
<p class='meta'>Items that need a committee or board decision, or that block the work: licences and budget, policy choices, repeat findings and coverage gaps. Discuss, decide, minute.</p>
$(New-HtmlTable $issueRows -Columns 'Issue','Type','What the committee is asked to decide','Who','Phase' -EmptyText 'No open issues for the committee; the remaining work is for the IT provider.')
<h3 id='progress'>Progress</h3>
<p class='meta'>$(HtmlEnc $trend.Text)</p>
<div class='charts'><div class='chart'>$scoreTrend</div><div class='chart'>$findingTrend</div></div>
<details class='more'><summary>Show every change since the last report</summary>
$progressHtml
</details>
$(if ($repeatRows.Count) { "<details class='more' open><summary>Repeat findings: open for three or more reports</summary><p class='meta'>Examiners treat a finding that stays open across review cycles as a program weakness, not just a technical gap. Each needs an owner, a date and a reason in the minutes.</p>" + (New-HtmlTable $repeatRows -Columns 'Finding','Priority','Open in reports','First seen','Owner' -EmptyText '') + '</details>' })
<details class='more'><summary>More figures: users, devices, MFA</summary><div class='kpis'>$($kpiMoreHtml.ToString())</div></details>
$legendHtml
<hr class='divider'>
<h3>For the IT committee, auditors and examiners</h3>
<details class='more' open><summary>The plan by phase</summary>
$(New-HtmlTable $planRows -Columns 'Phase','Open items','Target date','Owners','Closed when' -EmptyText 'No findings.')
<p class='meta'>Goal: every item closed and verified by this report, then a monthly run to catch drift and indications of compromise, a quarterly committee review and an annual board summary. The script changes nothing in the tenant.</p>
</details>
<details class='more'><summary>How this report serves the annual GLBA board report and the IT examination</summary>
<p class='meta'>The Interagency Guidelines Establishing Information Security Standards (12 CFR Part 30 / 364 / 208 App. B, D-2; NCUA Part 748 App. A) require management to report to the board at least annually on the overall status of the information security program, covering risk assessment, risk management and control decisions, service provider arrangements, results of testing, security breaches and management responses, and recommendations for changes. This report supplies the Microsoft 365 portion of that evidence each month; core systems, vendors outside Microsoft 365, physical security and the written program remain part of the wider report.</p>
$(New-HtmlTable $glbaRows -Columns 'Board report element (GLBA Interagency Guidelines)','Where it is in this report')
<p class='meta'>What examiners look for in board reporting, and where it is here: evidence the board was told (this page, minuted); decisions the board made (the issues list); findings tracked with owners and dates (rocks, the plan by phase and section 4); closure proven rather than asserted (next report's progress); trends across cycles (the charts); and no repeat findings (the repeat-findings table).</p>
</details>
<details class='more'><summary>Summary in full and how to read this report by role</summary>
<p class='exec'>$execNarrative</p>
<div class='howto'><b>How to use this report.</b><ul>
<li><b>Board:</b> the verdict, the headlines and the scorecard answer three questions: are we secure, are we improving, is there a plan. Minute this page to record that the board was told.</li>
<li><b>IT committee:</b> the rocks, the issues list and the plan by phase are the agenda; approve owners, dates, licences and policy choices and minute them.</li>
<li><b>Security officer and IT provider / MSP:</b> work section 2 (review these now) today, then section 4 from the top. Every finding has step-by-step instructions, an owner and an effort estimate, and copies straight into a ticket.</li>
<li><b>Auditors:</b> section 3 is the control test, section 5 maps each control to NIST CSF 2.0, NIST SP 800-63B, the CIS Microsoft 365 Benchmark, CIS Controls v8 and FFIEC / NCUA expectations, and sections 6 to 19 hold the evidence behind every status.</li>
<li><b>Examiners:</b> the monthly history, the progress section, the repeat-findings table and the committee minutes show that the institution measures, reviews, plans and improves its Microsoft 365 controls.</li></ul></div>
</details>
<div class='privacy'><b>Privacy.</b> This report was generated locally by the free Audit Hat M365 Audit Evidence script ($brandLink) on the computer of the signed-in user. The script reads the tenant with that user's own permissions and writes only to the local output folder. No data is sent to Audit Hat or any third party.</div>
<h3>Governance and review</h3>
$(New-KvTable $governanceKv)
<h3>Tenant details</h3>
$(New-KvTable $tenantKv -Issues $tenantIssues)
<h3>Verified domains</h3>
$(New-HtmlTable $domainRows -Columns 'Domain','Default','Initial (.onmicrosoft.com)','Type','Capabilities' -EmptyText 'No verified domains returned.')
<h3>Subscriptions</h3>
$(New-HtmlTable $skuRows -Columns 'License SKU','Enabled','Consumed','Status')
</section>

$iocSectionHtml
<section id='summary'><h2>3. Security Controls Summary</h2>
<p class='meta'>Each status is computed from the tenant configuration collected below. Status reflects what is enforced today, not what is configured in report-only or disabled policies.</p>
$legendHtml
<div class='filters' data-table='tbl-controls'><button class='on' data-tag='all'>All ($(@($script:summaryRows).Count))</button><button data-tag='warn'>Action needed ($(@($script:summaryRows | Where-Object { (Get-StatusKind $_.Status) -eq 'warn' }).Count))</button><button data-tag='mid'>Partial ($(@($script:summaryRows | Where-Object { (Get-StatusKind $_.Status) -eq 'mid' }).Count))</button><button data-tag='good'>Pass ($(@($script:summaryRows | Where-Object { (Get-StatusKind $_.Status) -eq 'good' }).Count))</button><button data-tag='na'>Not measured ($(@($script:summaryRows | Where-Object { (Get-StatusKind $_.Status) -eq 'na' }).Count))</button><input type='search' placeholder='Search controls...' aria-label='Search'><span class='n'></span></div>
$($summaryTable.ToString())
</section>

<section id='remediation'><h2>4. Remediation Plan</h2>
<p class='meta'>Every control from the summary that is not fully met, plus every Microsoft Secure Score improvement action that still has points available, ranked by priority (Critical, High, Medium, Low); within a priority the measured controls come first, then Secure Score actions by points available. Findings: $(HtmlEnc $remediationCounts). $(if ($null -ne $scoreCurrent) { "Secure Score $scoreCurrent of $scoreMax points; target $scoreTarget points ($TargetScorePercent%); gap $scoreGap points." } else { 'Secure Score not available.' })$(if ($scoreIgnored) { " $scoreIgnored Secure Score action(s) marked ignored, third-party or resolved in Defender are not listed." })</p>
<h3>Roadmap</h3>
<p class='meta'>The same findings laid out as a plan: what to do this week, within 30 days, within 90 days, and over the year. Each phase lists the finding, its reference, who does it and how long it takes. Click a finding to jump to its step-by-step instructions below. Critical items are this week because they are either active exposure or the controls that stop account takeover and data leakage (MFA, consent, device compliance, guest and anonymous access, a second Global Administrator); High items close the gaps attackers use most; Medium items harden the environment; Low items are housekeeping and licence decisions for the committee.</p>
$($roadmapHtml.ToString())
<h3>All findings with step-by-step instructions</h3>
$ticketBar
$($remHtml.ToString())
<script type='application/json' id='tkData'>$ticketJson</script><script type='application/json' id='tkMeta'>$ticketMeta</script>
</section>
$frameworkSectionHtml
$mdmSection

<section id='profiles'><h2>7. Intune Policies</h2>
<p class='meta'>Configuration profiles, settings catalog policies, and endpoint security policies with the settings configured in each. Compliance policies are in section 4.</p>
<h3>Policy inventory</h3>
$(New-HtmlTable $profileRows -Columns 'Policy','Kind','Platform','Type','Last modified','Included','Excluded' -EmptyText $(if (Get-GraphFailure 'device configuration profiles') { "Intune policies could not be read: $(Get-GraphFailure 'device configuration profiles')" } else { 'No configuration profiles, settings catalog policies, or endpoint security policies exist in this tenant.' }))
<h3>Enrollment restrictions and enrollment configuration</h3>
$(New-HtmlTable $enrollRows -Columns 'Configuration','Type','Priority','Settings','Assigned to' -EmptyText 'No enrollment configurations returned.')
<h3>App protection (MAM) policies</h3>
$(New-HtmlTable $mamRows -Columns 'App protection policy','Platform','Assigned','Last modified','Key settings' -EmptyText 'No app protection policies exist in this tenant.')
<h3>Policy detail</h3>
$($profileDetailHtml.ToString())
</section>

<section id='devices'><h2>8. Managed Devices</h2>
$(New-KvTable $deviceStats -Issues $deviceStatIssues)
<h3>Device list</h3>
<p class='meta'>Every device object in Entra ID joined with its Intune record, plus Intune devices that have no Entra object. "Stale" means no sign-in or Intune sync in $StaleDays days. "Last user" is the most recent account Intune saw log on; "Owner / primary user" is the registered owner.</p>
$(New-HtmlTable $allDeviceRows -Id 'tbl-devices' -Filters $DeviceFilters -EmptyText 'No devices were returned from Entra ID or Intune.')
</section>

$caSection

<section id='authmethods'><h2>10. Authentication Methods</h2>
<h3>Security defaults</h3>
$(New-KvTable ([ordered]@{ 'Security defaults' = $(if ($null -ne $secDefaults.isEnabled) { $(if ($secDefaults.isEnabled) { 'Enabled' } else { 'Disabled' }) } else { 'Could not read' }) }) -Issues $secDefIssues)
<h3>Temporary Access Pass</h3>
$(New-KvTable $tapKv -Issues $tapIssues)
<h3>Authentication method policy</h3>
$(New-HtmlTable $authMethodRows -Columns 'Method','State','Included targets','Excluded targets','Settings' -EmptyText 'Could not read the authentication methods policy.')
<h3>MFA product and registered methods</h3>
<p class='meta'>Which product provides MFA (Microsoft Entra, a federated provider, or a third-party method), which methods users have actually registered, and how many rely only on phone-based methods that NIST SP 800-63B treats as restricted.</p>
$(New-KvTable $mfaProductKv -Issues $mfaProductIssues)
$(New-HtmlTable $methodCountRows -Id 'tbl-methods' -Filters $MethodFilters -Columns 'Method','Strength','Users registered','Share of enabled users' -EmptyText 'Registration report not available (requires Entra ID P1).')
<h3>Passkey and passwordless rollout</h3>
<p class='meta'>Phishing-resistant MFA (passkeys in Microsoft Authenticator, FIDO2 security keys, Windows Hello for Business, certificates) cannot be relayed by phishing kits. This tracks the rollout month by month; the Executive Summary charts the percentage.</p>
$(New-KvTable $passkeyKv -Issues $passkeyIssues)
<h3>Device registration settings</h3>
$(New-KvTable $devRegKv -Issues $devRegIssues)
</section>

<section id='apps'><h2>11. Enterprise Applications</h2>
<h3>Application governance settings</h3>
$(New-KvTable $appGovKv -Issues $appGovIssues)
<p class='note'>Approver and approval date come from the Entra directory audit log, which retains $AuditDays days in this export (30 days with Entra ID P1/P2, 7 days on Free). Applications consented before that window show the date they were added to the tenant and "Not in audit log retention" for the approver.</p>
<h3>Applications and consented permissions</h3>
<p class='meta'>Risk: High = application permissions that read or write mail, files or the directory without a user, a known-abused app, or an unverified publisher with mail or file access; Medium = user-consented permissions, admin-consented high-privilege delegated permissions, or an unverified publisher; Low = everything else. The first four permissions are listed; Show all opens the complete list grouped by type.</p>
$(New-HtmlTable $appRowsSorted -Id 'tbl-apps' -Filters $AppFilters -Columns 'Application','Risk','Permissions','Approved by / when','Added to tenant' -EmptyText 'No enterprise applications returned.')
</section>

<section id='consent'><h2>12. Application Consent Audit Events (last $AuditDays days)</h2>
$(New-HtmlTable $auditRows -Columns 'Date (local)','Activity','Performed by','Application','Admin consent','Result' -EmptyText "No application consent or add events were recorded in the last $AuditDays days.")
</section>

<section id='users'><h2>13. Users</h2>
$(New-KvTable $userStats -Issues $userStatIssues)
<h3>Enabled member accounts without MFA registered</h3>
$(if (@($regDetails).Count) { New-HtmlTable $noMfaRows -Columns 'User','UPN','Licensed','Last sign-in','Methods registered' -EmptyText 'Every enabled member account has an MFA method registered.' } else { "<p class='empty'>MFA registration report not available (requires Entra ID P1).</p>" })
<h3>User list</h3>
$(New-HtmlTable $userRows -Id 'tbl-users' -Filters $UserFilters -EmptyText 'No users returned.')
</section>

<section id='roles'><h2>14. Privileged Users</h2>
<p class='meta'>Active directory role assignments, grouped by role. Roles with no members are omitted. Groups assigned to a role are shown as a group row; their members are not expanded.</p>
$(New-KvTable $privStats)
$($privHtml.ToString())
$(if (@($activeRoles).Count -eq 0) { "<p class='empty'>No directory roles have active members.</p>" })
</section>

<section id='exchange'><h2>15. Exchange Online</h2>
<p class='meta'>Automatic forwarding and legacy protocol settings from Exchange Online PowerShell.</p>
$(New-KvTable $exoKv -Issues $exoIssues)
$exoDetailHtml
<h3>Email authentication: SPF, DKIM, DMARC</h3>
<p class='meta'>Public DNS records for every email-enabled custom domain, checked from this computer at report time, plus the DKIM signing state from Exchange Online. SPF and DKIM prove mail really came from your tenant; DMARC tells the world's mail servers to reject mail that fails and sends you reports of spoofing attempts.</p>
$(if (-not $dnsAvailable) { "<p class='empty'>DNS lookups are not available on this computer (Resolve-DnsName missing).</p>" } else { New-HtmlTable $emailAuthRows -Columns 'Domain','MX','SPF','SPF record','DKIM signing (Exchange)','DKIM DNS (selector1 / selector2)','DMARC policy','DMARC coverage','DMARC reporting (rua)','DMARC record' -EmptyText 'No email-enabled custom domains are verified in this tenant.' })
</section>

<section id='defender'><h2>16. Defender and Email Protection</h2>
<p class='meta'>Exchange Online Protection and Microsoft Defender for Office 365 policies (anti-spam, anti-malware, anti-phishing, Safe Links, Safe Attachments, preset security policies) and Microsoft Defender for Endpoint device status.</p>
$(New-KvTable ([ordered]@{
    'Defender for Office 365 licence'      = $(if ($script:HasDefenderP2) { 'Plan 2' } elseif ($script:HasDefenderP1) { 'Plan 1' } else { 'Not licensed: Exchange Online Protection only' })
    'Defender for Endpoint licence'        = $(if ($script:HasMDE) { $script:MdePlan } else { 'Not licensed' })
    'Preset security policies enabled'     = $(if ($defenderCollected) { "$presetOn of $(@($presetRows).Count)" } else { 'Not collected' })
    'Defender for SharePoint, OneDrive and Teams' = $(if ($atpGlobal) { $atpGlobal.EnableATPForSPOTeamsODB } else { 'Not collected' })
    'Safe Documents for Office clients'    = $(if ($atpGlobal) { $atpGlobal.EnableSafeDocs } else { 'Not collected' })
    'Connection filter: IP allow list entries / safe list' = $(if ($connFilter) { "$(@($connFilter.IPAllowList).Count) / $($connFilter.EnableSafeList)" } else { 'Not collected' })
}) -Issues $defenderKvIssues)
$(if ($exoCollected -and $defenderCollected) { @"
<h3>Preset security policies</h3>
$(New-HtmlTable $presetRows -Columns 'Preset','Scope','State','Priority','Applies to' -EmptyText 'No preset security policy rules exist (Standard and Strict presets are both off).')
<h3>Anti-spam (inbound) policies</h3>
$(New-HtmlTable $spamRows -Columns 'Policy','Default','Spam','High-confidence spam','Phishing','High-confidence phishing','Bulk threshold','Bulk action','ZAP (spam / phish)','Quarantine retention (days)','End-user notifications' -EmptyText 'No anti-spam policies returned.')
<h3>Anti-malware policies</h3>
$(New-HtmlTable $malwareRows -Columns 'Policy','Default','Common attachments filter','Blocked file types','File filter action','ZAP for malware','Notify admins (internal / external senders)' -EmptyText 'No anti-malware policies returned.')
<h3>Anti-phishing policies</h3>
$(New-HtmlTable $phishRows -Columns 'Policy','Default','Enabled','Spoof intelligence','Phishing threshold','Impersonation protection (users / domains / mailbox intelligence)','Protected users','Impersonation actions (user / domain / mailbox intelligence)','Safety tips (first contact / similar user / similar domain)','Unauthenticated sender indicator' -EmptyText 'No anti-phishing policies returned.')
<h3>Safe Links policies (Defender for Office 365)</h3>
$(if ($script:HasDefenderP1) { New-HtmlTable $safeLinksRows -Columns 'Policy','Email','Teams','Office apps','Scan URLs in real time','Wait for scan before delivery','Track clicks','Users may click through','Internal mail' -EmptyText 'Defender for Office 365 is licensed but no Safe Links policy exists.' } else { "<p class='empty'>Not licensed (requires Defender for Office 365).</p>" })
<h3>Safe Attachments policies (Defender for Office 365)</h3>
$(if ($script:HasDefenderP1) { New-HtmlTable $safeAttachRows -Columns 'Policy','Enabled','Action','Redirect blocked mail','Redirect to','If scan errors or times out' -EmptyText 'Defender for Office 365 is licensed but no Safe Attachments policy exists.' } else { "<p class='empty'>Not licensed (requires Defender for Office 365).</p>" })
"@ } else { "<p class='empty'>Exchange Online Protection and Defender for Office 365 policies were not collected ($(HtmlEnc $exoKv['Exchange Online'])).</p>" })
<h3 id='mde'>Endpoint protection: Defender for Endpoint</h3>
<p class='meta'>Whether workstations and laptops have endpoint detection and response (EDR): licensed, deployed to every active Windows and macOS device in Entra ID and Intune, and healthy.</p>
$mdeStatusHtml
$(if ($mdeCollected) { "<h4>Onboarded devices</h4>" + (New-KvTable $mdeKv) + (New-HtmlTable $mdeRows -Id 'tbl-mde' -Filters $MdeFilters -EmptyText 'No devices are onboarded to Defender for Endpoint.') + $(if ($mdeMissingRows.Count) { "<h4>Active Windows and macOS devices not onboarded ($($mdeMissingRows.Count))</h4><p class='meta'>Signed in or synced in the last $StaleDays days according to Entra ID or Intune, but not found in Defender for Endpoint by device name.</p>" + (New-HtmlTable $mdeMissingRows -Id 'tbl-mde-missing') } else { '' }) } else { '' })
</section>
$teamsSectionHtml
$spoSectionHtml
$shareSectionHtml
<div id='moreModal' class='modal' hidden><div class='modal-box' role='dialog' aria-modal='true' aria-labelledby='moreTitle'><div class='modal-head'><b id='moreTitle'></b><button type='button' id='moreClose'>Close</button></div><div class='modal-body' id='moreBody'></div></div></div>
<footer class='brand'>$(if ($logoImg) { "<a href='https://audithat.com' target='_blank' rel='noopener'>$logoImg</a>" })<div>Microsoft 365 Audit Evidence for $(HtmlEnc $ClientLabel), generated $($ReportDate.ToString('MMMM d, yyyy HH:mm')) by $(HtmlEnc $ctx.Account).<br>Produced locally by the free, read-only Audit Hat M365 Audit Evidence script, version $ScriptVersion (<a href='$RepoUrl' target='_blank' rel='noopener'>source and updates on GitHub</a>). Provided by $brandLink. No data left your organization.<br>Running on-premises Active Directory as well? The companion <a href='$AdRepoUrl' target='_blank' rel='noopener'>AuditHat AD Governance Report</a> covers it the same way.</div></footer>
<script>
(function(){
  document.querySelectorAll('.filters').forEach(function(bar){
    var table=document.getElementById(bar.getAttribute('data-table')); if(!table||!table.tBodies.length) return;
    var rows=Array.prototype.slice.call(table.tBodies[0].rows), active='all', q='', count=bar.querySelector('.n');
    function apply(){ var shown=0; rows.forEach(function(r){ var tags=' '+(r.getAttribute('data-tags')||'')+' ';
      var ok=(active==='all'||tags.indexOf(' '+active+' ')>=0)&&(!q||r.textContent.toLowerCase().indexOf(q)>=0);
      r.style.display=ok?'':'none'; if(ok) shown++; });
      if(count) count.textContent=shown+' of '+rows.length+' rows'; }
    bar.querySelectorAll('button').forEach(function(b){ b.addEventListener('click',function(){
      bar.querySelectorAll('button').forEach(function(x){x.classList.remove('on')}); b.classList.add('on'); active=b.getAttribute('data-tag'); apply(); }); });
    var inp=bar.querySelector('input'); if(inp) inp.addEventListener('input',function(){ q=inp.value.toLowerCase(); apply(); });
    apply();
  });
  // Remediation tickets: copy one finding, copy all, or download a CSV for the IT provider / MSP
  var tkEl=document.getElementById('tkData'), tkMetaEl=document.getElementById('tkMeta');
  if(tkEl&&tkMetaEl){
    var tk=JSON.parse(tkEl.textContent), meta=JSON.parse(tkMetaEl.textContent);
    function ticketText(t){
      var L=['['+t.Ref+'] '+t.Priority+': '+t.Finding,
             'Client: '+meta.Client,
             'Source: AuditHat M365 Audit Evidence report '+meta.Date+' (section: '+t['Report section']+')',
             'Priority: '+t.Priority,
             'Current state: '+t['Current state'],
             'Requested change: '+t['Recommended action']];
      if(t.Phase) L.push('Phase: '+t.Phase);
      if(t.Owner) L.push('Suggested owner: '+t.Owner+(t.Effort?' (effort: '+t.Effort+')':''));
      if(t.Prerequisite) L.push('Prerequisite: '+t.Prerequisite);
      if(t.Steps) L.push('Steps:\n'+t.Steps);
      if(t['Done when']) L.push('Done when: '+t['Done when']);
      if(t['Why it matters']) L.push('Why it matters / effort: '+t['Why it matters']);
      if(t.Standards) L.push('Standards: '+t.Standards);
      if(t['Secure Score points']) L.push('Microsoft Secure Score points available: '+t['Secure Score points']);
      if(t.Link) L.push('Reference: '+t.Link);
      L.push('Acceptance: the next AuditHat report lists '+t.Ref+' as resolved.');
      return L.join('\n');
    }
    function flash(b,msg){ var o=b.textContent; b.textContent=msg; b.disabled=true; setTimeout(function(){ b.textContent=o; b.disabled=false; },1600); }
    function copyText(s,b){
      function fallback(){ var ta=document.createElement('textarea'); ta.value=s; ta.setAttribute('readonly',''); ta.style.position='fixed'; ta.style.left='-9999px';
        document.body.appendChild(ta); ta.select(); var ok=false; try{ ok=document.execCommand('copy'); }catch(e){} document.body.removeChild(ta); flash(b, ok?'Copied':'Copy failed'); }
      if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(s).then(function(){ flash(b,'Copied'); },fallback); } else { fallback(); }
    }
    function csvCell(v){ v=(v==null?'':String(v)); return '"'+v.replace(/"/g,'""')+'"'; }
    var modal=document.getElementById('tkModal'), tkTextEl=document.getElementById('tkText'), tkTitleEl=document.getElementById('tkTitle'), tkCopyBtn=document.getElementById('tkCopy'), tkCloseBtn=document.getElementById('tkClose'), currentTicket='';
    function openTicket(t){ if(!modal) return; currentTicket=ticketText(t); tkTitleEl.textContent='Ticket '+t.Ref+': '+t.Finding; tkTextEl.textContent=currentTicket; modal.removeAttribute('hidden'); document.body.classList.add('noscroll'); tkCopyBtn.focus(); }
    function closeTicket(){ if(!modal||modal.hasAttribute('hidden')) return; modal.setAttribute('hidden',''); document.body.classList.remove('noscroll'); }
    document.querySelectorAll('button.tk').forEach(function(b){ b.addEventListener('click',function(){ var t=tk[parseInt(b.getAttribute('data-i'),10)]; if(t) openTicket(t); }); });
    if(tkCopyBtn) tkCopyBtn.addEventListener('click',function(){ copyText(currentTicket,tkCopyBtn); });
    if(tkCloseBtn) tkCloseBtn.addEventListener('click',closeTicket);
    if(modal) modal.addEventListener('click',function(e){ if(e.target===modal) closeTicket(); });
    document.addEventListener('keydown',function(e){ if(e.key==='Escape') closeTicket(); });
    var all=document.getElementById('tkAll'); if(all) all.addEventListener('click',function(){ copyText(tk.map(ticketText).join('\n\n----------------------------------------\n\n'),all); });
    var csv=document.getElementById('tkCsv'); if(csv) csv.addEventListener('click',function(){
      var cols=['Ref','Priority','Phase','Finding','Current state','Recommended action','Owner','Effort','Prerequisite','Steps','Done when','Standards','Why it matters','Secure Score points','Report section','Link'];
      var lines=[cols.concat(['Client','Report date']).map(csvCell).join(',')];
      tk.forEach(function(t){ lines.push(cols.map(function(c){ return csvCell(t[c]); }).concat([csvCell(meta.Client),csvCell(meta.Date)]).join(',')); });
      var blob=new Blob(['\ufeff'+lines.join('\r\n')],{type:'text/csv;charset=utf-8'}); var a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download=meta.File;
      document.body.appendChild(a); a.click(); setTimeout(function(){ URL.revokeObjectURL(a.href); document.body.removeChild(a); },500); flash(csv,'Downloaded');
    });
  }
  window.addEventListener('beforeprint', function(){ document.querySelectorAll('details.steps, details.more').forEach(function(d){ d.setAttribute('open',''); }); });
  // "Show all" dialogs for long cells (permissions): the full content sits hidden next to the button
  var moreModal=document.getElementById('moreModal'), moreBody=document.getElementById('moreBody'), moreTitle=document.getElementById('moreTitle'), moreClose=document.getElementById('moreClose');
  function closeMore(){ if(!moreModal||moreModal.hasAttribute('hidden')) return; moreModal.setAttribute('hidden',''); document.body.classList.remove('noscroll'); }
  document.querySelectorAll('button.more').forEach(function(b){ b.addEventListener('click',function(){ var src=document.getElementById(b.getAttribute('data-target')); if(!src||!moreModal) return; moreTitle.textContent=b.getAttribute('data-title')||''; moreBody.innerHTML=src.innerHTML; moreModal.removeAttribute('hidden'); document.body.classList.add('noscroll'); moreClose.focus(); }); });
  if(moreClose) moreClose.addEventListener('click',closeMore);
  if(moreModal) moreModal.addEventListener('click',function(e){ if(e.target===moreModal) closeMore(); });
  document.addEventListener('keydown',function(e){ if(e.key==='Escape') closeMore(); });
  // Paged view: the navigation shows one section at a time; "Show all sections" restores the single long page.
  var mainEl=document.querySelector('main');
  var sections=Array.prototype.slice.call(document.querySelectorAll('main > section'));
  var navLinks=Array.prototype.slice.call(document.querySelectorAll('nav a[href^="#"]'));
  var viewToggle=document.getElementById('viewToggle'), prevBtn=document.getElementById('prevSec'), nextBtn=document.getElementById('nextSec'), pagerTitle=document.getElementById('pagerTitle'), pagerPos=document.getElementById('pagerPos');
  var paged=true; try{ paged = (sessionStorage.getItem('audithat-view')!=='all'); }catch(e){}
  function sectionFor(id){ var el=document.getElementById(id); if(!el) return null; var sec=el; while(sec&&sec!==mainEl&&!(sec.tagName==='SECTION'&&sec.parentNode===mainEl)){ sec=sec.parentNode; } return (sec&&sec!==mainEl)?sec:null; }
  function showSection(sec, targetId){
    if(!sec) return;
    sections.forEach(function(s){ s.classList.toggle('active', s===sec); });
    navLinks.forEach(function(a){ a.classList.toggle('active', a.getAttribute('href')==='#'+sec.id); });
    var idx=sections.indexOf(sec), h=sec.querySelector('h2');
    if(pagerTitle) pagerTitle.textContent = h?h.textContent:'';
    if(pagerPos) pagerPos.textContent='Section '+(idx+1)+' of '+sections.length;
    if(prevBtn) prevBtn.disabled=(idx<=0); if(nextBtn) nextBtn.disabled=(idx>=sections.length-1);
    if(targetId&&targetId!==sec.id){ var t=document.getElementById(targetId); if(t){ t.scrollIntoView(); window.scrollBy(0,-60); } } else { window.scrollTo(0,0); }
  }
  function applyMode(){
    document.body.classList.toggle('paged', paged);
    if(viewToggle) viewToggle.textContent = paged?'Show all sections':'One section at a time';
    if(prevBtn) prevBtn.style.display=paged?'':'none'; if(nextBtn) nextBtn.style.display=paged?'':'none'; if(pagerPos) pagerPos.style.display=paged?'':'none';
    if(paged){ var id=(location.hash||'#overview').substring(1); var sec=sectionFor(id)||sections[0]; showSection(sec,id); }
    else { sections.forEach(function(s){ s.classList.add('active'); }); if(pagerTitle) pagerTitle.textContent='Whole report on one page'; navLinks.forEach(function(a){ a.classList.remove('active'); }); }
  }
  document.addEventListener('click', function(e){
    var a=e.target; while(a&&a!==document&&!(a.tagName==='A'&&a.getAttribute('href')&&a.getAttribute('href').charAt(0)==='#')){ a=a.parentNode; }
    if(!a||a===document||!paged) return;
    var id=a.getAttribute('href').substring(1); var sec=sectionFor(id); if(!sec) return;
    e.preventDefault(); try{ history.replaceState(null,'','#'+id); }catch(err){} showSection(sec,id);
  });
  window.addEventListener('hashchange', function(){ if(!paged) return; var id=location.hash.substring(1); var sec=sectionFor(id); if(sec) showSection(sec,id); });
  if(viewToggle) viewToggle.addEventListener('click', function(){ paged=!paged; try{ sessionStorage.setItem('audithat-view', paged?'paged':'all'); }catch(e){} applyMode(); });
  if(prevBtn) prevBtn.addEventListener('click', function(){ var i=sections.findIndex(function(s){ return s.classList.contains('active'); }); if(i>0){ try{ history.replaceState(null,'','#'+sections[i-1].id); }catch(err){} showSection(sections[i-1]); } });
  if(nextBtn) nextBtn.addEventListener('click', function(){ var i=sections.findIndex(function(s){ return s.classList.contains('active'); }); if(i<sections.length-1){ try{ history.replaceState(null,'','#'+sections[i+1].id); }catch(err){} showSection(sections[i+1]); } });
  window.addEventListener('beforeprint', function(){ document.body.classList.remove('paged'); });
  window.addEventListener('afterprint', function(){ document.body.classList.toggle('paged', paged); });
  function setPagerHeight(){ var pg=document.querySelector('.pager'); if(pg) document.documentElement.style.setProperty('--pager-h', (pg.offsetHeight+1)+'px'); }
  setPagerHeight(); window.addEventListener('resize', setPagerHeight);
  applyMode();
  // one info panel open at a time; click anywhere else closes it
  document.addEventListener('click', function(e){
    var open=document.querySelectorAll('details.info[open]');
    open.forEach(function(d){ if(!d.contains(e.target)) d.removeAttribute('open'); });
  });
})();
</script>
</main></body></html>
"@

$reportPath = Join-Path $OutputFolder "$safeClient - M365 Audit Evidence $($ReportDate.ToString('yyyy-MM-dd')).html"
[System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
Write-Host "HTML report: $reportPath" -ForegroundColor Green
if ($ticketRows.Count) {
    $ticketCsvPath = Join-Path $OutputFolder $ticketCsvName
    $ticketRows | Select-Object *, @{ Name = 'Client'; Expression = { $ClientLabel } }, @{ Name = 'Report date'; Expression = { $ReportDate.ToString('yyyy-MM-dd') } } |
        Export-Csv -Path $ticketCsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "Remediation tickets (CSV): $ticketCsvPath" -ForegroundColor Green
}
if (($SaveSignIn -or $script:SignInFromCache) -and $script:GraphToken -and $script:GraphToken.RefreshToken) { Save-SignInCache }   # the token rotates during the run; keep the newest
if ($script:ConsentNeeded.Count) {
    Write-Host ''
    Write-Warning "One-time approval needed in this tenant: the connector's read-only permissions for $(Join-List @($script:ConsentNeeded.Keys) ' and ') are not approved here, so those parts of the report are incomplete. A Global Administrator opens each link below once and clicks Accept. The browser then shows a localhost page that does not load; that is expected. Run the report again afterwards."
    foreach ($k in $script:ConsentNeeded.Keys) { Write-Host "  $k :" -ForegroundColor Yellow; Write-Host "  $(Get-ConsentLink $script:ConsentNeeded[$k])" }
}
if (-not $NoOpen -and -not $NonInteractive) { try { Start-Process -FilePath $reportPath -ErrorAction Stop } catch { Write-Verbose "Could not open the report automatically: $($_.Exception.Message)" } }

# ----------------------------------------------------------------------------------------------
# Optional Word documents (requires Microsoft Word installed on this machine)
# ----------------------------------------------------------------------------------------------
if ($WordDocs) {
    $docCss = $css -replace 'nav\{[^}]*\}', '' -replace 'main\{[^}]*\}', 'main{padding:0}'
    $docs = @(
        @{ Name = 'MDM Configuration';            Title = 'Intune MDM Compliance Policy Configuration'; Body = $mdmSection },
        @{ Name = 'Conditional Access Policies';  Title = 'Conditional Access Policy Configuration';    Body = $caSection }
    )
    $word = $null
    try { $word = New-Object -ComObject Word.Application; $word.Visible = $false } catch { Write-Warning "Microsoft Word is not available on this machine; HTML versions of the two documents were written instead." }
    foreach ($d in $docs) {
        $docHtml = @"
<!DOCTYPE html><html><head><meta charset='utf-8'><title>$(HtmlEnc $d.Title)</title><style>$docCss</style></head><body><main>
<p style='font-size:20pt;color:#005D9C;font-weight:bold;margin:0'>$(HtmlEnc $ClientLabel)</p>
<p style='font-size:14pt;color:#005D9C;margin:0 0 4px'>$(HtmlEnc $d.Title)</p>
<p class='meta'>Exported from the Microsoft 365 tenant on $($ReportDate.ToString('MMMM d, yyyy HH:mm')) by $(HtmlEnc $ctx.Account)</p>
$($d.Body)
</main></body></html>
"@
        $htmlPath = Join-Path $OutputFolder "$($d.Name).html"
        [System.IO.File]::WriteAllText($htmlPath, $docHtml, [System.Text.Encoding]::UTF8)
        if ($word) {
            $docxPath = Join-Path $OutputFolder "$($d.Name).docx"
            try {
                $doc = $word.Documents.Open($htmlPath, $false, $true)
                $doc.SaveAs2($docxPath, 16)   # 16 = wdFormatXMLDocument; plain arguments work on both 5.1 and 7.x
                $doc.Close($false)
                Remove-Item $htmlPath -Force
                Write-Host "Word document: $docxPath" -ForegroundColor Green
            } catch { Write-Warning "Word conversion failed for $($d.Name): $($_.Exception.Message). HTML left at $htmlPath" }
        } else { Write-Host "HTML document: $htmlPath" }
    }
    if ($word) { $word.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}

Disconnect-MgGraph | Out-Null
Write-Host "Done." -ForegroundColor Green



