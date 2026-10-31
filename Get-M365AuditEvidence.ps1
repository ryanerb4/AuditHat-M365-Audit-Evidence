<#
.SYNOPSIS
  Collects Microsoft 365 / Entra ID / Intune audit evidence and builds one self-contained HTML
  report with a left navigation bar. Optionally also produces two Word documents:
    - MDM Configuration.docx              (Intune compliance policies for Android and iOS/iPadOS)
    - Conditional Access Policies.docx    (all Conditional Access policies)

.DESCRIPTION
  Sections produced:
    1. Tenant overview
    2. Intune compliance policies (settings, actions for noncompliance, assignments) grouped by platform
    3. Intune device configuration profiles (list and assignments)
    4. Conditional Access policies (full conditions, grant and session controls)
    5. Enterprise applications (publisher, consent type, delegated and application permissions,
       date added, who added or consented when available in the audit log)
    6. Application consent and add events from the directory audit log
    7. Users (status, type, licenses, last sign-in, MFA registration)
    8. Privileged role members

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
    [string]$TenantId,             # tenant GUID or domain (contoso.onmicrosoft.com); required when -ClientId is a single-tenant app registration
    [string]$ClientId,             # override the default Audit Hat public-client app with your own registration (README "Register your own app")
    [string]$OutputFolder = ".\M365-Audit-Evidence",
    [int]$AuditDays       = 30,
    [int]$TargetScorePercent = 80,   # target Microsoft Secure Score as a percentage of the maximum achievable score
    [switch]$WordDocs,
    [switch]$SkipExchange,         # skip the Exchange Online checks (auto-forwarding, POP/IMAP/SMTP AUTH)
    [switch]$SkipInboxRules,       # skip per-mailbox inbox rule scan (slow on large tenants)
    [switch]$DeviceCode            # sign in with a device code instead of the browser (for remote sessions or if the browser flow fails)
)

$ErrorActionPreference = 'Stop'
$ReportDate = Get-Date
$ScriptVersion = '2026.10.31'   # date-based version: year.month.day of the last change
$RepoUrl = 'https://github.com/ryanerb4/AuditHat-M365-Audit-Evidence'

# Optional settings file next to the script so -ClientId / -TenantId do not have to be typed every run:
#   { "ClientId": "<app registration client id>", "TenantId": "<tenant id or domain>" }
$settingsPath = Join-Path $PSScriptRoot 'M365AuditEvidence.settings.json'
if (Test-Path $settingsPath) {
    try {
        $settings = Get-Content $settingsPath -Raw | ConvertFrom-Json
        if (-not $ClientId -and $settings.ClientId) { $ClientId = [string]$settings.ClientId }
        if (-not $TenantId -and $settings.TenantId) { $TenantId = [string]$settings.TenantId }
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
    'SecurityEvents.Read.All'          # Microsoft Secure Score (read-only)
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
Write-Host "connector you consent to only grants this local script read-only rights; it has no server side and receives nothing." -ForegroundColor DarkCyan
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
$script:AuthTenant    = if ($TenantId) { $TenantId } else { 'organizations' }   # 'organizations' excludes personal Microsoft accounts
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
function Request-ResourceToken {
    # Uses the refresh token from the browser sign-in to get a delegated token for another Microsoft API (same app, same user).
    param([string]$ScopeStr)
    if (-not $script:GraphToken -or -not $script:GraphToken.RefreshToken) { return $null }
    $body = @{ grant_type = 'refresh_token'; refresh_token = $script:GraphToken.RefreshToken; client_id = $script:GraphClientId; scope = $ScopeStr }
    $r = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$($script:AuthTenant)/oauth2/v2.0/token" -Body $body -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop
    if ($r.refresh_token) { $script:GraphToken.RefreshToken = $r.refresh_token }
    return $r.access_token
}
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

$signedIn = $false
if (-not $DeviceCode) { $signedIn = Connect-GraphBrowser }
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
$missingScopes = @($Scopes | Where-Object { $ctx.Scopes -notcontains $_ })
if ($missingScopes.Count) {
    Write-Warning "The following Graph permissions were not granted; related sections will show as not available: $($missingScopes -join ', ')"
    Write-Warning "A Global Administrator must run the script once and accept the consent prompt to grant them."
}

if (-not (Test-Path $OutputFolder)) { New-Item -ItemType Directory -Path $OutputFolder | Out-Null }
$OutputFolder = (Resolve-Path $OutputFolder).Path

# ----------------------------------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------------------------------
function Get-GraphAll {
    param([Parameter(Mandatory)][string]$Uri, [switch]$NoPaging)
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
        $next = if ($NoPaging) { $null } else { $r.'@odata.nextLink' }
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
    param([string]$Uri, [string]$What, [switch]$NoPaging)
    try { return (Get-GraphAll -Uri $Uri -NoPaging:$NoPaging) }
    catch {
        $reason = Get-GraphErrorText $_
        # Graph returns a bare 400/403 for unlicensed features; explain from the tenant's own subscriptions when we can.
        if ($Uri -match 'deviceManagement|deviceAppManagement' -and $script:HasIntune -eq $false) { $reason += ' Intune is not licensed in this tenant (no Intune service plan in the subscribed SKUs).' }
        elseif ($Uri -match 'userRegistrationDetails|signInActivity' -and $script:HasEntraP1 -eq $false) { $reason += ' Requires Entra ID P1 or P2, which is not licensed in this tenant.' }
        elseif ($Uri -match 'roleEligibilitySchedule|roleAssignmentSchedule' -and $script:HasEntraP2 -eq $false) { $reason += ' Privileged Identity Management requires Entra ID P2, which is not licensed in this tenant.' }
        $script:GraphFailures[$What] = $reason
        Write-Warning "Could not retrieve $What : $reason"
        return @()
    }
}
function Get-GraphFailure { param([string]$What) if ($script:GraphFailures.ContainsKey($What)) { return $script:GraphFailures[$What] } ; return $null }

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
    'guest'                = @{ Title = 'Guest account';                                     Why = 'Guests are external people with access to shared content; stale or unreviewed guests are a common data-leak path.'; Fix = 'Review guests periodically, remove those no longer needed, and consider access reviews or restricting guest invitations.' }
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
    $Columns = @($Columns | Where-Object { $_ -notin @('_tags', '_issues') })
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
        [void]$sb.Append($(if ($tags) { "<tr data-tags='$(HtmlEnc $tags)'>" } else { '<tr>' }))
        foreach ($c in $Columns) {
            $v = if ($r -is [System.Collections.IDictionary]) { $r[$c] } else { $r.$c }
            if ($v -is [System.Collections.IEnumerable] -and $v -isnot [string]) { $v = Join-List $v '; ' }
            $cell = HtmlEnc $v
            $cls = ''
            switch -Regex ([string]$v) {
                '^(True|Enabled|enabled|Compliant|Required|Yes|Block|block|On)$' { $cls = " class='good'" }
                '^(False|Disabled|disabled|No|Not registered|Off|enabledForReportingButNotEnforced)$' { $cls = " class='warn'" }
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
$tenantKv['Licensed capabilities'] = @(
    "Microsoft Intune (device management): $(if ($script:HasIntune) { 'Licensed' } else { 'Not licensed' })",
    "Entra ID P1 (Conditional Access, MFA registration report, sign-in logs): $(if ($script:HasEntraP1) { 'Licensed' } else { 'Not licensed' })",
    "Entra ID P2 (Privileged Identity Management, Identity Protection): $(if ($script:HasEntraP2) { 'Licensed' } else { 'Not licensed' })",
    "Defender for Office 365 (Safe Links, Safe Attachments, impersonation protection): $(if ($script:HasDefenderP2) { 'Licensed (Plan 2)' } elseif ($script:HasDefenderP1) { 'Licensed (Plan 1)' } else { 'Not licensed (Exchange Online Protection only)' })",
    "Defender for Endpoint / Defender for Business (endpoint detection and response): $(if ($script:HasMDE) { 'Licensed' } else { 'Not licensed' })"
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

# ----------------------------------------------------------------------------------------------
# 11. Exchange Online: external auto-forwarding and legacy protocols
# ----------------------------------------------------------------------------------------------
$exoKv = [ordered]@{}
$fwdMailboxRows = @(); $fwdRuleRows = @(); $fwdTransportRows = @(); $remoteDomainRows = @(); $casRows = @()
$dkimConfigs = @(); $spamRows = @(); $malwareRows = @(); $phishRows = @(); $safeLinksRows = @(); $safeAttachRows = @(); $presetRows = @(); $atpGlobal = $null; $connFilter = $null; $defenderCollected = $false
$exoCollected = $false
if (-not $SkipExchange) {
    if (Get-Module -ListAvailable ExchangeOnlineManagement) {
        try {
            Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
            Import-Module ExchangeOnlineManagement
            $exoParams = @{ ShowBanner = $false }
            if ($TenantId) { $exoParams.Organization = $TenantId }
            if ($ctx.Account) { $exoParams.UserPrincipalName = $ctx.Account }   # reuse the same account signed in to Graph
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
                        $rules = Get-InboxRule -Mailbox $mb.Identity -ErrorAction Stop | Where-Object { $_.ForwardTo -or $_.ForwardAsAttachmentTo -or $_.RedirectTo }
                        foreach ($r in $rules) {
                            $targets = @(@($r.ForwardTo) + @($r.ForwardAsAttachmentTo) + @($r.RedirectTo) | Where-Object { $_ })
                            $ext = @($targets | Where-Object { Test-ExternalAddress ([string]$_) }).Count -gt 0
                            $fwdRuleRows += [ordered]@{ 'Mailbox' = $mb.PrimarySmtpAddress; 'Rule' = $r.Name; 'Enabled' = $r.Enabled; 'Forwards to' = (Join-List ($targets | ForEach-Object { [string]$_ }) '; '); 'External' = $(if ($ext) { 'Yes' } else { 'No' }) }
                        }
                    } catch { Write-Warning "Inbox rules not readable for $($mb.PrimarySmtpAddress): $($_.Exception.Message)" }
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

            $exoKv = [ordered]@{
                'Default remote domain (*) allows auto-forward'   = $(if ($defaultRd) { $defaultRd.AutoForwardEnabled } else { 'Not found' })
                'Outbound spam policy automatic forwarding'       = $spamText
                'Mailboxes forwarding to an external address'     = @($fwdMailboxRows | Where-Object { $_.External -eq 'Yes' }).Count
                'Inbox rules forwarding to an external address'   = $(if ($SkipInboxRules) { 'Not scanned' } else { @($fwdRuleRows | Where-Object { $_.External -eq 'Yes' }).Count })
                'Transport rules that redirect, copy, or bcc mail' = @($fwdTransportRows).Count
                'Modern authentication (OAuth2) enabled'          = $orgConfig.OAuth2ClientProfileEnabled
                'SMTP AUTH disabled tenant-wide'                  = $transportConfig.SmtpClientAuthenticationDisabled
                'Mailboxes with POP enabled'                      = @($casMailboxes | Where-Object { $_.PopEnabled }).Count
                'Mailboxes with IMAP enabled'                     = @($casMailboxes | Where-Object { $_.ImapEnabled }).Count
                'Default authentication policy'                   = $(if ($defaultAuthPolicy) { [string]$defaultAuthPolicy } else { 'None set' })
                'Authentication policies defined'                 = $(if ($authPolicies) { Join-List ($authPolicies | ForEach-Object { $ap = $_; $basic = @($ap.PSObject.Properties | Where-Object { $_.Name -like 'AllowBasicAuth*' -and $_.Value -eq $true } | ForEach-Object { $_.Name -replace 'AllowBasicAuth', '' }); "$($ap.Name) (basic auth allowed: $(if ($basic.Count) { $basic -join ',' } else { 'none' }))" }) '; ' } else { 'None' })
            }
        } catch {
            Write-Warning "Exchange Online collection failed: $($_.Exception.Message)"
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
$emailAuthComplete = @($emailAuthRows | Where-Object { -not $_['_issues'] }).Count
$emailAuthStatus = if (-not $dnsAvailable) { 'Not available' } elseif ($emailAuthRows.Count -eq 0) { 'No email domains' } elseif ($emailAuthComplete -eq $emailAuthRows.Count) { 'Complete' } elseif (@($emailAuthRows | Where-Object { $_['DMARC policy'] -ne 'Missing' -or $_['SPF'] -ne 'Missing' }).Count) { 'Partial' } else { 'Missing' }

# ----------------------------------------------------------------------------------------------
# Defender for Endpoint (separate API; delegated token for the same user, only after a browser sign-in)
# ----------------------------------------------------------------------------------------------
$mdeRows = @(); $mdeKv = [ordered]@{}; $mdeCollected = $false; $mdeNote = ''
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
                'Devices onboarded to Defender for Endpoint' = $mdeRows.Count
                'Sensor health: active / inactive / impaired' = "$(@($mdeRows | Where-Object { $_['Sensor health'] -eq 'Active' }).Count) / $(@($mdeRows | Where-Object { $_['Sensor health'] -eq 'Inactive' }).Count) / $(@($mdeRows | Where-Object { $_['Sensor health'] -like 'Impaired*' }).Count)"
                'Devices at high risk (active alerts)'       = @($mdeRows | Where-Object { $_['Risk'] -eq 'High' }).Count
                'Devices with high exposure (vulnerabilities)' = @($mdeRows | Where-Object { $_['Exposure'] -eq 'High' }).Count
                'Organization exposure score (0-100, lower is better)' = $(if ($null -ne $expo) { [math]::Round([double]$expo, 1) } else { 'Not available' })
                'Microsoft Secure Score for Devices (configuration score)' = $(if ($null -ne $conf) { [math]::Round([double]$conf, 1) } else { 'Not available' })
            }
            $mdeIntuneWin = @($allDeviceRows | Where-Object { $_['Intune'] -eq 'Enrolled' -and $_['OS'] -match 'Windows' }).Count
            if ($mdeIntuneWin -gt $mdeRows.Count) { $mdeKv['Intune Windows devices not seen by Defender for Endpoint'] = $mdeIntuneWin - $mdeRows.Count }
        } else { $mdeNote = 'Not collected: device-code sign-in was used (no refresh token for the Defender API).' }
    } catch {
        $msg = $_.Exception.Message
        try { $body = $_.ErrorDetails.Message; if ($body) { $j = $body | ConvertFrom-Json; if ($j.error_description) { $msg = "$msg $($j.error_description)" } elseif ($j.error.message) { $msg = "$msg $($j.error.message)" } } } catch { Write-Verbose 'Defender error body was not JSON.' }
        $mdeNote = if ($msg -match 'AADSTS65001|AADSTS650|consent') { "The Defender for Endpoint permissions (Machine.Read, Software.Read, Vulnerability.Read) have not been consented for this app in the tenant. A Global Administrator can grant them at https://login.microsoftonline.com/$($ctx.TenantId)/adminconsent?client_id=$ClientId" } elseif ($msg -match '403|Forbidden|Unauthorized|401') { "Defender for Endpoint API refused the request ($($msg -replace '\s+', ' ')). The tenant is not licensed for Defender for Endpoint / Defender for Business, or the signed-in account lacks a Defender role (Security Reader)." } else { "Not collected: $($msg -replace '\s+', ' ')" }
        Write-Warning "Defender for Endpoint: $mdeNote"
    }
} else { $mdeNote = 'Not collected: device-code sign-in was used (no refresh token for the Defender API).' }
$MdeFilters = @(@{ Label = 'High risk'; Tag = 'highrisk' }, @{ Label = 'High exposure'; Tag = 'highexposure' }, @{ Label = 'Inactive sensor'; Tag = 'inactive' }, @{ Label = 'Stale'; Tag = 'stale' })

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
$mdeStatus = if ($mdeCollected) { $(if (@($mdeRows | Where-Object { $_['Risk'] -eq 'High' -or $_['Sensor health'] -ne 'Active' }).Count) { 'Attention needed' } elseif ($mdeRows.Count) { 'Onboarded' } else { 'No devices onboarded' }) } elseif ($script:HasMDE -eq $false) { 'Not licensed' } else { 'Not available' }
$mdeEvidence = if ($mdeCollected) { "$($mdeRows.Count) device(s) onboarded; $($mdeKv['Devices at high risk (active alerts)']) at high risk; sensors active/inactive/impaired $($mdeKv['Sensor health: active / inactive / impaired']); exposure score $($mdeKv['Organization exposure score (0-100, lower is better)'])." } elseif ($script:HasMDE -eq $false) { 'No Defender for Endpoint or Defender for Business service plan in the subscribed SKUs.' } else { $mdeNote }

$script:summaryRows = @()
function Add-SummaryRow { param($control, $status, $evidence, $link) $script:summaryRows += [ordered]@{ 'Control' = $control; 'Status' = $status; 'Evidence' = $evidence; 'Section' = $link } }

Add-SummaryRow 'Security defaults' $(if ($secDefaults.isEnabled) { 'Enabled' } else { 'Disabled' }) $(if ($secDefaults.isEnabled) { 'Security defaults enforce MFA registration and block legacy authentication tenant-wide; Conditional Access cannot be used alongside them.' } else { 'Security defaults are off; MFA and legacy authentication controls depend on Conditional Access.' }) 'Authentication Methods'

$mfaStatus = if ($secDefaults.isEnabled) { 'Enforced (security defaults)' } elseif ($mfaAllUsersEnforced.Count) { 'Enforced' } elseif (@($mfaPolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not enforced' }
$mfaEvidence = if ($mfaAllUsersEnforced.Count) { "Enabled policy covering all users and all cloud apps: " + (Join-List ($mfaAllUsersEnforced | ForEach-Object { "$($_.displayName) [$(Get-CaScopeText $_)]" }) ' || ') }
               elseif ($mfaPolicies.Count) { "MFA policies exist but none that is enabled targets all users and all cloud apps: " + (Join-List ($mfaPolicies | ForEach-Object { "$($_.displayName) ($($_.state))" }) '; ') }
               elseif ($secDefaults.isEnabled) { 'Security defaults enabled.' } else { 'No Conditional Access policy requires MFA and security defaults are off.' }
Add-SummaryRow 'MFA enforced for sign-in' $mfaStatus $mfaEvidence 'Conditional Access'

Add-SummaryRow 'MFA registered by all users' $(if (-not @($regDetails).Count) { 'Not available' } elseif ($mfaRegistered -eq $enabledMembers.Count) { 'Yes' } else { 'No' }) $(if (@($regDetails).Count) { "$mfaRegistered of $($enabledMembers.Count) enabled member accounts have an MFA method registered; $(@($noMfaRows).Count) do not (listed under Users)." } else { 'Registration report requires Entra ID P1.' }) 'Users'

$legacyStatus = if ($secDefaults.isEnabled) { 'Blocked (security defaults)' } elseif ($legacyBlockedAll.Count) { 'Blocked' } elseif (@($legacyPolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not blocked by Conditional Access' }
$legacyEvidence = if ($legacyBlockedAll.Count) { "Block policy for legacy clients covering all users: " + (Join-List ($legacyBlockedAll | ForEach-Object { $_.displayName }) '; ') } elseif ($legacyPolicies.Count) { "Legacy block policies exist but are scoped or not enabled: " + (Join-List ($legacyPolicies | ForEach-Object { "$($_.displayName) ($($_.state))" }) '; ') } elseif ($secDefaults.isEnabled) { 'Security defaults block legacy authentication protocols tenant-wide (no Conditional Access policy is needed).' } else { 'No Conditional Access policy blocks Exchange ActiveSync and other (legacy) clients.' }
if ($exoCollected) { $legacyEvidence += " Exchange: SMTP AUTH disabled tenant-wide = $($exoKv['SMTP AUTH disabled tenant-wide']); POP enabled on $($exoKv['Mailboxes with POP enabled']) mailboxes; IMAP enabled on $($exoKv['Mailboxes with IMAP enabled']) mailboxes." }
Add-SummaryRow 'Legacy authentication protocols disabled' $legacyStatus $legacyEvidence 'Conditional Access'

$devStatus = if ($deviceEnforcedAll.Count) { 'Enforced' } elseif (@($devicePolicies | Where-Object { $_.state -eq 'enabled' }).Count) { 'Partial' } else { 'Not enforced' }
$devEvidence = if ($devicePolicies.Count) { Join-List ($devicePolicies | ForEach-Object { "$($_.displayName) ($($_.state)): $(Get-CaScopeText $_)" }) ' || ' } else { 'No Conditional Access policy requires a compliant or Entra joined device; any device with valid credentials can access Microsoft 365.' }
Add-SummaryRow 'Conditional Access restricts access to Entra joined or compliant devices' $devStatus $devEvidence 'Conditional Access'

Add-SummaryRow 'Intune deployed to Windows computers' $(if (Get-GraphFailure 'Intune managed devices') { 'Not available' } elseif ($winEntra.Count -eq 0) { 'No Windows devices in Entra' } elseif ($winManaged.Count -eq $winEntra.Count) { 'All enrolled' } elseif ($winManaged.Count -gt 0) { 'Partial' } else { 'Not deployed' }) "$(if (Get-GraphFailure 'Intune managed devices') { "Intune device list could not be read. $(Get-GraphFailure 'Intune managed devices') " })$($winManaged.Count) of $($winEntra.Count) enabled Windows devices in Entra ID are enrolled in Intune; $($winEntra.Count - $winManaged.Count) are not (listed under Managed Devices). Total Intune-enrolled devices: $(@($managedDevices).Count)." 'Managed Devices'

Add-SummaryRow 'Intune compliance policies' $(if (Get-GraphFailure 'compliance policies') { 'Not available' } elseif (@($compliancePolicies).Count) { "$(@($compliancePolicies).Count) policies" } else { 'None' }) $(if (Get-GraphFailure 'compliance policies') { "Could not be read. $(Get-GraphFailure 'compliance policies')" } elseif (@($compliancePolicies).Count) { Join-List ($compliancePolicies | ForEach-Object { "$($_.displayName) [$(Get-Platform $_.'@odata.type')]" }) '; ' } else { 'No compliance policies exist, so no device can be evaluated as compliant.' }) 'MDM Compliance Policies'

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

Add-SummaryRow 'Email authentication (SPF, DKIM, DMARC)' $emailAuthStatus $(if ($emailAuthRows.Count) { Join-List ($emailAuthRows | ForEach-Object { "$($_['Domain']): SPF $($_['SPF']); DKIM $($_['DKIM signing (Exchange)']) / $($_['DKIM DNS (selector1 / selector2)']); DMARC $($_['DMARC policy'])" }) ' || ' } elseif (-not $dnsAvailable) { 'Resolve-DnsName is not available on this computer.' } else { 'No email-enabled custom domains.' }) 'Exchange Online'

Add-SummaryRow 'Email protection (Defender for Office 365 / EOP)' $defenderStatus $defenderEvidence 'Defender'

Add-SummaryRow 'Endpoint protection (Defender for Endpoint)' $mdeStatus $mdeEvidence 'Defender'

Add-SummaryRow 'Microsoft Secure Score' $(if ($null -eq $scoreCurrent) { 'Not available' } elseif ($scoreCurrent -ge $scoreTarget) { 'Target met' } else { 'Below target' }) $(if ($null -ne $scoreCurrent) { "Current $scoreCurrent of $scoreMax points ($scorePct%); target $scoreTarget points ($TargetScorePercent%); gap $scoreGap points.$(if ($scoreCompareText) { " Comparison: $scoreCompareText." })" } else { "Could not be read. $(Get-GraphFailure 'Microsoft Secure Score')" }) 'Tenant Overview'

Add-SummaryRow 'Privileged role assignments' "$(@($roleRows).Count) assignments" $(if ($roles) { Join-List (($roles | Where-Object { @($_.members).Count -gt 0 } | Sort-Object displayName) | ForEach-Object { "$($_.displayName): $(@($_.members).Count)" }) '; ' } else { '' }) 'Privileged Users'

# ----------------------------------------------------------------------------------------------
# Remediation plan: every failed control plus every unimplemented Secure Score action, ranked by priority
# ----------------------------------------------------------------------------------------------
$PriorityRank = @{ 'Critical' = 1; 'High' = 2; 'Medium' = 3; 'Low' = 4 }
$script:remediationItems = New-Object System.Collections.Generic.List[object]
function Add-Remediation {
    param([string]$Priority, [string]$Finding, [string]$State, [string]$Action, $Points, [string]$Impact, [string]$Section, [string]$Url)
    $script:remediationItems.Add([pscustomobject]@{ Priority = $Priority; Rank = $PriorityRank[$Priority]; Finding = $Finding; State = $State; Action = $Action; Points = $Points; Impact = $Impact; Section = $Section; Url = $Url })
}
# Control -> (status pattern that counts as a finding, priority, recommended action)
$remediationGuides = @(
    @{ Control = 'MFA enforced for sign-in';                      Bad = '^(Not enforced|Partial)'; Priority = 'Critical'; Action = 'Create an enabled Conditional Access policy that requires MFA (or a phishing-resistant authentication strength) for all users and all cloud apps, excluding only a monitored break-glass account. Without Entra ID P1, turn on security defaults instead.' }
    @{ Control = 'MFA registered by all users';                   Bad = '^No$';                    Priority = 'High';     Action = 'Have every user listed under Users register Microsoft Authenticator (or a FIDO2 key). Use a registration campaign, then enforce MFA so unregistered accounts cannot sign in without completing it.' }
    @{ Control = 'Legacy authentication protocols disabled';      Bad = '^(Not blocked|Partial)';  Priority = 'High';     Action = 'Add a Conditional Access policy that blocks Exchange ActiveSync and "other clients" for all users. In Exchange Online disable SMTP AUTH tenant-wide and turn off POP and IMAP on every mailbox that does not need them.' }
    @{ Control = 'Automatic email forwarding outside the tenant'; Bad = '^(Active external|Allowed by policy)'; Priority = 'High'; Action = 'Set automatic forwarding to Off in the outbound spam filter policy, set the default remote domain to not allow auto-forward, then review and remove the mailbox forwards and inbox rules listed under Exchange Online.' }
    @{ Control = 'Conditional Access restricts access to Entra joined or compliant devices'; Bad = '^(Not enforced|Partial)'; Priority = 'Medium'; Action = 'Enroll company devices in Intune, then add a Conditional Access policy that requires a compliant or hybrid-joined device for Microsoft 365 apps. Start in report-only mode and review the sign-in impact before enforcing.' }
    @{ Control = 'Intune deployed to Windows computers';          Bad = '^(Not deployed|Partial)'; Priority = 'Medium';   Action = 'Enroll the Windows devices listed under Managed Devices into Intune (automatic enrollment through Entra join or group policy) so they can be measured against compliance policies.' }
    @{ Control = 'Intune compliance policies';                    Bad = '^None';                   Priority = 'Medium';   Action = 'Create a compliance policy per platform (Windows, iOS/iPadOS, Android, macOS) that requires encryption, a PIN or password, a minimum OS version and real-time protection, and marks devices noncompliant after a short grace period.' }
    @{ Control = 'Intune configuration and security policies';    Bad = '^0 policies';             Priority = 'Medium';   Action = 'Deploy a baseline: Microsoft security baseline, BitLocker, Defender Antivirus and firewall, screen lock, and Windows update rings. Assign to all devices and confirm the settings report as applied.' }
    @{ Control = 'User consent to applications';                  Bad = '^Allowed';                Priority = 'Medium';   Action = 'Set user consent to "Do not allow" (or limit it to verified publishers and low-impact permissions) and enable the admin consent request workflow with named reviewers, so every new application is approved by an administrator.' }
    @{ Control = 'Temporary Access Pass (TAP) implemented';       Bad = '^Disabled';               Priority = 'Low';      Action = 'Enable Temporary Access Pass for onboarding and MFA recovery: one-time use, a short lifetime, scoped to a help-desk-managed group.' }
    @{ Control = 'Users restricted from joining devices to Entra ID'; Bad = '^All users allowed';  Priority = 'Low';      Action = 'Restrict Entra ID device join to selected users or groups and require MFA to register or join a device.' }
    @{ Control = 'Email authentication (SPF, DKIM, DMARC)';      Bad = '^(Partial|Missing)';      Priority = 'High';     Action = 'For every email domain publish SPF ending in -all, enable DKIM signing in Exchange and publish both selector CNAMEs, and publish DMARC (start at p=none with reporting, then move to quarantine and reject). This stops criminals sending mail as your bank to customers and staff.' }
    @{ Control = 'Email protection (Defender for Office 365 / EOP)'; Bad = '^(Licensed, not configured|EOP only)'; Priority = 'High'; Action = 'If Defender for Office 365 is licensed, enable the Standard preset security policy for all users (Strict for executives and finance) so Safe Links, Safe Attachments and impersonation protection apply. If it is not licensed, license Plan 1 (included in Business Premium); phishing is the top attack path against bank staff.' }
    @{ Control = 'Endpoint protection (Defender for Endpoint)';  Bad = '^(Attention needed|No devices onboarded|Not licensed)'; Priority = 'Medium'; Action = 'Onboard every managed Windows and macOS device to Defender for Endpoint / Defender for Business through Intune, resolve high-risk alerts, and re-onboard or retire devices with inactive sensors.' }
)
foreach ($g in $remediationGuides) {
    $row = $script:summaryRows | Where-Object { $_.Control -eq $g.Control } | Select-Object -First 1
    if ($row -and ([string]$row.Status -match $g.Bad)) { Add-Remediation -Priority $g.Priority -Finding $g.Control -State "$($row.Status). $($row.Evidence)" -Action $g.Action -Section $row.Section }
}
# Licensing gaps that stop whole control families from existing
if ($script:HasIntune -eq $false)  { Add-Remediation -Priority 'Low' -Finding 'Intune is not licensed' -State 'No Intune service plan in the subscribed SKUs, so device compliance, configuration policies and device-based Conditional Access cannot be implemented.' -Action 'License Microsoft Intune (included in Microsoft 365 Business Premium, E3 and E5) and then complete the device controls above.' -Section 'Tenant Overview' }
if ($script:HasEntraP1 -eq $false) { Add-Remediation -Priority 'Low' -Finding 'Entra ID P1 is not licensed' -State 'No Entra ID P1/P2 service plan, so Conditional Access, the MFA registration report and sign-in activity are unavailable; protection relies on security defaults.' -Action 'License Entra ID P1 (included in Microsoft 365 Business Premium and E3) to gain Conditional Access and the MFA registration and sign-in reports this report measures.' -Section 'Tenant Overview' }
# Privileged access
$gaRole  = $roles | Where-Object { $_.displayName -eq 'Global Administrator' } | Select-Object -First 1
$gaCount = @($gaRole.members).Count
$gaNoMfa = @($roleRows | Where-Object { $_.Role -eq 'Global Administrator' -and $_.'MFA registered' -eq 'No' })
if ($gaNoMfa.Count) { Add-Remediation -Priority 'Critical' -Finding 'Global Administrators without MFA registered' -State ("$($gaNoMfa.Count) account(s): " + (Join-List ($gaNoMfa | ForEach-Object { $_.Member }))) -Action 'Register Microsoft Authenticator or a FIDO2 key for these accounts today and require MFA for all administrator roles through Conditional Access.' -Section 'Privileged Users' }
if ($gaCount -gt 4)  { Add-Remediation -Priority 'Medium' -Finding 'Too many Global Administrators' -State "$gaCount accounts hold Global Administrator (Microsoft recommends fewer than five)." -Action 'Reduce Global Administrators to two to four, move day-to-day administration to least-privilege roles (Exchange, Intune, User Administrator), and use PIM eligible assignments where Entra ID P2 is licensed.' -Section 'Privileged Users' }
if ($gaCount -eq 1)  { Add-Remediation -Priority 'Medium' -Finding 'Single Global Administrator' -State 'Only one account holds Global Administrator; losing access to it locks the organization out of the tenant.' -Action 'Create a second, cloud-only break-glass Global Administrator with a long random password stored offline, excluded from Conditional Access, and alerted on every sign-in.' -Section 'Privileged Users' }
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
.count{font-size:10pt;color:#555;font-weight:normal}
.empty{color:#555;font-style:italic}.note{background:#fff8e1;border:1px solid #f0d47a;padding:10px 12px;border-radius:4px;font-size:10pt}
.meta{font-size:9.5pt;color:#555;margin-bottom:18px}
ul.kvlist{margin:0;padding-left:18px}ul.kvlist li{margin:2px 0}
.filters{display:flex;flex-wrap:wrap;gap:6px;align-items:center;margin:10px 0 4px}
.filters button{font:inherit;font-size:9.5pt;border:1px solid var(--blue);background:#fff;color:var(--blue);border-radius:14px;padding:3px 11px;cursor:pointer}
.filters button.on{background:var(--blue);color:#fff}.filters input{font:inherit;font-size:9.5pt;padding:4px 8px;border:1px solid var(--line);border-radius:4px;min-width:220px;margin-left:auto}
.filters .n{font-size:9pt;color:#555}
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
.privacy{background:#eef6ee;border:1px solid #bfe0c2;border-left:5px solid var(--good);padding:10px 14px;border-radius:4px;font-size:10pt;margin:16px 0}
@media print{nav{display:none}main{margin:0;padding:0}section{page-break-before:always}section:first-of-type{page-break-before:avoid}table.grid thead th{position:static}.filters{display:none}details.info .tip{position:static;display:block;width:auto;box-shadow:none;margin-top:4px}details.info summary{display:none}details.info:not([open]) .tip{display:none}}
"@

$sections = @(
    @{ Id = 'overview';     Title = '1. Executive Summary' },
    @{ Id = 'summary';      Title = '2. Security Controls Summary' },
    @{ Id = 'remediation';  Title = '3. Remediation Plan' },
    @{ Id = 'mdm';          Title = '4. MDM Compliance Policies'; Sub = @($policyGroups | ForEach-Object { @{ Id = ('mdm-' + ($_.Name -replace '[^A-Za-z0-9]', '')); Title = $_.Name } }) },
    @{ Id = 'profiles';     Title = '5. Intune Policies' },
    @{ Id = 'devices';      Title = '6. Managed Devices' },
    @{ Id = 'ca';           Title = '7. Conditional Access' },
    @{ Id = 'authmethods';  Title = '8. Authentication Methods' },
    @{ Id = 'apps';         Title = '9. Enterprise Applications' },
    @{ Id = 'consent';      Title = '10. Application Consent Audit Events' },
    @{ Id = 'users';        Title = '11. Users' },
    @{ Id = 'roles';        Title = '12. Privileged Users' },
    @{ Id = 'exchange';     Title = '13. Exchange Online' },
    @{ Id = 'defender';     Title = '14. Defender and Email Protection' }
)
$SectionAnchors = @{ 'Authentication Methods' = 'authmethods'; 'Conditional Access' = 'ca'; 'Users' = 'users'; 'Managed Devices' = 'devices'; 'MDM Compliance Policies' = 'mdm'; 'Intune Policies' = 'profiles'; 'Exchange Online' = 'exchange'; 'Enterprise Applications' = 'apps'; 'Privileged Users' = 'roles' }
$SectionAnchors['Tenant Overview'] = 'overview'; $SectionAnchors['Executive Summary'] = 'overview'; $SectionAnchors['Remediation Plan'] = 'remediation'; $SectionAnchors['Secure Score'] = 'overview'; $SectionAnchors['Privileged Users'] = 'roles'; $SectionAnchors['Exchange Online'] = 'exchange'; $SectionAnchors['Defender'] = 'defender'; $SectionAnchors['Enterprise Applications'] = 'apps'
# ---- Directory synchronization (Entra Connect) for the overview
$syncEnabled = [bool]$org.onPremisesSyncEnabled
$syncLast    = $org.onPremisesLastSyncDateTime
$adDomains   = @($users | ForEach-Object { $_.onPremisesDomainName } | Where-Object { $_ } | Sort-Object -Unique)
$syncServers = @($users | Where-Object { $_.userPrincipalName -match '^Sync_(.+)_[0-9a-fA-F]{8,}@' } | ForEach-Object { $Matches[1] } | Sort-Object -Unique)
$syncedUserCount = @($users | Where-Object { $_.onPremisesSyncEnabled }).Count
$tenantKv['Directory synchronization (Entra Connect)'] = $(if ($syncEnabled) { "Enabled. Last sync $(Format-DateText $syncLast). $syncedUserCount synced user account(s)." } else { 'Not configured: all accounts are cloud-only.' })
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
    'MFA enforced for sign-in' = 'mfa-not-enforced'; 'MFA registered by all users' = 'mfa-not-registered'; 'Legacy authentication protocols disabled' = 'legacy-auth'
    'Conditional Access restricts access to Entra joined or compliant devices' = 'device-ca'; 'Intune deployed to Windows computers' = 'intune-coverage'
    'Intune compliance policies' = 'compliance-none'; 'Intune configuration and security policies' = 'config-none'; 'Temporary Access Pass (TAP) implemented' = 'tap-disabled'
    'Users restricted from joining devices to Entra ID' = 'device-join-open'; 'Automatic email forwarding outside the tenant' = 'forwarding'; 'User consent to applications' = 'consent-allowed'
    'Microsoft Secure Score' = 'secure-score'; 'Security defaults' = 'security-defaults-off'; 'Email authentication (SPF, DKIM, DMARC)' = 'dmarc-missing'; 'Email protection (Defender for Office 365 / EOP)' = 'defender-not-licensed'; 'Endpoint protection (Defender for Endpoint)' = 'mde-not-licensed'
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
foreach ($r in $appRows) {
    $is = @{}
    if ($r['Publisher'] -match 'unverified|Unknown') { $is['Publisher'] = 'app-unverified' }
    if ($r['User-consented delegated permissions'] -ne 'None') { $is['User-consented delegated permissions'] = 'app-user-consent' }
    if ($r['Application permissions'] -match $HighPrivPerm) { $is['Application permissions'] = 'app-high-privilege' }
    elseif ($r['Admin-consented delegated permissions'] -match $HighPrivPerm) { $is['Admin-consented delegated permissions'] = 'app-high-privilege' }
    if ($is.Count) { $r['_issues'] = $is }
}
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
$controlsMet   = @($controlsRated | Where-Object { [string]$_.Status -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met)' })
$findCritHigh  = @($remediationSorted | Where-Object { $_.Priority -in @('Critical', 'High') }).Count
$findMedium    = @($remediationSorted | Where-Object { $_.Priority -eq 'Medium' }).Count
$enabledMemberCount = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' }).Count
$noMfaCount    = if (@($regDetails).Count) { @($userRows | Where-Object { " $($_['_tags']) " -like '* nomfa *' }).Count } else { $null }
$staleUserCount = if ($signInAvailable) { @($userRows | Where-Object { " $($_['_tags']) " -like '* stale *' }).Count } else { $null }
$notInIntuneCount = @($allDeviceRows | Where-Object { " $($_['_tags']) " -like '* notintune *' -and $_['Enabled'] -ne $false }).Count
$kpiHtml = New-Object System.Text.StringBuilder
function Add-Kpi { param($Value, [string]$Label, [string]$Class = '') [void]$script:kpiHtml.Append("<div class='kpi $Class'><div class='v'>$(HtmlEnc $Value)</div><div class='l'>$(HtmlEnc $Label)</div></div>") }
if ($null -ne $scoreCurrent) { Add-Kpi "$scorePct%" "Secure Score: $scoreCurrent of $scoreMax points (target $scoreTarget)" $(if ($scoreCurrent -ge $scoreTarget) { 'good' } else { 'bad' }) } else { Add-Kpi 'n/a' 'Secure Score not available' 'mid' }
Add-Kpi "$($controlsMet.Count) of $($controlsRated.Count)" 'Security controls met' $(if ($controlsMet.Count -eq $controlsRated.Count) { 'good' } elseif ($controlsMet.Count * 2 -ge $controlsRated.Count) { 'mid' } else { 'bad' })
Add-Kpi $findCritHigh 'Critical and High findings' $(if ($findCritHigh -eq 0) { 'good' } else { 'bad' })
Add-Kpi $findMedium 'Medium findings' $(if ($findMedium -eq 0) { 'good' } else { 'mid' })
Add-Kpi $enabledMemberCount 'Enabled user accounts' ''
Add-Kpi $(if ($null -ne $noMfaCount) { $noMfaCount } else { 'n/a' }) 'Users without MFA registered' $(if ($null -eq $noMfaCount) { 'mid' } elseif ($noMfaCount -eq 0) { 'good' } else { 'bad' })
Add-Kpi $(if ($null -ne $staleUserCount) { $staleUserCount } else { 'n/a' }) "Stale users (no sign-in in $StaleDays days)" $(if ($null -eq $staleUserCount) { 'mid' } elseif ($staleUserCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi $gaCount 'Global Administrators' $(if ($gaCount -ge 2 -and $gaCount -le 4) { 'good' } else { 'mid' })
Add-Kpi $allDeviceRows.Count 'Devices (Entra ID and Intune)' ''
Add-Kpi $notInIntuneCount 'Enabled devices not in Intune' $(if ($notInIntuneCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi $staleDeviceCount "Stale devices ($StaleDays days)" $(if ($staleDeviceCount -eq 0) { 'good' } else { 'mid' })
$execNarrative = "This report was produced on $($ReportDate.ToString('MMMM d, yyyy')) from the live configuration of the $(HtmlEnc $ClientLabel) Microsoft 365 tenant, read-only, by $(HtmlEnc $ctx.Account). " +
    "It measures $($controlsRated.Count) security controls that regulators and examiners ask about; <b>$($controlsMet.Count) are met</b>. " +
    $(if ($null -ne $scoreCurrent) { "Microsoft Secure Score is <b>$scoreCurrent of $scoreMax points ($scorePct%)</b> against a target of $scoreTarget points ($TargetScorePercent%)$(if ($scoreCompareText) { "; $scoreCompareText" }). " } else { 'Microsoft Secure Score could not be read. ' }) +
    "The remediation plan contains <b>$($remediationSorted.Count) findings</b> ($remediationCounts). " +
    $(if ($enabledMemberCount) { "There are $enabledMemberCount enabled user accounts$(if ($null -ne $noMfaCount) { ", $noMfaCount without MFA registered" })$(if ($null -ne $staleUserCount) { ", $staleUserCount with no sign-in in $StaleDays days" }), $gaCount Global Administrator account(s), and $($allDeviceRows.Count) devices, of which $notInIntuneCount enabled device(s) are not managed by Intune." } else { '' })

$remHtml = New-Object System.Text.StringBuilder
[void]$remHtml.Append("<table class='grid'><thead><tr><th style='width:3%'>#</th><th style='width:7%'>Priority</th><th style='width:17%'>Finding</th><th style='width:20%'>Current state</th><th>Recommended action</th><th style='width:7%'>Secure Score points</th><th style='width:12%'>Impact / effort</th><th style='width:8%'>Detail</th></tr></thead><tbody>")
$remIdx = 0
foreach ($item in $remediationSorted) {
    $remIdx++
    $cls = switch ($item.Priority) { 'Critical' { 'warn' } 'High' { 'warn' } 'Medium' { 'mid' } default { '' } }
    $detail = if ($item.Url) { "<a href='$(HtmlEnc $item.Url)' target='_blank' rel='noopener'>Open in Defender</a>" } elseif ($SectionAnchors[$item.Section]) { "<a href='#$($SectionAnchors[$item.Section])'>$(HtmlEnc $item.Section)</a>" } else { HtmlEnc $item.Section }
    [void]$remHtml.Append("<tr><td>$remIdx</td><td class='$cls'>$(HtmlEnc $item.Priority)</td><td><b>$(HtmlEnc $item.Finding)</b></td><td>$(HtmlEnc $item.State)</td><td>$(HtmlEnc $item.Action)</td><td>$(if ($null -ne $item.Points) { HtmlEnc $item.Points })</td><td>$(HtmlEnc $item.Impact)</td><td>$detail</td></tr>")
}
[void]$remHtml.Append('</tbody></table>')
if ($remIdx -eq 0) { # ---- Directory synchronization (Entra Connect) for the overview
$syncEnabled = [bool]$org.onPremisesSyncEnabled
$syncLast    = $org.onPremisesLastSyncDateTime
$adDomains   = @($users | ForEach-Object { $_.onPremisesDomainName } | Where-Object { $_ } | Sort-Object -Unique)
$syncServers = @($users | Where-Object { $_.userPrincipalName -match '^Sync_(.+)_[0-9a-fA-F]{8,}@' } | ForEach-Object { $Matches[1] } | Sort-Object -Unique)
$syncedUserCount = @($users | Where-Object { $_.onPremisesSyncEnabled }).Count
$tenantKv['Directory synchronization (Entra Connect)'] = $(if ($syncEnabled) { "Enabled. Last sync $(Format-DateText $syncLast). $syncedUserCount synced user account(s)." } else { 'Not configured: all accounts are cloud-only.' })
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
    'MFA enforced for sign-in' = 'mfa-not-enforced'; 'MFA registered by all users' = 'mfa-not-registered'; 'Legacy authentication protocols disabled' = 'legacy-auth'
    'Conditional Access restricts access to Entra joined or compliant devices' = 'device-ca'; 'Intune deployed to Windows computers' = 'intune-coverage'
    'Intune compliance policies' = 'compliance-none'; 'Intune configuration and security policies' = 'config-none'; 'Temporary Access Pass (TAP) implemented' = 'tap-disabled'
    'Users restricted from joining devices to Entra ID' = 'device-join-open'; 'Automatic email forwarding outside the tenant' = 'forwarding'; 'User consent to applications' = 'consent-allowed'
    'Microsoft Secure Score' = 'secure-score'; 'Security defaults' = 'security-defaults-off'; 'Email authentication (SPF, DKIM, DMARC)' = 'dmarc-missing'; 'Email protection (Defender for Office 365 / EOP)' = 'defender-not-licensed'; 'Endpoint protection (Defender for Endpoint)' = 'mde-not-licensed'
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
foreach ($r in $appRows) {
    $is = @{}
    if ($r['Publisher'] -match 'unverified|Unknown') { $is['Publisher'] = 'app-unverified' }
    if ($r['User-consented delegated permissions'] -ne 'None') { $is['User-consented delegated permissions'] = 'app-user-consent' }
    if ($r['Application permissions'] -match $HighPrivPerm) { $is['Application permissions'] = 'app-high-privilege' }
    elseif ($r['Admin-consented delegated permissions'] -match $HighPrivPerm) { $is['Admin-consented delegated permissions'] = 'app-high-privilege' }
    if ($is.Count) { $r['_issues'] = $is }
}
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
$controlsMet   = @($controlsRated | Where-Object { [string]$_.Status -match '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met)' })
$findCritHigh  = @($remediationSorted | Where-Object { $_.Priority -in @('Critical', 'High') }).Count
$findMedium    = @($remediationSorted | Where-Object { $_.Priority -eq 'Medium' }).Count
$enabledMemberCount = @($users | Where-Object { $_.accountEnabled -and $_.userType -eq 'Member' }).Count
$noMfaCount    = if (@($regDetails).Count) { @($userRows | Where-Object { " $($_['_tags']) " -like '* nomfa *' }).Count } else { $null }
$staleUserCount = if ($signInAvailable) { @($userRows | Where-Object { " $($_['_tags']) " -like '* stale *' }).Count } else { $null }
$notInIntuneCount = @($allDeviceRows | Where-Object { " $($_['_tags']) " -like '* notintune *' -and $_['Enabled'] -ne $false }).Count
$kpiHtml = New-Object System.Text.StringBuilder
function Add-Kpi { param($Value, [string]$Label, [string]$Class = '') [void]$script:kpiHtml.Append("<div class='kpi $Class'><div class='v'>$(HtmlEnc $Value)</div><div class='l'>$(HtmlEnc $Label)</div></div>") }
if ($null -ne $scoreCurrent) { Add-Kpi "$scorePct%" "Secure Score: $scoreCurrent of $scoreMax points (target $scoreTarget)" $(if ($scoreCurrent -ge $scoreTarget) { 'good' } else { 'bad' }) } else { Add-Kpi 'n/a' 'Secure Score not available' 'mid' }
Add-Kpi "$($controlsMet.Count) of $($controlsRated.Count)" 'Security controls met' $(if ($controlsMet.Count -eq $controlsRated.Count) { 'good' } elseif ($controlsMet.Count * 2 -ge $controlsRated.Count) { 'mid' } else { 'bad' })
Add-Kpi $findCritHigh 'Critical and High findings' $(if ($findCritHigh -eq 0) { 'good' } else { 'bad' })
Add-Kpi $findMedium 'Medium findings' $(if ($findMedium -eq 0) { 'good' } else { 'mid' })
Add-Kpi $enabledMemberCount 'Enabled user accounts' ''
Add-Kpi $(if ($null -ne $noMfaCount) { $noMfaCount } else { 'n/a' }) 'Users without MFA registered' $(if ($null -eq $noMfaCount) { 'mid' } elseif ($noMfaCount -eq 0) { 'good' } else { 'bad' })
Add-Kpi $(if ($null -ne $staleUserCount) { $staleUserCount } else { 'n/a' }) "Stale users (no sign-in in $StaleDays days)" $(if ($null -eq $staleUserCount) { 'mid' } elseif ($staleUserCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi $gaCount 'Global Administrators' $(if ($gaCount -ge 2 -and $gaCount -le 4) { 'good' } else { 'mid' })
Add-Kpi $allDeviceRows.Count 'Devices (Entra ID and Intune)' ''
Add-Kpi $notInIntuneCount 'Enabled devices not in Intune' $(if ($notInIntuneCount -eq 0) { 'good' } else { 'mid' })
Add-Kpi $staleDeviceCount "Stale devices ($StaleDays days)" $(if ($staleDeviceCount -eq 0) { 'good' } else { 'mid' })
$execNarrative = "This report was produced on $($ReportDate.ToString('MMMM d, yyyy')) from the live configuration of the $(HtmlEnc $ClientLabel) Microsoft 365 tenant, read-only, by $(HtmlEnc $ctx.Account). " +
    "It measures $($controlsRated.Count) security controls that regulators and examiners ask about; <b>$($controlsMet.Count) are met</b>. " +
    $(if ($null -ne $scoreCurrent) { "Microsoft Secure Score is <b>$scoreCurrent of $scoreMax points ($scorePct%)</b> against a target of $scoreTarget points ($TargetScorePercent%)$(if ($scoreCompareText) { "; $scoreCompareText" }). " } else { 'Microsoft Secure Score could not be read. ' }) +
    "The remediation plan contains <b>$($remediationSorted.Count) findings</b> ($remediationCounts). " +
    $(if ($enabledMemberCount) { "There are $enabledMemberCount enabled user accounts$(if ($null -ne $noMfaCount) { ", $noMfaCount without MFA registered" })$(if ($null -ne $staleUserCount) { ", $staleUserCount with no sign-in in $StaleDays days" }), $gaCount Global Administrator account(s), and $($allDeviceRows.Count) devices, of which $notInIntuneCount enabled device(s) are not managed by Intune." } else { '' })

$remHtml = New-Object System.Text.StringBuilder; [void]$remHtml.Append("<p class='empty'>No findings. Every measured control passed and every Secure Score action is implemented.</p>") }
$summaryTable = New-Object System.Text.StringBuilder
[void]$summaryTable.Append("<table class='grid'><thead><tr><th style='width:26%'>Control</th><th style='width:14%'>Status</th><th>Evidence</th><th style='width:12%'>Detail</th></tr></thead><tbody>")
foreach ($row in $script:summaryRows) {
    $cls = switch -Regex ([string]$row.Status) {
        '^(Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded)' { 'good'; break }
        '^(Partial|Allowed|All users allowed|Not available|Not collected|No Windows|EOP only|No email domains|No devices onboarded|Not licensed)' { 'mid'; break }
        '^(Disabled|Not enforced|Not blocked|No$|Not deployed|None|Active external|Allowed by policy|Below target|Missing|Licensed, not configured|Attention needed)' { 'warn'; break }
        default { '' }
    }
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
            if ($row.Control -like 'Endpoint protection*' -and $st -eq 'No devices onboarded') { $iid = 'mde-not-onboarded' }
            $sumIcon = New-InfoIcon $iid
        }
    }
    [void]$summaryTable.Append("<tr><td><b>$(HtmlEnc $row.Control)</b></td><td class='$cls'>$(HtmlEnc $row.Status)$sumIcon</td><td>$(HtmlEnc $row.Evidence)</td><td><a href='#$anchor'>$(HtmlEnc $row.Section)</a></td></tr>")
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
[void]$navHtml.Append("$(if ($logoImg) { "<a href='https://audithat.com' target='_blank' rel='noopener'>$logoImg</a>" })<h1>$(HtmlEnc $ClientName)</h1><div class='sub'>$(if ($DefaultDomain) { "$(HtmlEnc $DefaultDomain)<br>" })Microsoft 365 Audit Evidence<br>$($ReportDate.ToString('MMMM d, yyyy'))<br>Provided by $brandLink<br>Script version $ScriptVersion &middot; <a href='$RepoUrl' target='_blank' rel='noopener'>GitHub</a></div>")
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
<section id='mdm'><h2>4. MDM Compliance Policies</h2>
<p class='meta'>Source: Microsoft Graph deviceManagement/deviceCompliancePolicies, exported $($ReportDate.ToString('yyyy-MM-dd HH:mm')). Values are read directly from the tenant configuration.</p>
$complianceHtmlText
</section>
"@

$caSection = @"
<section id='ca'><h2>7. Conditional Access</h2>
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

$html = @"
<!DOCTYPE html>
<html lang='en'><head><meta charset='utf-8'><title>$(HtmlEnc $ClientLabel) - Microsoft 365 Audit Evidence</title><style>$css</style></head>
<body>
<nav>$($navHtml.ToString())</nav>
<main>
<section id='overview'><h2>1. Executive Summary</h2>
<p class='exec'>$execNarrative</p>
<p class='legend'><span class='i'>i</span> marks a configuration issue anywhere in this report. Click the icon to read why it matters and what to do about it.</p>
<div class='kpis'>$($kpiHtml.ToString())</div>
<h3>Top priorities</h3>
<div class='exec'>$(if ($remediationSorted.Count) { '<ol>' + ((@($remediationSorted | Select-Object -First 5) | ForEach-Object { "<li><b>$(HtmlEnc $_.Priority):</b> $(HtmlEnc $_.Finding)$(if ($_.Points) { " ($($_.Points) Secure Score points)" })</li>" }) -join '') + '</ol>' } else { "<p class='empty'>No open findings.</p>" }) <p class='meta'>Full list with recommended actions in <a href='#remediation'>section 3, Remediation Plan</a>.</p></div>
<div class='privacy'><b>Privacy.</b> This report was generated locally by the free Audit Hat M365 Audit Evidence script ($brandLink) on the computer of the signed-in user. The script reads the tenant with that user's own permissions and writes only to the local output folder. No data is sent to Audit Hat or any third party.</div>
<h3>Tenant details</h3>
$(New-KvTable $tenantKv -Issues $tenantIssues)
<h3>Verified domains</h3>
$(New-HtmlTable $domainRows -Columns 'Domain','Default','Initial (.onmicrosoft.com)','Type','Capabilities' -EmptyText 'No verified domains returned.')
<h3>Subscriptions</h3>
$(New-HtmlTable $skuRows -Columns 'License SKU','Enabled','Consumed','Status')
</section>

<section id='summary'><h2>2. Security Controls Summary</h2>
<p class='meta'>Each status is computed from the tenant configuration collected below. Status reflects what is enforced today, not what is configured in report-only or disabled policies.</p>
$($summaryTable.ToString())
</section>

<section id='remediation'><h2>3. Remediation Plan</h2>
<p class='meta'>Every control from the summary that is not fully met, plus every Microsoft Secure Score improvement action that still has points available, ranked by priority (Critical, High, Medium, Low); within a priority the measured controls come first, then Secure Score actions by points available. Findings: $(HtmlEnc $remediationCounts). $(if ($null -ne $scoreCurrent) { "Secure Score $scoreCurrent of $scoreMax points; target $scoreTarget points ($TargetScorePercent%); gap $scoreGap points." } else { 'Secure Score not available.' })$(if ($scoreIgnored) { " $scoreIgnored Secure Score action(s) marked ignored, third-party or resolved in Defender are not listed." })</p>
$($remHtml.ToString())
</section>
$mdmSection

<section id='profiles'><h2>5. Intune Policies</h2>
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

<section id='devices'><h2>6. Managed Devices</h2>
$(New-KvTable $deviceStats -Issues $deviceStatIssues)
<h3>Device list</h3>
<p class='meta'>Every device object in Entra ID joined with its Intune record, plus Intune devices that have no Entra object. "Stale" means no sign-in or Intune sync in $StaleDays days. "Last user" is the most recent account Intune saw log on; "Owner / primary user" is the registered owner.</p>
$(New-HtmlTable $allDeviceRows -Id 'tbl-devices' -Filters $DeviceFilters -EmptyText 'No devices were returned from Entra ID or Intune.')
</section>

$caSection

<section id='authmethods'><h2>8. Authentication Methods</h2>
<h3>Security defaults</h3>
$(New-KvTable ([ordered]@{ 'Security defaults' = $(if ($null -ne $secDefaults.isEnabled) { $(if ($secDefaults.isEnabled) { 'Enabled' } else { 'Disabled' }) } else { 'Could not read' }) }) -Issues $secDefIssues)
<h3>Temporary Access Pass</h3>
$(New-KvTable $tapKv -Issues $tapIssues)
<h3>Authentication method policy</h3>
$(New-HtmlTable $authMethodRows -Columns 'Method','State','Included targets','Excluded targets','Settings' -EmptyText 'Could not read the authentication methods policy.')
<h3>Device registration settings</h3>
$(New-KvTable $devRegKv -Issues $devRegIssues)
</section>

<section id='apps'><h2>9. Enterprise Applications</h2>
<h3>Application governance settings</h3>
$(New-KvTable $appGovKv -Issues $appGovIssues)
<p class='note'>Approver and approval date come from the Entra directory audit log, which retains $AuditDays days in this export (30 days with Entra ID P1/P2, 7 days on Free). Applications consented before that window show the date they were added to the tenant and "Not in audit log retention" for the approver.</p>
<h3>Applications and consented permissions</h3>
$(New-HtmlTable $appRows -EmptyText 'No enterprise applications returned.')
</section>

<section id='consent'><h2>10. Application Consent Audit Events (last $AuditDays days)</h2>
$(New-HtmlTable $auditRows -Columns 'Date (local)','Activity','Performed by','Application','Admin consent','Result' -EmptyText "No application consent or add events were recorded in the last $AuditDays days.")
</section>

<section id='users'><h2>11. Users</h2>
$(New-KvTable $userStats -Issues $userStatIssues)
<h3>Enabled member accounts without MFA registered</h3>
$(if (@($regDetails).Count) { New-HtmlTable $noMfaRows -Columns 'User','UPN','Licensed','Last sign-in','Methods registered' -EmptyText 'Every enabled member account has an MFA method registered.' } else { "<p class='empty'>MFA registration report not available (requires Entra ID P1).</p>" })
<h3>User list</h3>
$(New-HtmlTable $userRows -Id 'tbl-users' -Filters $UserFilters -EmptyText 'No users returned.')
</section>

<section id='roles'><h2>12. Privileged Users</h2>
<p class='meta'>Active directory role assignments, grouped by role. Roles with no members are omitted. Groups assigned to a role are shown as a group row; their members are not expanded.</p>
$(New-KvTable $privStats)
$($privHtml.ToString())
$(if (@($activeRoles).Count -eq 0) { "<p class='empty'>No directory roles have active members.</p>" })
</section>

<section id='exchange'><h2>13. Exchange Online</h2>
<p class='meta'>Automatic forwarding and legacy protocol settings from Exchange Online PowerShell.</p>
$(New-KvTable $exoKv -Issues $exoIssues)
$exoDetailHtml
<h3>Email authentication: SPF, DKIM, DMARC</h3>
<p class='meta'>Public DNS records for every email-enabled custom domain, checked from this computer at report time, plus the DKIM signing state from Exchange Online. SPF and DKIM prove mail really came from your tenant; DMARC tells the world's mail servers to reject mail that fails and sends you reports of spoofing attempts.</p>
$(if (-not $dnsAvailable) { "<p class='empty'>DNS lookups are not available on this computer (Resolve-DnsName missing).</p>" } else { New-HtmlTable $emailAuthRows -Columns 'Domain','MX','SPF','SPF record','DKIM signing (Exchange)','DKIM DNS (selector1 / selector2)','DMARC policy','DMARC coverage','DMARC reporting (rua)','DMARC record' -EmptyText 'No email-enabled custom domains are verified in this tenant.' })
</section>

<section id='defender'><h2>14. Defender and Email Protection</h2>
<p class='meta'>Exchange Online Protection and Microsoft Defender for Office 365 policies (anti-spam, anti-malware, anti-phishing, Safe Links, Safe Attachments, preset security policies) and Microsoft Defender for Endpoint device status.</p>
$(New-KvTable ([ordered]@{
    'Defender for Office 365 licence'      = $(if ($script:HasDefenderP2) { 'Plan 2' } elseif ($script:HasDefenderP1) { 'Plan 1' } else { 'Not licensed: Exchange Online Protection only' })
    'Defender for Endpoint licence'        = $(if ($script:HasMDE) { 'Licensed' } else { 'Not licensed' })
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
<h3>Defender for Endpoint devices</h3>
$(if ($mdeCollected) { (New-KvTable $mdeKv) + (New-HtmlTable $mdeRows -Id 'tbl-mde' -Filters $MdeFilters -EmptyText 'No devices are onboarded to Defender for Endpoint.') } else { "<p class='empty'>$(HtmlEnc $mdeEvidence)</p>" })
</section>
<footer class='brand'>$(if ($logoImg) { "<a href='https://audithat.com' target='_blank' rel='noopener'>$logoImg</a>" })<div>Microsoft 365 Audit Evidence for $(HtmlEnc $ClientLabel), generated $($ReportDate.ToString('MMMM d, yyyy HH:mm')) by $(HtmlEnc $ctx.Account).<br>Produced locally by the free, read-only Audit Hat M365 Audit Evidence script, version $ScriptVersion (<a href='$RepoUrl' target='_blank' rel='noopener'>source and updates on GitHub</a>). Provided by $brandLink. No data left your organization.</div></footer>
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
  // one info panel open at a time; click anywhere else closes it
  document.addEventListener('click', function(e){
    var open=document.querySelectorAll('details.info[open]');
    open.forEach(function(d){ if(!d.contains(e.target)) d.removeAttribute('open'); });
  });
})();
</script>
</main></body></html>
"@

$safeClient = ($ClientLabel -replace '[^A-Za-z0-9 ._()-]', '').Trim()
if (-not $safeClient) { $safeClient = 'Client' }
$reportPath = Join-Path $OutputFolder "$safeClient - M365 Audit Evidence $($ReportDate.ToString('yyyy-MM-dd')).html"
[System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
Write-Host "HTML report: $reportPath" -ForegroundColor Green

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



