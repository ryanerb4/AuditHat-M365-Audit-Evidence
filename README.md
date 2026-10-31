# AuditHat M365 Audit Evidence

**Version:** 2026.10.31 (date-based) · **Source:** https://github.com/ryanerb4/AuditHat-M365-Audit-Evidence

**Author:** Ryan Erb · **Provided by:** [Audit Hat](https://AuditHat.com)

One read-only PowerShell script that signs in to Microsoft 365, collects the Entra ID, Intune, and
Exchange Online evidence an IT auditor asks a community bank or credit union for, and writes a
single self-contained HTML report with a left navigation bar. Optionally also produces two Word
documents (`MDM Configuration.docx`, `Conditional Access Policies.docx`).

The goal is to help banks measure their Microsoft 365 security and compliance posture, understand
their risk, and hand auditors factual evidence without anyone taking portal screenshots.

## Privacy: your data never leaves your organization

**This script sends no data to Audit Hat or to any third party.** It is a free, purely internal tool
provided for your own use to help you understand and reduce your Microsoft 365 risk.

- It runs entirely on the computer where you start it.
- It reads your tenant with the account **you** sign in with, read-only, and writes the report only
  to the output folder on that computer (default `.\M365-Audit-Evidence`).
- Audit Hat has no server side for this tool and receives nothing: no telemetry, no copies, no
  "phone home". You can confirm this by reading the script; every network call goes to Microsoft
  (`login.microsoftonline.com`, `graph.microsoft.com`, `outlook.office365.com`,
  `powershellgallery.com` for module install).

**Why there is an "AuditHat" app in the sign-in.** Microsoft requires every sign-in to identify an
application. The script identifies itself as the **AuditHat Pulse Reports (PowerShell)** connector,
an Entra enterprise application published by Audit Hat (the "Microsoft 365 AuditHat PowerShell
connector"). Consenting to it in your tenant **grants this local script read-only rights for the
signed-in user and nothing else**:

- it holds no client secret or certificate, so nobody can use it without a person signing in with
  your tenant's own MFA;
- it has delegated permissions only, so it can never do more than the signed-in person can;
- it has no application (unattended) permissions and no server behind it.

If your policy requires a tenant-owned app instead, see "Register your own app" below and pass
`-ClientId`; the script behaves identically.

---

## Quick start

```powershell
# 1. Allow scripts for this PowerShell window only
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass

# 2. Run it (modules install themselves on first run; a browser window opens for sign-in).
#    The report is named after the tenant display name and default domain; -ClientName overrides that.
.\Get-M365AuditEvidence.ps1

# With the two Word documents (needs Microsoft Word on this PC)
.\Get-M365AuditEvidence.ps1 -ClientName "Sample Bank" -WordDocs

# Tenant blocks Microsoft's Graph PowerShell app (AADSTS700016)? Use your own app registration (see below)
.\Get-M365AuditEvidence.ps1 -ClientName "Sample Bank" -ClientId <app id> -TenantId <tenant id>

# Target a specific tenant and output folder, skip the slow inbox-rule scan
.\Get-M365AuditEvidence.ps1 -TenantId contoso.onmicrosoft.com -OutputFolder C:\Audit\Contoso -SkipInboxRules
```

Output lands in `.\M365-Audit-Evidence\<Client> - M365 Audit Evidence <date>.html` by default.

---

## Requirements

| # | Requirement | Notes |
|---|-------------|-------|
| 1 | Windows 10/11 or Windows Server | Windows PowerShell 5.1 (built in) or PowerShell 7.x |
| 2 | Internet access | graph.microsoft.com, login.microsoftonline.com, outlook.office365.com, powershellgallery.com |
| 3 | `Microsoft.Graph.Authentication` module | **Auto-installed** for the current user on first run |
| 4 | `ExchangeOnlineManagement` module | **Auto-installed** for the current user on first run; only needed for the Exchange section |
| 5 | A Microsoft 365 account | Sign-in opens the normal Microsoft sign-in page in your default browser; approve the Microsoft Authenticator / MFA prompt and you are in. **MFA, Conditional Access and passwordless all work.** `-DeviceCode` is available for remote or locked-down sessions |
| 6 | Microsoft Word (optional) | Only for `-WordDocs` |

No local admin rights are needed. Modules install with `-Scope CurrentUser`.

### Which account to sign in with

- **First run in a tenant:** a **Global Administrator** (or Privileged Role Administrator / Cloud
  Application Administrator). Microsoft Graph will show a one-time consent prompt for the read-only
  permissions below. Accept it.
- **Every run after that:** **Global Reader** is enough for all Graph sections. The Exchange Online
  section additionally needs **View-Only Organization Management** or **Exchange Administrator**.

If a permission was not granted, the script warns at sign-in and marks the affected section
"not available" instead of failing.

### Graph permissions requested (all read-only, delegated)

`Organization.Read.All`, `User.Read.All`, `Group.Read.All`, `Directory.Read.All`,
`Application.Read.All`, `Policy.Read.All`, `Policy.Read.ConditionalAccess`,
`Policy.Read.AuthenticationMethod`, `DeviceManagementConfiguration.Read.All`,
`DeviceManagementManagedDevices.Read.All`, `DeviceManagementApps.Read.All`, `Device.Read.All`,
`AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`, `RoleManagement.Read.Directory`, `SecurityEvents.Read.All` (Microsoft Secure Score)

The script never requests a write scope and only runs `Get-*` Exchange cmdlets. Nothing in the
tenant is changed.

### Register your own app (needed if sign-in fails with AADSTS700016)

By default the script signs in through Microsoft's own public app, **Microsoft Graph Command Line
Tools** (client ID `14d82eec-204b-4c2f-b7e8-296a70dae67e`), the same app `Connect-MgGraph` uses.
Some tenants do not have that app and block it from being added, and sign-in then fails with:

> AADSTS700016: Application with identifier '14d82eec-...' was not found in the directory

The fix is a small app registration owned by the tenant. It is a **public client with no secret
and no certificate**, so there is nothing to protect or rotate, and it only holds the read-only
delegated permissions listed above. Takes about five minutes as a Global Administrator:

1. **Entra admin center** (https://entra.microsoft.com) > Identity > Applications >
   **App registrations** > **New registration**.
   - Name: `M365 Audit Evidence` (any name works)
   - Supported account types: **Accounts in this organizational directory only**
   - Redirect URI: platform **Mobile and desktop applications**, value `http://localhost`
   - Register.
2. On the app's **Overview** page copy the **Application (client) ID** and the
   **Directory (tenant) ID**.
3. **Authentication** > scroll to **Advanced settings** > *Allow public client flows* = **Yes** >
   Save. (This enables the `-DeviceCode` fallback; the browser flow works either way.)
4. **API permissions** > **Add a permission** > **Microsoft Graph** > **Delegated permissions** >
   tick each of these, then **Add permissions**:
   `Organization.Read.All`, `User.Read.All`, `Group.Read.All`, `Directory.Read.All`,
   `Application.Read.All`, `Policy.Read.All`, `Policy.Read.ConditionalAccess`,
   `Policy.Read.AuthenticationMethod`, `DeviceManagementConfiguration.Read.All`,
   `DeviceManagementManagedDevices.Read.All`, `DeviceManagementApps.Read.All`, `Device.Read.All`,
   `AuditLog.Read.All`, `UserAuthenticationMethod.Read.All`, `RoleManagement.Read.Directory`, `SecurityEvents.Read.All` (Microsoft Secure Score)
5. Still on **API permissions**, click **Grant admin consent for \<tenant\>** and confirm. Every
   row should show a green check.
6. Run the script with both IDs:

   ```powershell
   .\Get-M365AuditEvidence.ps1 -ClientName "Sample Bank" -ClientId <Application (client) ID> -TenantId <Directory (tenant) ID>
   ```

   Or save them once in `M365AuditEvidence.settings.json` next to the script and run with no
   extra parameters:

   ```json
   { "ClientId": "11111111-2222-3333-4444-555555555555", "TenantId": "8505d674-44f6-4a98-aa7a-20ef5bba7225" }
   ```

   The settings file contains no secrets (public client), so it is safe to keep next to the script.

The Exchange Online section still signs in through Microsoft's own Exchange Online PowerShell app;
if that is blocked too, run with `-SkipExchange`.

---

## Info icons

Every value in the report that represents a configuration issue carries a red **i** icon. Clicking it
opens a panel with the issue title, *why it matters* to a bank or credit union, and the *recommended
fix*. The catalogue of about fifty issues lives in the script (`$IssueCatalog`) and covers MFA, legacy
authentication, device management, application consent, privileged roles, stale accounts and devices,
mail forwarding, POP/IMAP/SMTP AUTH, SPF/DKIM/DMARC, anti-spam, anti-phishing, Safe Links/Attachments,
Defender licensing and directory sync. Icons appear in the summary table, key/value tables and data
tables alike, so an executive can click through from any page without reading the detail sections.

## What the report contains

| # | Section | Source | What it answers |
|---|---------|--------|-----------------|
| 1 | Executive Summary | organization, subscribedSkus, security/secureScores, computed | Narrative paragraph and KPI tiles (Secure Score vs target, controls met, Critical/High and Medium findings, users without MFA, stale users, Global Administrators, devices, devices not in Intune, stale devices), top five priorities, privacy statement, tenant details with bulleted licensed capabilities, verified domains table, subscriptions |
| 2 | Security Controls Summary | computed from everything below | One row per control (MFA, legacy auth, device access, Intune, TAP, device join, forwarding, consent, **email authentication**, **email protection**, **endpoint protection**, Secure Score, privileged roles): Status, Evidence, link to detail. Only **enabled** CA policies count as enforced |
| 3 | Remediation Plan | summary results, secureScoreControlProfiles, directoryRoles, subscribedSkus | Every failed control with a recommended action, privileged-access findings (Global Admins without MFA, too many or a single Global Admin), licensing gaps, and every unimplemented Secure Score action with points available. Sorted Critical, High, Medium, Low, then by points |
| 4 | MDM Compliance Policies | deviceCompliancePolicies | Intune compliance policies per platform: Basics / Settings / Actions for noncompliance / Assignments |
| 5 | Intune Policies | deviceConfigurations, configurationPolicies (settings catalog), intents (endpoint security), enrollment configs, app protection | Every Intune policy and the settings behind it |
| 6 | Managed Devices | managedDevices, devices (+registeredOwners) | Device counts, then one device list joining every Entra ID device with its Intune record: owner, **last user** seen by Intune, last activity, compliance, encryption. Quick filters: Stale (90 days), Not in Intune, Noncompliant, Not encrypted, Personal, Windows, Disabled, plus text search |
| 7 | Conditional Access | conditionalAccess/policies, namedLocations, authenticationStrength | Full policy detail with IDs resolved to names; filtered tables for MFA, legacy-auth block, device-based policies |
| 8 | Authentication Methods | security defaults, authenticationMethodsPolicy, deviceRegistrationPolicy | Security defaults, Temporary Access Pass, each method's state, who may join devices |
| 9 | Enterprise Applications | servicePrincipals, oauth2PermissionGrants, appRoleAssignments, authorizationPolicy | Publisher, admin vs user consent, delegated and application permissions, date added, approver |
| 10 | Application Consent Audit Events | auditLogs/directoryAudits | Consent / add-app events in the last `-AuditDays` days |
| 11 | Users | users (signInActivity), userRegistrationDetails, directoryRoles | Counts, enabled members without MFA registered, full user list with last sign-in, licenses, MFA methods, Admin flag and **number of devices**. Quick filters: Stale (90 days, needs P1), No MFA, Admin, Guests, Disabled, Licensed, plus text search |
| 12 | Privileged Users | directoryRoles, roleEligibilityScheduleInstances | One table per role, Global Administrator first, MFA and last sign-in per member, PIM eligible list |
| 13 | Exchange Online | ExchangeOnlineManagement cmdlets, public DNS (Resolve-DnsName) | Auto-forwarding (remote domain, outbound spam policy, mailbox, inbox rules, transport rules), POP/IMAP/SMTP AUTH, modern auth, and **email authentication per custom domain: MX, SPF (and its all-mechanism), DKIM signing state plus selector CNAMEs, DMARC policy, coverage and reporting** |
| 14 | Defender and Email Protection | Get-HostedContentFilterPolicy, Get-MalwareFilterPolicy, Get-AntiPhishPolicy, Get-SafeLinksPolicy, Get-SafeAttachmentPolicy, Get-AtpPolicyForO365, Get-EOPProtectionPolicyRule, Get-ATPProtectionPolicyRule, Defender for Endpoint API | Licence status for Defender for Office 365 and Defender for Endpoint, preset security policies, anti-spam actions and bulk threshold, ZAP, anti-malware common attachments filter, anti-phishing impersonation settings and threshold, Safe Links, Safe Attachments, and Defender for Endpoint device list (sensor health, risk, exposure, last seen) with exposure and configuration scores |

### The auditor questions this answers

1. MDM compliance policies for Android and iOS/iPadOS (Word doc)
2. Conditional Access policies (Word doc)
3. Enterprise applications, permissions granted, when approved and by whom
4. User list with status, licenses, last sign-in, MFA registration
5. Is MFA enforced, and registered by every user
6. Are legacy authentication protocols disabled
7. Does Conditional Access restrict access to Entra joined or compliant devices only
8. Is Intune deployed, and on which computers
9. Every Intune policy with its settings
10. Is Temporary Access Pass implemented
11. Is automatic email forwarding outside the tenant allowed or in use
12. Privileged users by role

### A note on "who approved and when"

Microsoft Graph does not store an approver on a permission grant. The only record is the Entra
directory audit log, which keeps **30 days** with Entra ID P1/P2 and **7 days** on Free. Apps
consented earlier show the date they were added to the tenant and "Not in audit log retention"
for the approver. The report never guesses.

---

## Parameters

| Parameter | Default | Purpose |
|-----------|---------|---------|
| `-ClientName` | tenant display name | Optional override. By default the report is named after the tenant display name and default domain read from the tenant (for example `Sample Bank (samplebank.com)`), on the cover, nav, file name, and Word docs |
| `-TenantId` | current account's home tenant | Tenant GUID or `.onmicrosoft.com` name. Required together with `-ClientId` |
| `-ClientId` | Microsoft Graph Command Line Tools | Your own app registration's Application (client) ID. See "Register your own app". Can also come from `M365AuditEvidence.settings.json` |
| `-OutputFolder` | `.\M365-Audit-Evidence` | Where the report is written |
| `-AuditDays` | `30` | Directory audit log window |
| (none) | | **Version** is printed in the console banner, the report navigation and footer (`2026.10.31` style: year.month.day of the last change). The report links to this GitHub repository for source and updates |
| `-TargetScorePercent` | `80` | Target Microsoft Secure Score as a percentage of the maximum achievable score; drives the target, gap and the "Below target" status |
| `-WordDocs` | off | Also produce the two `.docx` files (needs Word) |
| `-SkipExchange` | off | Skip the Exchange Online section |
| `-SkipInboxRules` | off | Skip the per-mailbox inbox rule scan (slow on large tenants) |
| `-DeviceCode` | off | Sign in with a device code (type a code at https://microsoft.com/devicelogin) instead of the browser. Use it in remote sessions or if the browser flow cannot open a local port. The script also falls back to this automatically if the browser flow fails |

---

## How the script works

1. **Prerequisites.** Checks for the two modules and installs them for the current user if missing
   (enables TLS 1.2 and the NuGet provider on Windows PowerShell 5.1 when needed).
2. **Sign-in.** The script runs the standard OAuth 2.0 authorization-code flow with PKCE itself:
   it opens `https://login.microsoftonline.com/organizations/...` in the default browser (work
   accounts only, so a personal Microsoft account signed in to Windows is never picked up), you sign
   in with the admin account and approve the Authenticator / MFA prompt, and the browser returns to a
   loopback listener (`http://localhost:<random port>/`) on the PC. The token is handed to
   `Connect-MgGraph -AccessToken` and refreshed automatically during long runs. It uses the same
   public Microsoft app ("Microsoft Graph Command Line Tools") that `Connect-MgGraph` uses, with only
   the read-only scopes listed above. If the browser flow fails, or `-DeviceCode` is given, the
   script falls back to `Connect-MgGraph -UseDeviceCode`. After connecting it compares granted
   scopes to the requested list and warns about any missing ones. Why not plain `Connect-MgGraph`?
   See the WAM row under Troubleshooting.
3. **Collection.** Every Graph call goes through `Get-GraphAllSafe`, which follows `@odata.nextLink`
   paging and turns any failure into a warning plus an empty result, so one missing license or
   permission never stops the run. Object IDs (groups, users, roles, apps) are resolved to display
   names in batches through `directoryObjects/getByIds` and cached.
4. **Exchange Online.** Connects with the same account used for Graph (interactive, MFA-capable),
   runs only `Get-*` cmdlets, and always disconnects.
5. **Evaluation.** The Security Controls Summary is computed from the collected data. A control is
   only "Enforced" or "Blocked" when an **enabled** Conditional Access policy targets **all users
   and all cloud apps**; report-only or scoped policies show "Partial".
6. **Rendering.** All HTML is built from strings with every value HTML-encoded. The CSS is inline, so
   the report is a single file that can be emailed or attached to an audit workbook.
7. **Word export** (optional). The MDM and Conditional Access sections are written as HTML and
   converted to `.docx` through Word COM automation.

### Design constraints

- Read-only. No write scopes, no changes to the tenant.
- Single module dependency for Graph (`Microsoft.Graph.Authentication` + `Invoke-MgGraphRequest`),
  not the per-resource Graph modules.
- Windows PowerShell 5.1 compatible. No PS7-only syntax (no ternary, no `??`, no scriptblock `-replace`).
- Where data cannot exist, the report says so instead of guessing.
- Executive-facing wording: states what is or is not in place.
- Brand: Arial, blue `#005D9C` headings and table header fills, nothing smaller than 10pt in tables.
  The Audit Hat logo (`AuditHat-Logo.png`, kept next to the script) is embedded in the report as a
  data URI so the HTML stays a single file; if the file is missing the report simply has no logo.
- Table filters are plain JavaScript inside the HTML; no external scripts, fonts or network calls.
- Privacy by construction: nothing is transmitted anywhere except Microsoft's own endpoints.

---

## Troubleshooting

| Symptom | Cause / fix |
|---------|-------------|
| `...cannot be loaded because running scripts is disabled` | Run `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass` first |
| Module install fails | Corporate proxy or no Gallery access. Install manually: `Install-Module Microsoft.Graph.Authentication -Scope CurrentUser` and `Install-Module ExchangeOnlineManagement -Scope CurrentUser` |
| Why does the script not just call `Connect-MgGraph` for browser sign-in? | On Windows the Graph module signs in through the Web Account Manager (WAM) broker. WAM silently picks the account signed in to Windows (often a personal Microsoft account) and fails with no error. The module's `Set-MgGraphOption -DisableLoginByWAM` switch is not honored in builds 2.35.1 and 2.38.1 (MSAL still reports `IsBrokerConfigured - True`). The script therefore drives the browser sign-in itself (see How the script works, step 2) and never touches WAM |
| No sign-in page opens after "BROWSER SIGN-IN" | Copy the address printed under that heading into any browser **on the same PC** (it must redirect back to localhost). If a local port cannot be opened, run with `-DeviceCode` |
| Sign-in page says the redirect URI is invalid | Something intercepts localhost traffic (some endpoint-security products). Run with `-DeviceCode` |
| Nothing opens after "DEVICE CODE SIGN-IN" (`-DeviceCode` or fallback) | Device code never opens a window by itself. Go to https://microsoft.com/devicelogin in any browser on any device and type the code printed in the console. It can take a few seconds for the code line to appear |
| `AADSTS700016: Application with identifier '14d82eec-...' was not found in the directory` | The tenant does not have Microsoft's Graph Command Line Tools app and will not add it. Register your own app (section above) and run with `-ClientId` and `-TenantId` |
| `AADSTS50194` single-tenant application | `-ClientId` given without `-TenantId`. Add `-TenantId <tenant GUID or domain>` |
| `AADSTS65001` or `AADSTS90008` after registering your own app | Admin consent was not granted on the app's API permissions page, or a permission is missing. Compare the list in step 4 |
| Defender for Endpoint: "permissions ... have not been consented" | The connector''s Defender for Endpoint permissions (Machine.Read, Software.Read, Vulnerability.Read) need one admin consent per tenant: `https://login.microsoftonline.com/<tenant>/adminconsent?client_id=ae173bf3-ba47-4ab8-ac54-3ea92fa6c337`. Without it the Defender device table is skipped and everything else still runs |
| Defender for Endpoint: "Not collected: device-code sign-in" | The Defender API needs a refresh token, which only the browser sign-in provides. Run without `-DeviceCode` |
| SPF/DKIM/DMARC show "Not available" | `Resolve-DnsName` is missing (non-Windows PowerShell host) or DNS is blocked. Run from a Windows PC with normal DNS |
| `Microsoft Secure Score: Not available` | The signed-in account needs Security Reader (or Global Reader) and the `SecurityEvents.Read.All` scope must be consented. Re-run as a Global Administrator and accept the consent prompt |
| Intune or PIM sections say "not licensed" | Derived from the tenant's subscribed SKUs: no Intune or Entra ID P1/P2 service plan was found. The remediation plan lists the licence as a Low finding |
| "Need admin approval" at sign-in | The tenant blocks user consent. A Global Administrator must run the script once and accept the consent prompt |
| Warning: permissions not granted | Same as above; the related sections show "not available" |
| `signInActivity not available` | Requires Entra ID P1. Last sign-in columns are blank |
| MFA registration "Not available" | Requires Entra ID P1 (`userRegistrationDetails` report) |
| Exchange section "Collection failed" | Account lacks View-Only Organization Management, or the EXO module is outdated: `Update-Module ExchangeOnlineManagement` |
| Inbox rule scan is slow | Use `-SkipInboxRules`; it is one call per mailbox |
| Word conversion fails | Word is not installed or is blocked by policy. The HTML versions are left in the output folder |

---

## Validation performed

- Parses cleanly under both Windows PowerShell 5.1 and PowerShell 7.5.
- PSScriptAnalyzer: no errors. Remaining warnings are `Write-Host` (intentional console progress) and
  plural function nouns.
- Bugs fixed from the original draft: the HTML-encode helper was named `H`, which is a built-in
  alias for `Get-History` and would have replaced every encoded value with command history; the nav
  rendered an empty sub-link for sections with no sub-items; assignment-filter labels could attach
  to the wrong group; the platform label for Android policies came out as "Android Enterprise Android
  device administrator" because a regex switch fell through; the `-WordDocs` `SaveAs2` call used
  `[ref]` arguments that fail on PowerShell 7; O(n²) device and user lookups replaced with hashtables.
- HTML helper functions smoke-tested offline under both PowerShell 5.1 and 7.5 (encoding, tables,
  key/value tables, platform labels, assignment formatting, Conditional Access target resolution).
- **Run end-to-end against a live tenant (Oct 31, 2026):** a Microsoft 365 Business Standard tenant without
  Intune or Entra ID P1. Browser sign-in, consent, all Graph sections, Secure Score, Exchange Online
  and the HTML build all completed. Findings from that run fixed in the script: empty tables rendered a
  blank row; the authentication-methods endpoint was wrong; the forwarding verdict compared a string to
  a number; Intune sections said "none exist" when the API refused (now "could not be read" with the
  licence reason); Secure Score text carried raw HTML. Not yet exercised live: Intune sections with real
  data, Conditional Access policies, PIM, `-WordDocs`, and `Get-InboxRule` on a large tenant.

## Possible next steps

- Parallel inbox-rule scan for large tenants
- Entra sign-in log sample showing legacy auth actually in use (P1 tenants)
- Export the Security Controls Summary as CSV for the audit workbook
- Replace Word COM with a docx library so `-WordDocs` works without Word

## License

Apache License 2.0 (see `LICENSE`). Provided as-is, without warranty. Review the output before
sharing it outside your organization; it contains user, device, and application inventory.
