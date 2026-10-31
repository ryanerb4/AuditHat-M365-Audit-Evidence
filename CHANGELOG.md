# Changelog

All notable changes to the AuditHat M365 Audit Evidence script. Versions are date-based
(`year.month.day` of the last change) and appear in the console banner, the report navigation and
the report footer.

## 2026.10.31

First public release.

- Read-only collection across Entra ID, Intune, Conditional Access, enterprise applications and
  consent audit events, users, privileged roles, Exchange Online and Defender.
- Browser sign-in with MFA / Microsoft Authenticator through the script's own OAuth
  authorization-code flow (bypasses the Windows Web Account Manager problem in the Graph module);
  device-code fallback.
- Default sign-in app: AuditHat Pulse Reports (PowerShell), a multi-tenant public client with
  delegated read-only permissions; `-ClientId` / `-TenantId` and a settings file for tenant-owned
  registrations.
- Executive Summary with narrative, KPI tiles, top five priorities and privacy statement.
- Security Controls Summary with computed status and evidence per control.
- Remediation Plan: failed controls, privileged-access and licensing findings, and every open
  Microsoft Secure Score action, ranked Critical to Low.
- Microsoft Secure Score with configurable target (`-TargetScorePercent`), gap and comparison
  averages.
- Clickable info icons on every flagged value, backed by a catalogue of about fifty issues with
  "why it matters" and "recommended" text.
- Filterable user table (Stale, No MFA, Admin, Guests, Disabled, Licensed, search) with devices per
  user; combined device list with owner, last user, last activity and filters.
- Email authentication per domain: MX, SPF, DKIM (Exchange signing and DNS selectors), DMARC.
- Defender and Email Protection section: licences, preset security policies, anti-spam, anti-malware,
  anti-phishing, Safe Links, Safe Attachments, Defender for Endpoint devices and scores.
- Directory synchronization status (Entra Connect, on-premises AD domains, sync server) and a link to
  the companion AuditHat AD Governance Report.
- `-TestPermissions`: probes every endpoint, token, report, module and licence the report uses and
  prints what is missing.
- Entra ID Free fallback: when the P1 registration report is refused, MFA registration, MFA product and
  passkey facts are built from each user's authentication methods (batched Graph calls).
- New Critical control "Audit logging (unified audit log)" with playbook when ingestion is off; Defender
  for Endpoint reports "Not licensed" when the API says so; SPF/DKIM/DMARC read Partial rather than
  Missing when a weak record exists.
- Paged navigation: one section per page with Previous / Next, a "Show all sections" toggle (remembered
  in the browser), in-report links that open the right page, and whole-report printing.
- Indications of compromise is section 2, directly after the Executive Summary, and opens with a
  situation table (Urgent / Review / Clear / Not checked, what we found, why it matters). New
  indicators: mail forwarded outside the organization, known-abused OAuth applications (matched by
  application ID), unverified-publisher apps with mail or file access, Anyone links and external
  sharing events, unusual download and sync volume per user, mailbox rule and permission changes with
  actor and IP address (all from the unified audit log via Search-UnifiedAuditLog), and malicious
  link clicks from Safe Links. `-SkipAuditLog` skips the audit log search.
- This-week (Critical) defaults for financial institutions: Teams guest and external access, device join
  open to all users, single Global Administrator, mobile device compliance (new control "Intune
  compliance policies for mobile devices"), and Secure Score actions for anonymous meeting join, lobby
  admission, more than one global admin and user consent. `PriorityOverrides` in the settings file
  re-tiers any finding by title.
- "Review These Now: Possible Indications of Compromise" is section 3, ahead of the Remediation Plan
  (now section 4).
- View ticket: each finding opens a preview of the full ticket text with a Copy button, instead of
  copying blind.
- Email authentication split into three controls and findings: SPF record (DNS), DKIM signing
  (Exchange Online and DNS, linked to the Defender DKIM page) and DMARC policy (DNS), each with its own
  playbook, framework mapping and guidance link.
- Priorities: User consent to applications and Intune compliance policies are Critical; Intune not
  licensed is High. New High finding when usage-report names are concealed, with the admin-center
  path to clear the setting.
- Remediation Plan as a roadmap: findings grouped into this week / 30 days / 90 days / 12 months,
  and every finding carries a playbook (owner, effort, prerequisite, numbered steps in the admin
  centers, "done when" test) plus framework references; Secure Score actions are matched to the same
  playbooks. Tickets and the CSV include steps, owner, phase and standards. Links name the admin center.
- Section 5 "Framework Alignment": every control mapped to NIST CSF 2.0 / SP 800-63B, CIS Microsoft 365
  Foundations Benchmark, CIS Controls v8 and FFIEC / NCUA (IT Handbook, 2021 Authentication Guidance,
  CAT / ACET domains), with Met / Gap filters and a statement of what the report does not test.
- Authentication Methods: MFA product detection (Entra MFA, federated provider, third-party external
  authentication method), registered methods by strength, phone-only MFA users, Authenticator
  number-matching settings, and passkey / passwordless rollout tracking (policy, users, administrators,
  stage). New controls "Phishing-resistant MFA (passkeys) rollout" and "Phishing-resistant MFA for
  administrators", a KPI tile, a monthly trend chart and history tracking.
- Executive Summary: "How to use this report" by role (board, IT committee, security officer / MSP,
  auditor, examiner), a Governance and review table (report number, previous report, progress, open
  findings by phase, monthly cadence, next report due, standards referenced) and a sign-off block.
- Recommended cadence is now monthly; the README has an operating model table.
- Section 4 "Possible Indications of Compromise": directory audit log (role assignments, MFA method
  changes, password resets, policy changes, domain/federation changes, app credentials and owners,
  account lifecycle), app registration secrets and certificates added, new enterprise apps and accounts,
  Entra ID Protection risky users and detections (P2), sign-in log analysis (P1: countries per user,
  legacy sign-ins, MFA denials, Conditional Access blocks, password sprays, lockouts, risky sign-ins),
  and Exchange: inbox rules that hide, delete or forward mail, Full Access delegates, Send As grants,
  transport rule changes, mailbox auditing. Rolled up as urgent / to review with a KPI tile, summary
  row, Critical/High findings and history tracking.
- Section 16 "Microsoft Teams": guest and external access posture, every team with owners, members,
  guests, archived state and 90-day activity; optional tenant policies through the MicrosoftTeams
  module (federation, guest access, anonymous meeting join).
- Section 17 "SharePoint and OneDrive": tenant sharing settings, users sharing externally, every site
  and every OneDrive with the owner's account state, from the Microsoft 365 usage reports.
- Single sign-in: the browser sign-in's refresh token is exchanged for Exchange Online (and Teams)
  tokens; a separate Exchange prompt only appears when the connector lacks the Exchange.Manage
  permission, with the consent link printed.
- The report opens in the default browser when the run finishes (`-NoOpen` to suppress); the ticket
  CSV is linked from the Remediation Plan.
- Remediation links name the admin center they open (Defender, Entra, Intune, Exchange, Teams,
  SharePoint, Microsoft 365) and every control-based finding links to the page where it is changed.
- New Graph scopes: IdentityRiskyUser.Read.All, IdentityRiskEvent.Read.All, Reports.Read.All,
  SharePointTenantSettings.Read.All, TeamSettings.Read.All (all read-only; one-time admin consent).
- Sections renumbered: MDM Compliance Policies is now 5 and Defender and Email Protection 15.
- Remediation Plan as tickets: stable `AH-xxxxxx` reference per finding, Copy ticket and Copy all
  buttons (ready-to-paste text for the IT provider or MSP), priority filters and search, CSV download
  in the report and the same CSV written next to the report.
- Per-tenant run history with "Progress since last report": change log of controls, findings and
  headline counts, plus Secure Score and open-findings trend charts (inline SVG).
- Licence-aware explanations when Graph refuses unlicensed features (Intune, Entra ID P1/P2).
- Client identified from the tenant display name and default domain; Audit Hat logo and links.
- Windows PowerShell 5.1 and PowerShell 7 compatible; modules auto-install for the current user.