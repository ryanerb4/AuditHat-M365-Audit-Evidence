# Changelog

All notable changes to the AuditHat M365 Audit Evidence script. Versions are date-based
(`year.month.day` of the last change) and appear in the console banner, the report navigation and
the report footer.

## 2026.11.04

- Defender for Endpoint shows a clear deployment status in section 3 and in a new "Endpoint
  protection" block in section 16: Not licensed, Licensed, not deployed, Partly deployed,
  Attention needed, Deployed, or Not checked (permission not approved, access refused, device-code
  sign-in), with what it means and what to do. The licensed plan is named (Defender for Business,
  Plan 1, Plan 2). Coverage compares onboarded devices with active Windows and macOS devices in
  Entra ID and Intune and lists the ones not onboarded. Devices only discovered on the network are
  no longer counted as onboarded. When Defender could not be checked, the first page lists it as a
  coverage gap. The approval link now requests exactly the three Defender permissions.
- Defender for Endpoint status no longer depends on the Defender API alone. Licence detail
  (licences assigned, whether Microsoft has finished setting the plan up) and the Intune EDR
  onboarding policy give: Licensed, setup not finished; Licensed, not assigned; Licensed, no
  onboarding policy; Onboarding policy in place (devices not verified). The run output leads with
  that finding; a missing permission is a secondary note.
- Exchange Online, Teams and Defender tokens are checked for the expected permission before use.
  A tenant that approved only the Graph permissions now gets "the connector's Exchange permission
  is not approved" and a fallback Exchange sign-in instead of "UnAuthorized"; an UnAuthorized
  answer otherwise explains the Exchange role needed. `-NonInteractive` skips Exchange instead of
  prompting.
- Sign-in asks for everything in the app registration (`.default`) instead of a list of Graph
  permissions, so a new organization's Global Administrator approves Graph, Exchange Online,
  Defender for Endpoint and Teams on one screen at first sign-in. Missing permissions print the
  single approval link.
- Inbox rules and Full Access permissions look each mailbox up by its unique ID instead of its
  name, so two mailboxes that share a name or alias (for example "ryan") no longer fail with
  "isn't unique".
- End of run: "One-time approval needed in this tenant" lists each service whose permission is
  missing, with one approval link per service.

- Saved sign-ins are kept per tenant when `-TenantId` is given, so auditing a second organization
  from the same PC no longer reuses the first organization's sign-in. A saved sign-in for a different
  named tenant is ignored with a warning.
- README: new "Auditing more than one organization" section (consent link, account, `-TenantId`
  on every run, separate output folders, settings file caveat); corrected the Teams section
  number, the Framework Alignment filter names and the connector name in the quick start.

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
- Consistent pass / fail indicators: one classifier decides whether a status is Pass, Partial: review,
  Action needed or Not measured; every Status, Result and Level cell shows an icon badge in the same
  colours; the Security Controls Summary and Framework Alignment tables gain a Result column, filters by
  result and a legend, and the first page carries the same legend.
- First page: rocks list their open items as links to the plan; issues, repeat findings and top
  priorities link to their plan row and carry a View ticket button.
- The signature block was removed from the Executive Summary; governance is recorded in the committee and board minutes.
- Saved sign-in for scheduled runs: `-SaveSignIn` keeps the refresh token encrypted with Windows DPAPI for
  the current user and PC; later runs sign in silently and keep the rotated token; `-NonInteractive`
  never prompts (exit code 2 when the saved sign-in is unusable); `-ForgetSignIn` removes it;
  `-SignInCachePath` sets the location. Task Scheduler example in the README.
- Executive Summary rebuilt as a one-page board view in the EOS pattern: verdict (Act now / Needs
  attention / In good shape) with reason and chips; four headline tiles with the rest folded; Headlines
  (good and bad news); To-dos this week; Scorecard (goal, actual, on/off track, movement since last
  report); Rocks (90-day objectives with owner, due date, status); Issues list (licence and policy
  decisions, repeat findings, coverage gaps); progress charts with the change log folded; plan by phase;
  mapping to the GLBA annual board-report elements and to what examiners look for in board reporting.
- Review These Now: when the unified audit log is off, the sharing, download and mailbox-rule rows read
  Urgent ("cannot be checked") instead of Not checked, and the Anyone-links row combines the audit log
  with the section 19 share scan.
- Section 19 "Sharing and Access": every share and access grant across SharePoint sites, OneDrive,
  Teams and Exchange (sharing links and item permissions, guest and public teams, shared-channel
  external users, Full Access and Send As, calendar Default/Anonymous permissions), classified Anyone /
  Whole organization / External / Specific people with outside-owning-team and cross-department flags,
  guests with memberships, cross-department matrix. New scope Sites.Read.All; `-SkipShareScan`.
  Findings "Content open to anyone on the internet" (High) and "Content open to the whole
  organization" (Medium) with playbooks.
- Enterprise Applications table redesigned: one application column (name, publisher, type, sign-in,
  assignment, assigned users), a Risk level column, all permissions in one column with the first four
  inline and a Show all dialog grouped by permission type, app IDs removed, risk and consent filters.
  Credential additions show who added them; new enterprise apps show who added or consented.
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