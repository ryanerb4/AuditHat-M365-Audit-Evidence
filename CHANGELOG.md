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
- Remediation Plan as tickets: stable `AH-xxxxxx` reference per finding, Copy ticket and Copy all
  buttons (ready-to-paste text for the IT provider or MSP), priority filters and search, CSV download
  in the report and the same CSV written next to the report.
- Per-tenant run history with "Progress since last report": change log of controls, findings and
  headline counts, plus Secure Score and open-findings trend charts (inline SVG).
- Licence-aware explanations when Graph refuses unlicensed features (Intune, Entra ID P1/P2).
- Client identified from the tenant display name and default domain; Audit Hat logo and links.
- Windows PowerShell 5.1 and PowerShell 7 compatible; modules auto-install for the current user.