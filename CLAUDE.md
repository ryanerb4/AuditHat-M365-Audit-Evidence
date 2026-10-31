# CLAUDE.md

Working notes for AI assistants and contributors on this repository.

## What this is

`Get-M365AuditEvidence.ps1` is Audit Hat's free, read-only Microsoft 365 audit evidence collector
for banks and credit unions. One run produces one self-contained HTML report (executive summary,
controls, remediation plan with ticket export, indications of compromise, and 13 evidence sections),
a ticket CSV and a per-tenant history file. See `README.md` for the user view and
`docs/DEVELOPMENT.md` for the code map, conventions and verification steps. Read both before
changing the script.

## Non-negotiables

- Read-only: no write scopes, no `Set-*`/`New-*`/`Remove-*` against the tenant.
- No data leaves the customer's organization. The script has no server side and must stay that way.
- Windows PowerShell 5.1 compatible as well as PowerShell 7.
- Every change updates `README.md`, `CHANGELOG.md`, `$ScriptVersion` (date-based) and, when the
  layout or conventions change, `docs/DEVELOPMENT.md`.
- Never commit generated reports, history files, CSVs, settings files or real tenant names.
  Screenshots use "Sample Bank" data.

## Working practices

- Verify before committing: parse in PS 5.1 and PS 7, run PSScriptAnalyzer with the exclusions
  listed in `docs/DEVELOPMENT.md`, `node --check` the report JavaScript, and render changed sections
  with sample data in headless Chrome.
- Live runs need the user at the browser for MFA. Ask them to run the script from their own
  terminal; do not try to complete a sign-in from a background session.
- Edit with scripts that assert each anchor string occurs exactly once; do not use blanket
  string replacement on this file.
- Write report text for a bank board: plain language, what is or is not in place, why it matters,
  what to do. Status words are matched by regex for colour and trend; see the conventions section
  in `docs/DEVELOPMENT.md`.
