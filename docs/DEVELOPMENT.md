# Development guide

How `Get-M365AuditEvidence.ps1` is put together, so a contributor (human or AI assistant) can add a
check, a section or a fix without reading all 3,000 lines first. Keep this file current: it is the
map that future work is built from.

## Ground rules

- **Read-only.** Delegated `*.Read.*` Graph scopes only, `Get-*` Exchange and Teams cmdlets only.
  Never add a write scope or a cmdlet that changes anything.
- **Windows PowerShell 5.1 and PowerShell 7.** No ternary, no `??`, no `-replace` with a
  scriptblock, no `ForEach-Object -Parallel`. Parse the file in both engines before committing.
- **One file, no external assets at run time.** CSS, JavaScript, SVG charts and the logo are inline.
- **Privacy by construction.** Every network call goes to Microsoft endpoints. Nothing is sent to
  Audit Hat. Generated reports, history files and CSVs are never committed (`.gitignore`).
- **Keep in sync on every change:** `README.md` (sections table, parameters, scopes, FAQ),
  `CHANGELOG.md`, `$ScriptVersion` (date-based `yyyy.MM.dd`), this file, and the header comment.
- **Screenshots** in `docs/screenshots` are sanitized ("Sample Bank"). Never commit real tenant data.

## Layout of the script (top to bottom)

| Region | What lives there |
|---|---|
| Header comment | Synopsis, section list, requirements, privacy and connector notes |
| `param(...)` | All switches. Adding one means adding a README parameters row |
| Module install | `Install-RequiredModule` for Microsoft.Graph.Authentication and ExchangeOnlineManagement (MicrosoftTeams is optional and never auto-installed) |
| `$Scopes` | Delegated Graph scopes. Adding one triggers a one-time admin re-consent in every tenant; document it in README and CHANGELOG |
| Run logs | `Start-Transcript` at start; `Complete-RunLog` (called at the end, before `exit 2`, after `-TestPermissions`, and from a script-level `trap`) renames the transcript after the client and writes the errors log from shown warnings, `$script:RunIssues` (full module messages) and new `$Error` entries |
| Sign-in | `Connect-GraphBrowser` (loopback OAuth code + PKCE, scope `https://graph.microsoft.com/.default` so one approval covers every permission in the app registration; `$Scopes` is the list checked against the token afterwards; a new Graph permission must also be added to the connector registration), `Request-GraphToken`, `Confirm-GraphToken` (refresh), `Request-ResourceToken` (same refresh token exchanged for Exchange Online, Defender for Endpoint, Teams admin API; `-RequireScope` reads the token's `scp` claim with `Get-TokenScope` and throws a consent error when the tenant has not approved that permission, because Microsoft still issues a token without it). Missing approvals collect in `$script:ConsentNeeded` and print at the end of the run with `Get-ConsentLink` (v2 adminconsent with explicit scopes). `-DeviceCode` fallback. Saved sign-in: `Save-SignInCache` / `Read-SignInCache` (DPAPI); the default cache file is per tenant when `-TenantId` is given (`M365AuditEvidence-<tenant>.signin`) and a cache saved for another named tenant is ignored |
| Graph helpers | `Get-GraphAll` (paging, optional `-MaxItems`), `Get-GraphAllSafe` (failure becomes a warning plus licence hint in `$script:GraphFailures`), `Get-GraphFailure`, `Get-GraphReportCsv` (usage reports), `Test-ConcealedReportNames` |
| HTML helpers | `HtmlEnc`, `Format-DateText`, `Join-List`, `New-InfoIcon`, `Get-RowIssueIcon`, `New-HtmlTable` (`-Id` + `-Filters` give a filter bar; rows carry `_tags` and `_issues`), `New-KvTable -Issues` |
| `$IssueCatalog` | Every info icon: key -> Title / Why / Fix. Written for a bank board, not an engineer |
| Collection | One block per data source, in report order: tenant, Secure Score, Intune, devices, Conditional Access, audit events, apps, users, roles, auth methods, Exchange (+ EOP/Defender policies, email auth DNS), Defender for Endpoint, indications of compromise, Teams, SharePoint/OneDrive |
| Evaluation | `Add-SummaryRow` per control; status strings drive colour via regexes in the summary table builder and `Get-StatusRank`; `$remediationGuides` + `Add-Remediation` build the plan; `$ControlPortalUrl` maps controls to the admin center page |
| Playbooks and frameworks | `$Playbooks` (control name -> Owner, Effort, Prereq, Steps, Verify), `$ScorePlaybooks` (regex -> playbook for Secure Score titles), `$GenericScorePlaybook`, `Get-Playbook`; `$FrameworkMap` (control -> Nist, Cis365, CisV8, Ffiec), `$ScoreFrameworkRules`, `Get-FrameworkText`; `$PhaseByPriority` / `$PhaseOrder` drive the Roadmap. `Add-Remediation` attaches all three automatically by finding name |
| Authentication analysis | After `$mfaRegistered`: `$MethodLabels`, `$PhishingResistantMethods`, `$WeakMethods`, MFA product detection (`$mfaProduct` from external authentication method configurations, federated domains, else Entra MFA), `$passkeyKv`, `$passkeyStatus`, `$adminPhishStatus` |
| Flagging | `_issues` hashtables added to rows, `$SummaryIssueMap` for summary rows |
| Executive figures | `Add-Kpi`, `$execNarrative` |
| History | `<Client> - history.json`: snapshot, diff vs previous (`$changeRows`), SVG charts (`New-SvgTrendChart`, `New-SvgGroupedBars`) |
| Rendering | `$sections` (nav), `$SectionAnchors`, `$css`, section here-strings (`$iocSectionHtml`, `$frameworkSectionHtml`, `$mdmSection`, `$teamsSectionHtml`, `$spoSectionHtml`), `$governanceKv`, the main `$html` here-string, the `<script>` block (filters, info panels, ticket copy/CSV, open step lists before print) |
| Output | HTML written with UTF-8, ticket CSV, history saved, report opened unless `-NoOpen`, optional Word documents |

## Conventions

- **Status words matter.** The summary table colours a status by regex: green for
  `Enabled|Enforced|Blocked|Yes|All enrolled|Restricted|Target met|Complete|Configured|Onboarded|None found`,
  amber for `Partial|Partly deployed|Allowed|...|Guests present|N to review`, red for `Disabled|Not enforced|...|Licensed, not deployed|N urgent|Anyone links`, green includes `Deployed`; `Not checked...` is grey.
  `Get-StatusRank` uses the same families to decide Improved / Regressed in the change log. A new
  status word must be added to both places, or it renders grey and never counts as a change.
- **Rows are ordered hashtables.** Hidden keys: `_tags` (space-separated filter tags) and `_issues`
  (column -> issue key). `New-HtmlTable` strips them from the output.
- **Every value is HTML-encoded** through `HtmlEnc`. Do not concatenate raw tenant data into HTML.
- **Here-strings expand `$`.** The JavaScript block and the CSS live inside expandable here-strings,
  so JavaScript must not contain a literal `$` or backtick. The `node --check` step below catches
  syntax errors, not interpolation surprises; grep for `$` in the script block after editing.
- **Edit scripts.** Large edits have been applied with small Python scripts that assert the anchor
  text occurs exactly once before replacing (String.Replace replaces every occurrence; a duplicated
  executive summary block came from forgetting that).
- **Status words for new controls:** `Complete`, `In progress (N%)`, `Not started`, `None found`,
  `N urgent`, `N to review`, `Guests present`, `Anyone links` are already in the regex families.
- **Playbook steps are written for the person receiving the ticket:** admin center > page > setting
  > value, in click order, then a "done when" that names the report row that will change.
- **Framework references:** CIS M365 Benchmark numbering follows v4.0 at family level (for example
  5.2.2) unless the item is stable across editions; FFIEC CAT / ACET codes are at domain.component
  level (D3.PC.Am). Keep the disclaimer in section 5 when editing.
- **Never wrap a `List[object]` of custom objects in `@()`.** `@($list)` throws "Argument types do not
  match" in both engines when the list holds `[pscustomobject]` items; use `$list.ToArray()`.
  Piping the list (`$list | Where-Object ...`) is fine.
- **Numbers in section titles** appear in three places: `$sections`, the `<h2>` of each section
  template, and the README table. Renumber all three.

## Adding a check (the seven steps, in code)

```powershell
# 1. collect
$thing = Get-GraphAllSafe 'v1.0/some/endpoint?$select=a,b' 'my thing'
# 2. status + evidence
$thingStatus = if (Get-GraphFailure 'my thing') { 'Not available' } elseif ($thing.ok) { 'Enabled' } else { 'Disabled' }
# 3. summary row (also feeds KPI, history, change log)
Add-SummaryRow 'My control name' $thingStatus "Why we say so: $($thing.detail)" 'Section title'
# 4. explain (info icon) -> $IssueCatalog['my-issue'] = @{ Title=..; Why=..; Fix=.. }; $SummaryIssueMap['My control name'] = 'my-issue'
# 5. finding -> add to $remediationGuides: @{ Control='My control name'; Bad='^Disabled'; Priority='High'; Action='...'; Url='https://...' }
# 6. (optional) section -> $sections entry, $SectionAnchors['Section title'] = 'id', <section id='id'> in the template
# 7. document -> README table + CHANGELOG + $ScriptVersion + this file
```

## Verifying a change

```powershell
# parse in both engines
$t=$null;$e=$null;[System.Management.Automation.Language.Parser]::ParseFile('.\Get-M365AuditEvidence.ps1',[ref]$t,[ref]$e)|Out-Null;$e.Count
powershell.exe -NoProfile -Command "`$t=`$null;`$e=`$null;[System.Management.Automation.Language.Parser]::ParseFile('$PWD\Get-M365AuditEvidence.ps1',[ref]`$t,[ref]`$e)|Out-Null;`$e.Count"
# static analysis (these rules are intentionally excluded)
Invoke-ScriptAnalyzer -Path .\Get-M365AuditEvidence.ps1 -ExcludeRule PSAvoidUsingWriteHost, PSUseSingularNouns, PSUseShouldProcessForStateChangingFunctions, PSAvoidUsingPositionalParameters, PSAvoidUsingConvertToSecureStringWithPlainText
# JavaScript inside the report: extract the <script> block to a file and run
node --check report.js
```

For rendering, build a page from the section here-strings with sample variables (see
`sections-test.ps1` in the session scratchpad for the pattern) and screenshot it with headless
Chrome: `chrome --headless=new --screenshot=out.png --window-size=1500,2400 file:///...html`.

A full live run needs a person at the browser for MFA; it cannot be driven from a background
session. Ask the user to run `pwsh -NoProfile -File .\Get-M365AuditEvidence.ps1` and report back.

## Data sources worth knowing

- Graph returns bare 400/403 for unlicensed features. Licence flags from `subscribedSkus`
  (`$script:HasIntune`, `HasEntraP1`, `HasEntraP2`, `HasDefenderP1`, `HasDefenderP2`, `HasMDE`)
  turn those into explanations. `INTUNE_O365` is not real Intune.
- Directory audit log: 30 days with Entra ID P1/P2, 7 days on Free. Sign-in log needs P1.
  Identity Protection needs P2. Secure Score needs `SecurityEvents.Read.All`.
- Usage reports (`/reports/get*Detail(period='D90')`) return CSV behind a redirect, lag up to
  48 hours, and conceal names (hashes) when the tenant's report privacy setting is on;
  `Test-ConcealedReportNames` detects that and the report explains how to turn it off.
- Defender for Endpoint and the Teams admin API use the same refresh token through
  `Request-ResourceToken`; a missing consent surfaces as AADSTS65001 and the report prints the
  admin-consent URL instead of failing.
- Exchange Online: `Connect-ExchangeOnline -AccessToken` with a token for
  `https://outlook.office365.com/.default` needs the connector to hold the delegated
  `Exchange.Manage` permission; otherwise the script falls back to the interactive prompt.
