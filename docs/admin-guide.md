# Administrator guide

The admin surface is **Microsoft Lists** — no separate admin app to deploy.
You create two lists once; coordinators point the add-in at the site in
Settings.

## 1. Create the routing list

On your chosen SharePoint site, create a list named `LegislativeRoutingMatrix`
with these columns (all "Single line of text" unless noted):

| Column | Notes |
| --- | --- |
| Title | = DivisionCode (SharePoint's built-in Title column) |
| DivisionCode | MVD, TDD, SOD, ELT, AG Office, … |
| DivisionName | Motor Vehicle Division, … |
| Aliases | semicolon-separated alternative names ("Motor Vehicle; Motor Vehicles") |
| Emails | semicolon-separated targeted-email recipients |
| TeamsTeamId | Teams team GUID (Teams → team → ⋯ → Get link to team → `groupId=`) |
| TeamsChannelId | channel ID (⋯ on the channel → Get link to channel → `19:...@thread.tacv2`) |
| TeamsChannelName | display name (for the review UI) |
| TeamsTagId | tag GUID — use the add-in's Settings → "Teams tag ID lookup" |
| TeamsTagName | tag display name as it appears in Teams |
| CodeChapters | semicolon-separated Iowa Code chapters this division owns ("321; 321A; 322") — powers division SUGGESTIONS for unrouted bills (optional) |
| MentionUserIds | semicolon-separated Entra object IDs (optional, for individual mentions) |
| MentionUserEmails | matching display emails (optional) |
| IsActive | Yes/No |
| Priority | Number — highest wins when multiple rules match a division |
| EffectiveStartDate | Date (optional) |
| EffectiveEndDate | Date (optional) |
| Notes | free text |

**Rule semantics:** a division token from the report matches a rule when it
equals the DivisionCode, DivisionName, or any alias (case-insensitive). Among
matches that are active and within their effective dates, the highest
Priority wins. Session handoffs = add the new rule with a later
EffectiveStartDate and end-date the old one; history stays intact.

## 2. Create the audit list

Same site, list named `LegislativeAudit`, columns (Single line of text):

```
Title (bill number)   ReportKey   IdempotencyKey   TeamId   ChannelId
TeamsMessageId   Status   Error   Divisions   EmailRecipients
PublishedBy   SourceSubject
```

The add-in appends one row per publish attempt. **Do not delete rows** for
reports that may be re-run — the IdempotencyKey rows are what prevent
duplicate posts.

## 2b. Create the bill tracker list (recommended)

Same site, list named `BillTracker` — this is the shared "who's waiting on
who" board and replaces individual completeness spreadsheets. Columns:

| Column | Type | Notes |
| --- | --- | --- |
| Title | text | bill number (written by the add-in) |
| Division | text | one row per bill × division |
| Status | **Choice** | `Pending review` / `In review` / `Commented` / `No comment needed` |
| DueDate | Date | auto-set to +2 business days at publish |
| BillLink | text | BillBook URL |
| Brief | text | first 250 chars |
| ReportKey | text | which report published it |

The add-in writes one `Pending review` row per bill × division at publish
time; division staff update Status themselves.

**Pin it in Teams:** in the legislative channel, **+ Add a tab → Lists →
existing list → BillTracker**. Create these views:

- *My division* — filter Division = yours (the "filter tagged bills" ask)
- *Still waiting* — Status = Pending review, grouped by Division, sorted by
  DueDate (the cross-divisional "who's waiting on who" board)
- *Overdue* — Pending review AND DueDate < today

## 3. Import the existing Excel routing matrix

Microsoft Lists imports Excel directly: List → ⋯ → "Export/Import" or create
the list *from* the Excel file, then rename columns to match the table
above. For tag IDs, use the add-in's tag lookup (Settings → paste the Team
ID → Fetch tags) and copy each `TeamsTagId` into the matching row.

## 4. Permissions to grant

- Coordinators: **edit** on both lists (audit rows are written as them).
- Division reviewers: no list access needed — they just receive posts/emails.
- The Entra app needs one-time admin consent (`docs/permissions.md`).

## 5. Validating a route

In the add-in: Settings → connect the site → parse a test report → the
Review screen shows exactly which rule each division resolved to and flags
rules missing tag IDs. Preview shows the rendered post before anything is
published.

## 6. Using a different state (or Congress)

The add-in is state-configurable:

1. Settings → **State / legislature** — pick your state. This presets the
   bill identifier prefixes the parser recognizes (editable — e.g. add
   "LB" for Nebraska's unicameral if needed).
2. **New-filings feed:** Iowa uses the Legislature's own RSS. Every other
   state uses the Open States API via the deployment's feed mirror
   (verified live: Iowa 1,917 and Texas 1,185 bills in a 1-year window).
   Congress: the parser preset works for federal bill reports, but Open
   States carries state data only - a congress.gov API adapter is a
   planned addition for the federal new-filings feed. Mirror setup: add the state's name to `states.json` in the repository
   and set the `OPENSTATES_API_KEY` repository secret (free key from
   openstates.org). The scheduled mirror then publishes
   `feeds/openstates-<state>.json` automatically.
3. Everything else (routing matrix, tracker, audit) is already
   org-specific — each organization points at its own SharePoint site.
4. **Org profile:** after configuring, use Settings → Org profile → Copy
   profile, and send the code to your team; they paste + Apply and are
   fully configured.

## 7. Code chapter tracking

Bill briefs reference the Iowa Code ("Amends Code Chapter 9", "Code
chapters 6A and 6B"). The add-in extracts these references and:

- shows them as chips on each bill in Review (checkmarked when they hit
  the agency's tracked list),
- flags newly filed bills in the Filings tab whose descriptions touch
  tracked chapters (even when no watch term matches),
- suggests divisions for unrouted bills when routing rules claim
  chapters via the optional `CodeChapters` column.

The tracked list ships seeded from the DOT's 2015 tracking list
(Settings → "Tracked Code chapters") and is FULLY EDITABLE — the 2015
list predates several Code changes, so review it with your legislative
coordinator and paste the updated set. It travels with org profiles.
