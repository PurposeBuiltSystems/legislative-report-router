# Legislative Report Router

*From bill report to the right people — without manual routing.*

Outlook add-in for legislative coordinators: parse a Daily Bill Report email
(or DOCX attachment), review the extracted bills and their division routing,
then publish — one Teams post per bill with **real tag/user mentions**,
consolidated targeted emails per division, optional send of the original
report, idempotent re-runs, and a durable SharePoint audit trail.

## Architecture

**No backend.** Task pane + MSAL nested app authentication + delegated
Microsoft Graph as the signed-in coordinator. Rationale: every operation the
workflow needs (parse, route, post to channels, send mail, read/write
SharePoint lists) is available with delegated permissions; removing the
backend removes Azure cost, a second auth surface, and a data processor from
the privacy story. Cloud endpoints are configurable — Commercial, GCC, GCC
High, DoD (see `docs/government-cloud.md`).

Pure-logic modules (offline-testable, no Office/Graph imports):

- `src/parser.js` — staged pipeline: HTML→text (links kept) → header strip →
  **line-aware entry boundaries** (a bill number alone on a line; internal
  references like "Successor to HSB171" can never split an entry) → field
  extraction → division normalization (`MVD/TDD`, `ELT and MVD`, …) →
  confidence + warnings. See `docs/parser-rules.md`.
- `src/routing.js` — rule matching: aliases, priority, effective dates,
  active flags; storage-agnostic rule shape (SharePoint today, Dataverse/SQL
  later via the same mapper seam).
- `src/teams.js` — Graph `chatMessage` payloads with real mention entities
  (`<at id>` ↔ `mentions[]`, tag + user, deduped), division email builder,
  idempotency keys. Templates live here, separate from wiring.
- `src/docx.js` — client-side DOCX text extraction (zip + XML; no Word, no
  macro formats, nothing executed).
- `src/graph.js` — auth + SharePoint lists + channel posts + mail, with
  throttle retry and cloud endpoint map.
- `src/taskpane/` — Overview / Review / Preview / Publish / Audit screens.

## Configuration

Two SharePoint lists on a site of your choice (see `docs/admin-guide.md` for
column-by-column setup):

- **LegislativeRoutingMatrix** — one row per division route (code, aliases,
  emails, Teams team/channel/tag IDs, mention users, priority, effective
  dates). This list — edited in Microsoft Lists — is the admin interface.
- **LegislativeAudit** — written by the add-in on every publish attempt;
  doubles as the idempotency store (a published key is never re-posted).

## Safety properties

- Nothing publishes without the explicit confirmation checkbox.
- Unknown divisions are prominent in Review; unmatched items can't silently post.
- Idempotent: re-running a report skips everything already published.
- Partial failure preserves successes; "Retry failed only" retries the rest.
- All user content is HTML-escaped in Teams posts and emails; only the
  mention markers are markup.
- No data leaves the tenant: no publisher server, no analytics, no AI.

## Tests

`node test/router.test.js` — 50+ checks: boundary detection against the
representative report, internal-reference safety, multi-division parsing,
HTML entities/links, routing aliases/priority/dates, mention payload schema
and dedupe, HTML-injection escaping, idempotency keys, SharePoint field
mapping, email grouping.

## Deployment

Internal tool: sideload `manifest.xml` (aka.ms/olksideload) or deploy
centrally via the M365 admin center. Entra app is multitenant; org admins
consent once (`docs/permissions.md` has the consent URL and scope matrix).
