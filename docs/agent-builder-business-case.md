# Business case: enable Copilot Agent Builder for the legislative pilot

**Request:** Enable Microsoft 365 Copilot Agent Builder for a named pilot
group (~10 users), or alternatively have IT build and share one declarative
agent on the pilot's behalf.

**Business sponsor:** Renee Jerman, State and Federal Relations Compliance
Officer, Office of the Director.

**Cost:** $0. Declarative agents are included in the M365 Copilot licenses
the DOT already owns. This is a settings change, not a purchase. (Copilot
Studio, which does carry cost, is NOT being requested.)

## This responds to a documented business request

From the sponsor's enhancement request of July 10, 2026 ("Possible
enhancements for the TEAMS channel, legislative comment system?", attached):

> "Co-Pilot app within the TEAMS Channel that could provide a quick summary
> of the bill and amendment with the summary output from Co-Pilot placed in
> the TEAMS channel for all to see."

IT is being asked to enable a capability the business has already requested
in writing — not to adopt a technology looking for a use.

## The workload, quantified

The 91st General Assembly has filed **4,224 bills to date** (2,367 in the
2025 session; 1,857 so far in 2026 — counted from the Legislature's own
newly-filed feed). The Daily Bill Report routes the DOT-relevant slice to
divisions on a **48-business-hour comment deadline**; entries range from
one-paragraph briefs to multi-page appropriations text (e.g., HF2792, the
RIIF bill, spans seven DOT appropriation line items across three
divisions).

Conservatively: ~400 DOT-routed bills and amendments per session × 15
minutes of manual reading × 2–3 reviewers per bill ≈ **300–450 staff-hours
per session** spent on first-pass comprehension — the exact step a
grounded summarization agent compresses to seconds, ahead of a deadline
measured in hours.

## This reduces AI risk rather than adding it

Staff working against a 48-hour clock already have every incentive to paste
bill text into whatever consumer AI tool they personally use — ungoverned,
unlogged, and outside the tenant. Bills are public documents, so today that
usage is invisible rather than impossible. An approved agent **displaces
unmanaged AI use with managed use**: grounded only on the legislative
SharePoint site, running under each caller's existing permissions, logged
in Purview, and disable-able by IT at any time. Declining the request does
not prevent AI usage; it only prevents the governed kind.

## Why the risk is contained

1. **Internal grounding only.** Knowledge sources are limited to the
   legislative SharePoint site; web grounding stays off. The agent cannot
   see content the requesting user couldn't already open.
2. **No autonomy.** Declarative agents respond only when invoked; they take
   no actions, run no code, call no external APIs, and store nothing
   outside the tenant.
3. **Public-record content.** Ground truth is Iowa legislative bills. No
   PII, CJIS, tax, or health data in scope.
4. **Granular admin control.** Enable creation for one security group only
   (not tenant-wide); separately control who may use shared agents; require
   admin approval for sharing; full agent inventory in the admin center.
5. **Auditability.** Interactions log to Purview like all Copilot usage; no
   new data flows.

## Pilot plan and measurable exit criteria

**Phase 1 (July–August 2026):** enable creation for the pilot group, or IT
creates the agent from the one-page spec (`docs/copilot-agent.md`) and
shares it to the Legislation 91st GA team.

**Phase 2 (December 2026, before session):** go/no-go review against
defined criteria:

- Agent invoked on ≥ 50% of bills routed during the pilot window
- Zero Purview/DLP findings attributable to the agent
- Median time from bill posting to first division comment decreases
  (baseline: current tracker data)
- Pilot participants rate it worth keeping (short survey, ≥ 70% yes)

If criteria aren't met, the agent is retired and the enablement reversed —
a settings change in both directions.

## The alternative if declined

The legislative routing tooling functions without it (official LSA bill
descriptions are included in every routed Teams post), but staff continue
manually summarizing bill text under deadline — the single most
time-consuming step the sponsor's enhancement request asked us to address —
with the shadow-AI exposure described above left unmanaged.
