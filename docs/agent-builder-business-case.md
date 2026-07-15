# Business case: enable Copilot Agent Builder for the legislative pilot

**Request:** Enable Microsoft 365 Copilot Agent Builder for a named pilot
group (Office of the Director legislative staff + Systems Operations
sponsors, ~10 users), or alternatively have IT build and share one
declarative agent on the pilot's behalf.

**Cost:** $0. Declarative agents are included in the M365 Copilot licenses
the DOT already owns. This is a settings change, not a purchase. (Copilot
Studio, which does carry cost, is NOT being requested.)

## The use case

During legislative session, every division must review and comment on
assigned bills within 48 business hours. Staff currently read bill text
manually from the Daily Bill Report. A "Bill Summarizer" agent in the
legislative Teams channel would let any member request a plain-language
summary of a bill or amendment — posted in-channel where the whole review
team sees it — directly supporting the State & Federal Relations office's
enhancement request for the 2027 session (R. Jerman, July 2026).

## Why the risk is contained

1. **Internal grounding only.** The agent's knowledge sources are limited to
   the legislative SharePoint site. Web grounding stays off. It cannot see
   any content the requesting user couldn't already open — declarative
   agents run entirely on the caller's own permissions.
2. **No autonomy.** Declarative agents only respond when invoked; they take
   no actions, run no code, call no external APIs, and store nothing
   outside the tenant.
3. **Public-record content.** The ground truth is Iowa legislative bills —
   public documents. No PII, no CJIS/tax/health data in scope.
4. **Granular admin control.** The Microsoft 365 admin center allows
   enabling agent creation for a specific security group only (not
   tenant-wide), separately controlling who may *use* shared agents, and
   requiring admin approval before an agent is shared beyond its creator.
   IT retains an inventory of all agents and can disable any of them.
5. **Auditability.** Copilot interactions are logged in Purview like other
   Copilot usage; the agent adds no new data flows.

## Proposed rollout

- **Phase 1 (July–Aug 2026):** enable creation for the pilot security
  group, or IT creates the agent from the one-page spec
  (`docs/copilot-agent.md`) and shares it to the Legislation 91st GA team.
- **Phase 2 (before session, Jan 2027):** review pilot feedback and Purview
  logs; decide whether to keep the agent, broaden availability, or retire.

## The alternative if declined

The legislative workflow tooling functions without it (official LSA bill
descriptions are included in every routed Teams post), but staff will
continue summarizing bill text manually — the single most time-consuming
step the enhancement request asked us to address.
