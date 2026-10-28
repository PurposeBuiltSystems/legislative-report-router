# Administrator handoff: Teams tab and contact-sync requirements

[Download the administrator handoff](../teams/Legislative-Router-Admin-Handoff.zip). It includes the ready Teams app ZIP, this explanation, and a permission/architecture review request. **Only the inner Legislative-Work-Queue.zip is uploaded to Teams.**

## Ready now: staff queue in Teams

An administrator can publish the enclosed Teams ZIP to the organization's app catalog and make it available to the intended users/Team. A missing Upload a custom app option is governed by Teams policy, not the Outlook manifest. Publishing centrally does not require allowing custom uploads for everyone.

The Teams ZIP contains its own manifest.json and required icons. It reuses the existing Entra application client ID `0860a653-ddbd-4455-8bff-affda2a8879f`. The deployed staff queue requires only the already-used delegated Sites.ReadWrite.All permission. The Outlook XML manifest is unchanged. [Install and configure the channel tab](../teams-install.html).

## Important correction: full contact sync needs a backend

The current permission-free matching/additive refresh is not full two-way tag membership synchronization. Microsoft Graph documents **adding a member to an existing Teams tag as unsupported with delegated work/school permissions**. That operation requires the **TeamworkTag.ReadWrite.All application permission**. Delegated TeamworkTag.ReadWrite supports other tag operations, including documented removal, but does not by itself enable the required additions.

Therefore a revised Outlook manifest, an additional delegated-consent checkbox, or the Teams ZIP alone cannot enable full synchronization. There is no client credential in the browser package, and no application permission or tenant policy has been changed.

The full-sync design needs the following separate review before implementation/deployment:

| Component | Required change |
| --- | --- |
| Router identity matching | Delegated TeamMember.Read.All, scope ID `2497278c-d82d-46a2-b1ce-39d4cdde5570`, to resolve current selected-Team members to unique email/user IDs. |
| Secure backend identity | A server-side service identity with admin-approved **application** TeamworkTag.ReadWrite.All, role ID `a3371ca5-911d-46d6-901c-42c8c7a937d8`, for membership additions to existing tags. This is separate from delegated TeamworkTag.ReadWrite. |
| Backend authentication | An authenticated API and explicit coordinator authorization, with configured Team/tag allowlists. Its app registration/API scope and hosting details must be finalized; a complete backend deployment manifest does not exist yet. |
| Credential handling | Managed identity or server-held certificate/credential in approved secure hosting. No client secret or application token may be included in the Teams ZIP, Outlook scripts, connection code or public repository. |

Microsoft Graph resource app ID is `00000003-0000-0000-c000-000000000000`. Preserve all existing permissions; the accompanying permission-request.json is a **review request**, not a replacement Entra manifest and not an authorization to grant application access. The publisher/registration owner handles registration changes as needed; the organization's authorized administrator approves tenant access. Do not grant an application permission to a browser-only SPA and assume it can then perform app-only operations.

The membership-comparison and preview prototype is preserved privately, but it is **not deployed or enabled**. No live tenant membership or consent tests have been performed. The production router retains its existing matching/additive refresh; the staff queue remains usable independently of full sync.

## Intended full-sync behavior for review

- Broad division contact rosters use a dedicated contact tag, separate from coordinator notification tags and workflow roles.
- An initial reviewed comparison combines additions without deleting existing people. Subsequent comparisons use the last verified common membership to propagate additions/removals from either side.
- Show exact additions/removals and require confirmation before applying. Ambiguous identities block the plan. Team owners still handle joining the Team; automatic Team-membership writes are not proposed.
- Recheck source versions, protect simultaneous edits, and advance the baseline only after both memberships converge. Preserve failed/partial outcomes for review; there is no cross-service transaction.
- A short routing Emails column may need conversion to plain multiline text by its list owner for larger rosters. No list-type changes have been made.
- No email, real mentions or Teams posts occur as a side effect of contact sync.

## Primary Microsoft references

- [Custom-app policy and administrator publishing](https://learn.microsoft.com/en-us/microsoftteams/teams-custom-app-policies-and-settings)
- [Create a tag member: application-only permission](https://learn.microsoft.com/en-us/graph/api/teamworktagmember-post?view=graph-rest-1.0)
- [Remove a tag member](https://learn.microsoft.com/en-us/graph/api/teamworktagmember-delete?view=graph-rest-1.0)
- [List Team members](https://learn.microsoft.com/en-us/graph/api/team-list-members?view=graph-rest-1.0)
- [Permission IDs and consent requirements](https://learn.microsoft.com/en-us/graph/permissions-reference)
