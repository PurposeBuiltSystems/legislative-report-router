# Required Microsoft Graph permissions

All permissions are **delegated** — the add-in acts as the signed-in
coordinator and can never reach anything the coordinator can't. There are
no application permissions and no backend.

## Requested at sign-in

| Scope | Used for | Why not less |
| --- | --- | --- |
| `Mail.ReadWrite` | Read the report message body/attachments; save the compose draft; draft the daily report and setup invitations | Reading a draft's attachments requires more than `Mail.Read` |
| `Mail.Send` | Send consolidated division emails; send the original report on publish | Sending is the product's job; drafts-only would break the workflow |
| `ChannelMessage.Send` | Post one message per bill to Teams channels | The narrowest channel-posting scope |
| `ChannelMessage.Read.All` | Read replies **to the add-in's own bill posts**, so divisions' cost estimates can be collected without re-typing (History → Harvest) | Graph has no "replies to messages I posted" scope; this is the narrowest available that can read a thread |
| `Sites.ReadWrite.All` | Read `LegislativeRoutingMatrix`; write `LegislativeAudit` and `BillTracker` rows | Graph has no list-scoped delegated permission; see hardening below |
| `Team.ReadBasic.All` | List the teams you belong to, for the setup wizard's Team picker | Required for `/me/joinedTeams` |
| `Channel.ReadBasic.All` | List a team's channels in the setup wizard | Required for `/teams/{id}/channels` |
| `TeamworkTag.Read` | Read a team's tags so a division's tag can be selected | Read-only |

## Requested only when you use the feature (incremental consent)

These are **not** requested at sign-in. The add-in asks for them at the
moment the action is taken, so a tenant that never grants them still gets
a fully working add-in everywhere else.

| Scope | Requested when | If refused |
| --- | --- | --- |
| `Sites.Manage.All` | You click **Create my lists** during setup | Create the three lists by hand in Microsoft Lists (schemas in the admin guide); everything else works |
| `TeamworkTag.ReadWrite` | You click **➕ Create the tag for this division** | Create the tag in Teams (Manage team → Tags) and pick it from the dropdown |

> Why `Sites.Manage.All` is needed at all: Microsoft splits these precisely —
> `Sites.ReadWrite.All` is *"edit or delete documents and list items"*, while
> creating a **list** requires `Sites.Manage.All`, *"create or delete document
> libraries and lists"*. The add-in only ever uses it to provision its own
> three lists.

## Hardening option

If security review balks at `Sites.ReadWrite.All` (it is delegated —
user's-reach-only — but broad), switch the app registration to
**`Sites.Selected`** and have the SharePoint admin grant the app access to
only the legislative site. No code changes required. Note that list
*creation* still needs `Sites.Manage.All`, so provision the three lists by
hand first if you take this route.

## Admin consent

Gov tenants typically block user consent. One-time admin consent URL:

```
https://login.microsoftonline.com/organizations/adminconsent?client_id=0860a653-ddbd-4455-8bff-affda2a8879f
```

(Use the `.us` authority host for GCC High/DoD.)

## What is never requested

No **application** permissions of any kind, no `Directory.*`, no
`User.Read.All`, no `Mail.Read.Shared`. The add-in cannot read other
people's mailboxes, the directory, or any Teams conversation it did not
itself post into — `ChannelMessage.Read.All` is used solely to follow
replies to the add-in's own bill posts, and only in channels the
signed-in coordinator already belongs to.

## Audience

The app registration is **AzureADMultipleOrgs** — work or school accounts
only. Personal Microsoft accounts are not supported, because the Teams and
SharePoint scopes above are not available to them.

## Shared Teams and email contacts — no additional permission

Step 4 matches Teams tag user IDs to user-supplied email addresses. Name/email/division paste suggests choices, but only confirmed matches are saved. The matching form reads the Team's tags and tag members using existing TeamworkTag.Read; it never requests TeamMember.Read.All, directory access, or application permissions.

Confirmed matches are stored as contact records in the existing routing list, using existing columns: Notes = LRR_CONTACT_MATCH_V1, TeamsTeamId, MentionUserIds, MentionUserEmails, DivisionName (person's name), and IsActive = false. These rows are excluded from division routing. One record per person avoids putting a whole division's map in a short text field. Other coordinators reuse the same matches.

Tag emails are added to the division's Emails column. Contacts outside its tag use direct user mentions. Email contacts without a confirmed Teams identity remain email-only and are shown as exceptions. New tag members without an email match block shared-audience publishing until matched. If a person is absent from the form, add them to a tag in the Team in Teams, then reload. The add-in does not create Team membership or change existing tag membership.

Refresh happens in the open add-in on demand and before Preview; Publish uses the reviewed audience. Reconciliation is additive. To remove a recipient, remove their email from the route and their division tag membership; to forget a saved identity/email match, delete its Contact row in the routing list. Existing list-column length limits still apply to the Emails column.

No new Outlook manifest, new Graph scope, schema provisioning, or additional admin consent is required for this matching path. Existing permissions and access must already be granted.

## Budget dashboard in Teams

The in-router dashboard filters BillTracker data by fiscal year and division. Publishing creates or updates a native SharePoint site page using the existing Sites.ReadWrite.All delegated permission; no new Graph scope, Teams app manifest or Power BI license is requested. Site/page editing rights and any existing page approval process still apply.

Load and review the fiscal rollup, then choose Publish budget snapshot to SharePoint. The published snapshot includes every loaded fiscal year and division and displays the data-load timestamp. In Teams, add a SharePoint tab and select the generated page. Refresh the rollup and republish to update it; the Teams snapshot does not refresh itself when tracker rows change.

Unknown estimates remain separate from zero-cost estimates. Savings, spending and net estimates are shown separately; totals include all tracked bills and carry status/duplicate-row warnings. Detail is limited to 150 displayed rows with a link to the full editable BillTracker.

## Daily report distribution from division contacts

Step 5 offers Use all division contacts for the daily report. It combines the email contacts of every currently active division route, excludes expired/future/inactive rules and Contact metadata records, and removes duplicate addresses without regard to case. The existing distribution field adds extra recipients when this option is enabled; when disabled, it remains a separate manual distribution list.

Setup shows the resolved recipient list. Preview reloads division contacts and requires a new confirmation before creating the daily report draft. The morning draft schedule uses the same recipient calculation and refreshes division/shared contacts first. This changes draft recipients only; it does not send the daily report automatically or change targeted per-division email behavior. No new permission or manifest is needed. Existing saved distribution behavior remains the default until the option is enabled.
