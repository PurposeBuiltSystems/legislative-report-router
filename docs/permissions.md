# Required Microsoft Graph permissions

All permissions are **delegated** — the add-in acts as the signed-in
coordinator and can never reach anything the coordinator can't. There are
no application permissions and no backend.

## Required

| Scope | Used for | Why not less |
| --- | --- | --- |
| `Mail.ReadWrite` | Read the report message body/attachments; save the compose draft | Reading a draft's attachments requires more than `Mail.Read` |
| `Mail.Send` | Send consolidated division emails; send the original report on publish | Sending is the product's job; drafts-only would break the workflow |
| `ChannelMessage.Send` | Post one message per bill to Teams channels | The narrowest channel-posting scope |
| `Sites.ReadWrite.All` | Read `LegislativeRoutingMatrix`; write `LegislativeAudit` | Graph has no list-scoped delegated permission; see hardening below |

## Optional

| Scope | Used for |
| --- | --- |
| `TeamworkTag.Read` | The Settings "tag ID lookup" helper (find a tag's GUID by team) |
| `Team.ReadBasic.All`, `Channel.ReadBasic.All` | Future rule-validation helpers |

## Hardening option

If security review balks at `Sites.ReadWrite.All` (it is delegated —
user's-reach-only — but broad), switch the app registration to
**`Sites.Selected`** and have the SharePoint admin grant the app access to
only the legislative site. No code changes required.

## Admin consent

Gov tenants typically block user consent. One-time admin consent URL:

```
https://login.microsoftonline.com/organizations/adminconsent?client_id=0860a653-ddbd-4455-8bff-affda2a8879f
```

(Use the `.us` authority host for GCC High/DoD.)

## What is never requested

No application permissions, no `Directory.*`, no `User.Read.All`, no
`ChannelMessage.Read.All` — the add-in cannot read Teams messages, other
mailboxes, or the directory.
