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
