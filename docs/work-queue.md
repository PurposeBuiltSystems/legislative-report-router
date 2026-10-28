# Work queue

Open **Work queue → Load shared work** after connecting the existing SharePoint lists in Setup. A request owner creates separate Comment and Fiscal data requests, using the same bill, General Assembly and session identifiers and bill conversation link. Give each reassessment its own request label. Nothing in this screen sends Teams messages, tags, email or submits to LSA.

The owner records the confirmed deadline and its source, related bill references, required divisions, and each division's coordinator, director and reviewers. Use the participants' **signed-in usernames**. The broad daily-report distribution list is independent of these roles. The owner can revise assignments; this resets approvals. No names or assignments are inferred from examples or imported tags.

1. An assigned reviewer or coordinator registers a draft response with a version label and HTTPS links to the exact document version(s). For fiscal responses, keep the written form and workbook together in that response package. Store documents in SharePoint/Teams as usual; this screen registers links and does not upload files. It does not monitor remote file edits: register every revised response before further review, and use version-specific links.
2. The reviewer requests review. The coordinator consolidates that response and the assigned director approves it. A request for changes must be followed by a registered revised response before review can resume.
3. Every required division must have current director approval before the assigned CFO can approve a fiscal request. Changes to a response, required divisions or assignments reset the affected approvals and CFO readiness. Earlier approvals remain in history.
4. After actually submitting the approved documents through the normal LSA process, the designated submitter records its actual time, document description and evidence link. This is a submission record, not an automated portal submission.

The queue filters by session, division, request type, overall or division stage, approaching/overdue or missing deadlines, and My assigned work. Filters work together. It shows the next responsible person. Due indicators compare the confirmed calendar date with the viewer's local date; they are visible reminders, not scheduled notifications. An unspecified date stays “Deadline needs confirmation.” The supplied LSA form's five-day wording does not establish its request date or holiday convention.

Comments and fiscal data have independent requests and statuses. Legislative progress is a separate, source-linked event shared across requests with the same bill, assembly and session. Recording submission does not stop legislative tracking. Related and successor bill references do not merge records. Link amendments to their parent bill and identify the amendment in the request label. This feature does not change the existing publication path to create Teams thread replies automatically.

FYI, consultation and approval-request notes retain intent without changing roles or granting approval. Filenames and text such as “final” or “Good to go” do not become workflow approvals automatically. Historical examples can be recorded as notes with verified source context; current approval requires the assigned person's explicit action.

Fiscal response detail can be kept in linked official forms. Optional narrative fields cover summary, background, assumptions, fiscal/operational impact and related sources. Optional financial lines take seven tab-separated fields:

```
FY2027	Revenue	Increase	1000	General Fund	Ongoing	Fees
FY2027	Expenditure	Increase	1000	General Fund	One-time	Implementation
FY2028	Expenditure	Increase	Unknown	Other fund	Ongoing	Staffing
```

Use nonnegative dollar amounts without commas, or `Unknown`, and explicitly choose Increase/Decrease. Blank or missing lines stay incomplete. Template zeros are not confirmed estimates. Totals keep revenue/expenditure increases/decreases separate and expose unknown lines. These structured figures are separate from the legacy BillTracker budget snapshot and are not automatically merged into it. The written response remains necessary.

## Storage, permissions and recovery

This feature uses the existing audit list's item fields and existing delegated Sites.ReadWrite.All permission. No manifest, Entra scope, list or column creation is needed. Workflow data is stored as append-only parts and a final checksum-verified commit. Partial writes never become approvals. Do not delete or edit `workflow-part` or `workflow-event` records in the audit list. The Work queue displays their readable history separately from publication history.

The acting username comes from the MSAL result for the token used for the save, rather than the Outlook mailbox or the first cached account. Workflow role checks govern this UI; SharePoint list access remains the underlying security boundary. This is not a server-enforced approval system for people who can directly edit the audit list. Retain appropriate SharePoint access and auditing.

Every save rereads shared history and rejects a stale view. Concurrent committed actions trigger an owner-review state that blocks approvals and submission. The owner reviews the retained history, then saves reviewed details to resolve the conflict; approvals reset. If a save times out or returns an uncertain result, load shared work and inspect history before retrying. If committed data is missing or corrupt, writes are blocked rather than showing an incomplete approval chain. Full-history paging limits produce an error, not a partial queue.

The same interactive queue is available in the Outlook router and in a standalone Teams channel tab. See [Teams setup](../teams-install.html). In the router, Work queue -> Teams setup creates a connection code containing the existing site/list IDs and default filters, with no tokens or workflow data. The channel owner pastes it once when configuring the Teams tab. The tab reads and writes the same audit records; it does not duplicate the workflow store. Each staff member signs in and uses their existing SharePoint access and assigned workflow role, without the Outlook add-in.

The Teams app package is separate from the Outlook XML manifest. It reuses the existing Entra client ID and broker domain, requesting only the existing Sites.ReadWrite.All delegated scope. TeamsJS initializes before MSAL. Tenant policies may require an administrator to allow or distribute the custom Teams app. This package does not change tenant policies, grant permissions or install itself.

[Practice mode](../src/staff/index.html?demo=1) supplies fictional requests, roles and dates in page memory; no Graph calls, sign-in, messages or notifications. It is clearly labeled and resets on reload. No sample records are inserted into the shared audit list.

The live queue is loaded/refreshed explicitly and shows its load time. Filter changes retain open response edits, and refreshing/navigating while editing requires saving or discarding the edits. It does not provide automatic reminders, automatic approval extraction from old conversations or automatic Teams parent-thread publication. Existing Teams document/conversation links remain available. No live tenant messages or email are used by the automated tests.


## Share the contact list through the built-in Teams Lists app

The router's **Open your SharePoint lists → Set up the same contact list in Teams** guide provides the connected routing/contact list's canonical SharePoint URL, a Copy button with a selectable-text fallback, and a chooser for existing routing channels. Division routes sharing a channel are grouped. The router retrieves the selected channel's canonical Teams link using existing Channel.ReadBasic.All access; it does not create a tab, post a message or install an app.

Copy the list URL, open the selected channel, and complete **+ → Lists → Add an existing list** in Teams. This uses the same list the router reads, with existing SharePoint access, without the custom Teams package or additional Graph permissions. It does not synchronize native tag membership. For list-only roster management, turn off the legacy shared-contacts option that adds contacts from tags. Reconnect the router after edits to load the current list.

The final Add/Save action remains in Teams. Microsoft Graph tab creation needs extra permission, and Microsoft documents SharePoint page/list tab configuration as unsupported through Graph. Source: https://learn.microsoft.com/en-us/graph/teams-configuring-builtin-tabs .
