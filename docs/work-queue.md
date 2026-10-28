# Work queue

Open **Work queue → Load shared work** after connecting the existing SharePoint lists in Setup. A request owner creates separate Comment and Fiscal data requests, using the same bill, General Assembly and session identifiers and bill conversation link. Give each reassessment its own request label. Nothing in this screen sends Teams messages, tags, email or submits to LSA.

The owner records the confirmed deadline and its source, related bill references, required divisions, and each division's coordinator, director and reviewers. Use the participants' **signed-in usernames**. The broad daily-report distribution list is independent of these roles. The owner can revise assignments; this resets approvals. No names or assignments are inferred from examples or imported tags.

1. An assigned reviewer or coordinator registers a draft response with a version label and HTTPS links to the exact document version(s). For fiscal responses, keep the written form and workbook together in that response package. Store documents in SharePoint/Teams as usual; this screen registers links and does not upload files. It does not monitor remote file edits: register every revised response before further review, and use version-specific links.
2. The reviewer requests review. The coordinator consolidates that response and the assigned director approves it. A request for changes must be followed by a registered revised response before review can resume.
3. Every required division must have current director approval before the assigned CFO can approve a fiscal request. Changes to a response, required divisions or assignments reset the affected approvals and CFO readiness. Earlier approvals remain in history.
4. After actually submitting the approved documents through the normal LSA process, the designated submitter records its actual time, document description and evidence link. This is a submission record, not an automated portal submission.

The queue filters by division, request type, overall or division stage, and approaching/overdue or missing deadlines. It shows the next responsible person. Due indicators compare the confirmed calendar date with the viewer's local date; they are visible reminders, not scheduled notifications. An unspecified date stays “Deadline needs confirmation.” The supplied LSA form's five-day wording does not establish its request date or holiday convention.

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

The interactive queue currently runs in the Outlook router. It does not yet provide a standalone Teams staff interface, automatic reminders, automatic approval extraction from old conversations, or automatic Teams parent-thread publication. Existing Teams document/conversation links remain available. No live tenant messages or email are used by the automated tests.
