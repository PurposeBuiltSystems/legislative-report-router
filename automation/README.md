# Scheduled legislative reminders

The router projects current assignments into its existing LegislativeAudit list. This optional cloud flow reads that feed and sends one consolidated Outlook email per responsible person per weekday at 8 a.m. Central time (including daylight saving time). It includes confirmed deadlines within two calendar days and overdue work. Comments and fiscal requests keep their separate review lifecycles. No Teams or Outlook manifest change is required by this flow.

## Install in another organization

1. In Power Automate, create a temporary scheduled cloud flow with SharePoint **Get items** bound to the router's LegislativeAudit list and Office 365 Outlook **Send an email (V2)** addressed only to the installer. Set a future start date. Save, then turn the flow off without running it.
2. Export that flow using **Export → Package (.zip)**. The export supplies Microsoft's actual package and connection references; no credentials are placed in this repository.
3. Run:

   ```sh
   python3 automation/package-reminders.py exported.zip reminders.zip \
     --site https://YOUR-TENANT.sharepoint.com/sites/YOUR-SITE \
     --audit-list YOUR-LEGISLATIVE-AUDIT-LIST-GUID
   ```

4. Use **My flows → Import → Import Package (Legacy)**. Choose **Update** for the temporary flow and select the existing SharePoint and Office 365 Outlook connections. Import and open the flow. Review the recurrence timezone/start date; change them for your organization. Keep exactly one enabled reminder flow per shared audit list.
5. Run once before connecting the feed: it should succeed without sending. In the router, load **Work queue**, then select **Connect automatic reminders**. Test with a clearly labeled self-only assignment, check the actual email and run history, and run again to verify deduplication. Close the test assignment through its ordinary review steps. Enable the schedule after verification.

The standard SharePoint and Office 365 Outlook connectors run as the connected Microsoft 365 account. That account needs access to the shared audit list and a mailbox it can send from. Organization connector, consent, licensing, and DLP policies still apply. Install the flow separately in each organization's own Power Automate environment; the Teams app package does not install cloud flows.

## What prevents stale or repeated reminders

Before including each assignment, the flow rereads its current row and newest committed workflow-event ID. Only matching revisions that remain pending and due are included. The SharePoint reservation uses that row's ETag; a concurrent change fails the reservation. Text and link attributes are HTML escaped. The flow rereads the site control before sending. There is a small unavoidable interval between these checks and mail delivery across the separate services.

Trigger and recipient loops run sequentially. A recipient/day attempt record is written before sending, and Outlook send retries are disabled. This favors avoiding duplicates after an ambiguous send result. If a send fails, inspect the run and mailbox before retrying: clear only the verified failed recipient/day attempt and its assignment `PublishedBy` dates to permit a same-day retry. Do not blindly resubmit a failed run. A failed reservation prevents that recipient's email; previously reserved rows may also need recovery.

Use **Pause automatic reminders** in the router to disable the site's feed, or **Turn off** in Power Automate to stop its schedule. Draft reminders remain a separate manual action. Router edits refresh the feed after committing the workflow; a failed refresh is shown in the router, while the revision comparison prevents the older feed from being sent. Reconnect the feed to refresh it after resolving an error.

## Existing list column mapping

These internal `reminder-*` rows are excluded from publication history. Do not edit them as routing contacts or published bills.

| Column | Reminder meaning |
| --- | --- |
| Status | `reminder-control`, `reminder-pending`, `reminder-closed`, or `reminder-attempt` |
| ReportKey | Request ID; for an attempt, `reminder:YYYY-MM-DD:recipient` |
| IdempotencyKey | Latest committed workflow row ID |
| Divisions | Assignment division or All divisions |
| EmailRecipients | Current responsible signed-in email |
| Error | Confirmed due date; control row uses enabled/disabled |
| SourceSubject | Request label |
| TeamId | Next action |
| ChannelId + TeamsMessageId | Teams/source link split across short columns |
| PublishedBy | Daily attempt reservation; router sync preserves it |

The definition requests pagination up to 5,000 pending rows, the verified limit of the test environment's plan. Larger queues need a reviewed pagination design and suitable SharePoint indexes before use. Retain attempt records according to organizational retention rules; flow run history alone is short-lived. Changes to the flow can be made later in Power Automate; keep this portable definition aligned with the tested version.
