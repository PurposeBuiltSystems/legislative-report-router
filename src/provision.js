/*
 * Legislative Report Router — SharePoint list provisioning (pure defs).
 *
 * One-click setup: the add-in creates the three lists it needs, with the
 * exact column schemas from docs/admin-guide.md, via Graph. Admins can
 * still create them by hand — this just removes the transcription work.
 */
(function (root) {
  "use strict";

  function text(name) { return { name: name, text: {} }; }

  /** Graph list definitions keyed by role. displayName comes from Settings. */
  function listDefinitions() {
    return {
      routing: {
        description: "Division routing rules for Legislative Report Router (edit in Microsoft Lists).",
        columns: [
          text("DivisionCode"), text("DivisionName"), text("Aliases"),
          text("Emails"), text("TeamsTeamId"), text("TeamsChannelId"),
          text("TeamsChannelName"), text("TeamsTagId"), text("TeamsTagName"),
          text("CodeChapters"), text("MentionUserIds"), text("MentionUserEmails"),
          { name: "IsActive", boolean: {} },
          { name: "Priority", number: {} },
          { name: "EffectiveStartDate", dateTime: {} },
          { name: "EffectiveEndDate", dateTime: {} },
          text("Notes"),
        ],
      },
      audit: {
        description: "Publication audit trail + duplicate-prevention record for Legislative Report Router. Do not delete rows.",
        columns: [
          text("ReportKey"), text("IdempotencyKey"), text("TeamId"),
          text("ChannelId"), text("TeamsMessageId"), text("Status"),
          text("Error"), text("Divisions"), text("EmailRecipients"),
          text("PublishedBy"), text("SourceSubject"),
        ],
      },
      tracker: {
        description: "Per-division bill review status. Pin as a Lists tab in the legislative Teams channel.",
        columns: [
          text("Division"),
          { name: "Status", choice: { displayAs: "dropDownMenu",
            choices: ["Pending review", "In review", "Commented", "No comment needed"] } },
          { name: "DueDate", dateTime: {} },
          text("BillLink"), text("Brief"), text("ReportKey"),
        ],
      },
    };
  }

  /** A starter routing rule so the list opens with visible structure. */
  function sampleRoutingRule() {
    return {
      Title: "MVD",
      DivisionCode: "MVD",
      DivisionName: "Motor Vehicle Division (SAMPLE - edit me)",
      Aliases: "Motor Vehicle; Motor Vehicles",
      CodeChapters: "321; 321A; 322",
      IsActive: true,
      Priority: 1,
      Notes: "Sample rule created by the add-in. Fill in Teams IDs (Settings > tag lookup) and emails, or delete.",
    };
  }

  var api = { listDefinitions: listDefinitions, sampleRoutingRule: sampleRoutingRule };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrProvision = api; }
})(typeof self !== "undefined" ? self : this);
