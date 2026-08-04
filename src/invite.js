/*
 * Legislative Report Router — setup by invitation (pure).
 *
 * Handing someone a base64 blob and telling them where to paste it is the
 * part a coordinator has to explain and a recipient can get wrong. Instead
 * the coordinator addresses a person, and that person's own mailbox carries
 * the setup to them.
 *
 * Nothing secret travels here: a site URL, list names, watch terms, and the
 * routing preferences already visible to anyone on the team.
 */
(function (root) {
  "use strict";

  var SUBJECT = "Legislative Report Router setup";
  var OPEN = "[[LRR-SETUP]]";
  var CLOSE = "[[/LRR-SETUP]]";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function setupSubject(orgLabel) {
    var org = String(orgLabel || "").trim();
    return SUBJECT + (org ? " — " + org : "");
  }

  /**
   * The invitation. Written for a person first; the configuration is a
   * footnote they never need to touch.
   */
  function inviteHtml(opts) {
    opts = opts || {};
    var who = esc(opts.fromName || "Your legislative coordinator");
    var site = esc(opts.siteLabel || "the team's SharePoint site");
    return '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px;color:#242424">' +
      "<p>You've been set up to run the <strong>Legislative Report Router</strong> — the " +
      "Outlook add-in that turns a Daily Bill Report into Teams posts, division emails and the " +
      "shared bill tracker.</p>" +
      "<p><strong>To finish: open any email, click Report Router on the ribbon, open Setup and " +
      "choose &ldquo;Find my setup&rdquo;.</strong> It reads this message and fills in the site, " +
      "lists and preferences — nothing to copy or paste. Then click " +
      "<strong>Connect &amp; load routing rules</strong> once to sign in on your device.</p>" +
      "<p style=\"color:#616161;font-size:12px\">Lists live on " + site + ". Keep this message; " +
      "the add-in looks for it. The settings below are configuration only — no passwords.</p>" +
      '<div style="background:#f3f2f1;border-radius:6px;padding:8px;font-family:ui-monospace,monospace;' +
      'font-size:11px;word-break:break-all;color:#616161">' +
      OPEN + esc(opts.code || "") + CLOSE + "</div>" +
      "<p style=\"color:#616161;font-size:12px\">Sent by " + who +
      (opts.fromEmail ? " (" + esc(opts.fromEmail) + ")" : "") + ".</p>" +
      "</div>";
  }

  /** Recover the code, tolerating what mail clients do to a body. */
  function extractSetupCode(body) {
    var text = String(body || "");
    var i = text.indexOf(OPEN);
    var j = text.indexOf(CLOSE, i + 1);
    if (i === -1 || j === -1) { return ""; }
    return text.slice(i + OPEN.length, j)
      .replace(/<[^>]*>/g, "")
      .replace(/&amp;/g, "&")
      .replace(/[^A-Za-z0-9+/=]/g, "");
  }

  /** Newest message that actually carries a code. */
  function pickInvite(messages) {
    var usable = (messages || []).filter(function (m) {
      return extractSetupCode(m.body || m.bodyPreview || "");
    });
    usable.sort(function (a, b) {
      return String(b.receivedDateTime || "").localeCompare(String(a.receivedDateTime || ""));
    });
    return usable[0] || null;
  }

  var api = {
    SUBJECT: SUBJECT,
    setupSubject: setupSubject,
    inviteHtml: inviteHtml,
    extractSetupCode: extractSetupCode,
    pickInvite: pickInvite,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrInvite = api; }
})(typeof self !== "undefined" ? self : this);
