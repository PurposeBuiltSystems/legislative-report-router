/*
 * Legislative Report Router — Teams message builder (pure logic).
 *
 * Builds Microsoft Graph chatMessage payloads with REAL mention entities:
 * <at id="N">…</at> markers in the HTML body paired 1:1 with entries in the
 * mentions[] array (mentioned.tag for Teams tags, mentioned.user for
 * people). All user content is HTML-escaped; only the <at> markers and
 * template markup are raw HTML. Never emits escaped tags like &lt;at&gt;.
 *
 * Also: idempotency keys and the email bodies for consolidated
 * per-division targeted mail. Templates live here, separate from wiring.
 */
(function (root) {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function nl2br(s) { return esc(s).replace(/\n/g, "<br>"); }

  /** Default per-bill Teams post template. Placeholders are functions of (item, rule). */
  var DEFAULT_TEMPLATE = {
    commentWindow: "Review comments requested within 48 business hours.",
  };

  /**
   * Build mention entities for a rule set: one tag mention per distinct
   * tagId, one user mention per distinct userId. Returns
   * {atHtml, mentions} where atHtml is the "<at…> <at…>" prefix.
   */
  function buildMentions(rules) {
    var mentions = [];
    var seen = {};
    var parts = [];
    (rules || []).forEach(function (r) {
      if (r.teamsTagId && !seen["tag:" + r.teamsTagId]) {
        seen["tag:" + r.teamsTagId] = true;
        var idx = mentions.length;
        var name = r.teamsTagName || r.divisionCode || "Team";
        mentions.push({
          id: idx,
          mentionText: name,
          mentioned: { tag: { id: r.teamsTagId, displayName: name } },
        });
        parts.push('<at id="' + idx + '">' + esc(name) + "</at>");
      }
      (r.mentionUserIds || []).forEach(function (uid, k) {
        if (!uid || seen["user:" + uid]) { return; }
        seen["user:" + uid] = true;
        var uidx = mentions.length;
        var display = (r.mentionUserEmails || [])[k] || r.divisionName || "Member";
        mentions.push({
          id: uidx,
          mentionText: display,
          mentioned: { user: { id: uid, displayName: display, userIdentityType: "aadUser" } },
        });
        parts.push('<at id="' + uidx + '">' + esc(display) + "</at>");
      });
    });
    return { atHtml: parts.join(" "), mentions: mentions };
  }

  /**
   * One Graph chatMessage payload for one legislative item posted to one
   * channel. `rulesForChannel` = all matched rules that share this channel
   * (their tags/users are all mentioned in this single post — deduped).
   */
  function buildChannelMessage(item, rulesForChannel, opts) {
    opts = opts || {};
    var t = opts.template || DEFAULT_TEMPLATE;
    var m = buildMentions(rulesForChannel);
    var linkHtml = (item.sourceLinks || []).slice(0, 1).map(function (l) {
      return '<br>Open bill: <a href="' + esc(l.href) + '">' + esc(l.text || l.href) + "</a>";
    }).join("");

    var content =
      (m.atHtml ? m.atHtml + "<br><br>" : "") +
      "<strong>" + esc(item.billNumber) + "</strong>" +
      (item.title && item.title !== item.brief.split("\n")[0] ? " — " + esc(item.title) : "") +
      "<br><br>Distributed To: " + esc((item.distributedTo || []).join(", ") || "—") +
      "<br>Comment Requested From: " + esc((item.commentRequestedFrom || []).join(", ") || "—") +
      "<br><br><strong>Brief:</strong><br>" + nl2br(item.brief || "(none)") +
      (t.commentWindow ? "<br><br>" + esc(t.commentWindow) : "") +
      linkHtml;

    return {
      body: { contentType: "html", content: content },
      mentions: m.mentions,
    };
  }

  /** Consolidated targeted-email HTML for one division rule and its items. */
  function buildDivisionEmail(rule, items, opts) {
    opts = opts || {};
    var t = opts.template || DEFAULT_TEMPLATE;
    var rows = (items || []).map(function (it) {
      return "<h3 style=\"margin:14px 0 2px\">" + esc(it.billNumber) + "</h3>" +
        "<p style=\"margin:2px 0;color:#616161\">Distributed To: " + esc((it.distributedTo || []).join(", ")) +
        " · Comment Requested From: " + esc((it.commentRequestedFrom || []).join(", ")) + "</p>" +
        "<p style=\"margin:4px 0\">" + nl2br(it.brief || "") + "</p>";
    }).join("");
    return {
      subject: (opts.subjectPrefix || "Legislative bills for review") + " — " +
        (rule.divisionName || rule.divisionCode) + " (" + (items || []).length + ")",
      html: '<div style="font-family:Segoe UI,Arial,sans-serif;max-width:640px">' +
        "<p>The following bills from today's report are assigned to <strong>" +
        esc(rule.divisionName || rule.divisionCode) + "</strong>:</p>" + rows +
        (t.commentWindow ? "<p><strong>" + esc(t.commentWindow) + "</strong></p>" : "") +
        "</div>",
      to: rule.emails || [],
    };
  }

  /**
   * Deterministic idempotency key: same report + bill + channel can only
   * publish once. (djb2 — collision-safe enough for a per-report audit list.)
   */
  function idempotencyKey(reportKey, billNumber, channelId) {
    var s = [reportKey, billNumber, channelId].join("|");
    var h = 5381;
    for (var i = 0; i < s.length; i++) { h = ((h << 5) + h + s.charCodeAt(i)) >>> 0; }
    return "lrr-" + h.toString(36) + "-" + s.length.toString(36);
  }

  var api = {
    buildMentions: buildMentions,
    buildChannelMessage: buildChannelMessage,
    buildDivisionEmail: buildDivisionEmail,
    idempotencyKey: idempotencyKey,
    DEFAULT_TEMPLATE: DEFAULT_TEMPLATE,
    _internals: { esc: esc },
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrTeams = api; }
})(typeof self !== "undefined" ? self : this);
