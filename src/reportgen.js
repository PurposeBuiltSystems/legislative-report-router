/*
 * Legislative Report Router — Daily Bill Report generator (pure).
 *
 * The reverse gear: instead of only parsing a hand-written report, compose
 * one FROM the newly-filed feed. Bills arrive from the Filings tab with
 * divisions assigned (auto-suggested from Code chapters, corrected in
 * Review); this module renders them into the coordinator's familiar
 * Daily Bill Report format — as email HTML and plain text.
 *
 * The generated text deliberately round-trips through our own parser
 * (bill numbers standalone on their own lines), so a recipient at another
 * agency could feed it straight back into their own Router.
 */
(function (root) {
  "use strict";

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function pad(n) { return (n < 10 ? "0" : "") + n; }

  /** "7-26-2026 Daily Bill Report" — matches the coordinator's convention. */
  function subjectFor(date) {
    var d = date instanceof Date ? date : new Date(date);
    return (d.getMonth() + 1) + "-" + d.getDate() + "-" + d.getFullYear() + " Daily Bill Report";
  }

  /**
   * Pull email addresses out of anything: a pasted Outlook To: line
   * ("Doe, Jane <jane@x.gov>; Roe, Bob <bob@x.gov>"), a semicolon list, a
   * spreadsheet column. Dedupes, lowercases.
   */
  function extractEmails(text) {
    var out = [];
    var re = /[A-Za-z0-9._%+'-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}/g;
    var m;
    while ((m = re.exec(String(text || ""))) !== null) {
      var e = m[0].toLowerCase();
      if (out.indexOf(e) === -1) { out.push(e); }
    }
    return out;
  }

  /** Drop feed entries whose bill number was already in a past report. */
  function filterNew(entries, reportedBills) {
    var seen = {};
    (reportedBills || []).forEach(function (b) { seen[String(b).toUpperCase()] = true; });
    return (entries || []).filter(function (e) { return !seen[String(e.bill).toUpperCase()]; });
  }

  /**
   * Render included items into the Daily Bill Report.
   * opts: {sessionName, commentWindow, date, preparedBy}
   * Returns {subject, html, text}.
   */
  function buildDailyReport(items, opts) {
    opts = opts || {};
    var date = opts.date instanceof Date ? opts.date : new Date();
    var included = (items || []).filter(function (it) { return it.routingStatus !== "excluded"; });

    var textLines = [];
    textLines.push("Daily Bill Report");
    if (opts.sessionName) { textLines.push(opts.sessionName); }
    textLines.push("");
    if (opts.commentWindow) { textLines.push(opts.commentWindow); textLines.push(""); }
    textLines.push("Bill Number:");
    textLines.push("Bill Distributed To:");
    textLines.push("Bill Comment Requested From:");
    textLines.push("Title:");
    textLines.push("");

    var h = [];
    h.push('<div style="font-family:Segoe UI,Arial,sans-serif;max-width:680px;color:#242424">');
    h.push('<h2 style="margin-bottom:2px">Daily Bill Report</h2>');
    if (opts.sessionName) { h.push('<p style="margin:0 0 12px;color:#616161">' + esc(opts.sessionName) + "</p>"); }
    h.push('<p style="margin:0 0 4px;color:#616161">' + esc(
      date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate())) +
      " · " + included.length + " bill(s)/amendment(s)</p>");
    if (opts.commentWindow) {
      h.push('<p style="background:#eff6fc;border-radius:6px;padding:8px 12px">' + esc(opts.commentWindow) + "</p>");
    }

    included.forEach(function (it) {
      var link = (it.sourceLinks || [])[0];
      var dist = (it.distributedTo || []).join(", ") || "TBD";
      var from = (it.commentRequestedFrom || []).join(", ") || dist;

      textLines.push(it.billNumber);
      textLines.push("");
      textLines.push(dist);
      textLines.push(from);
      textLines.push(it.brief || "(no summary)");
      if (link) { textLines.push(link.href); }
      textLines.push("");

      h.push('<div style="border-top:1px solid #e1e1e1;padding:10px 0 6px">');
      h.push("<p style=\"margin:0\"><b>" +
        (link ? '<a href="' + esc(link.href) + '">' + esc(it.billNumber) + "</a>" : esc(it.billNumber)) +
        "</b></p>");
      h.push('<p style="margin:2px 0;color:#616161">Distributed To: <b>' + esc(dist) +
        "</b> · Comment Requested From: <b>" + esc(from) + "</b></p>");
      h.push('<p style="margin:4px 0">' + esc(it.brief || "(no summary)").replace(/\n/g, "<br>") + "</p>");
      h.push("</div>");
    });

    if (opts.preparedBy) {
      h.push('<p style="color:#616161;font-size:12px">Prepared by ' + esc(opts.preparedBy) +
        " with Legislative Report Router.</p>");
      textLines.push("Prepared by " + opts.preparedBy + " with Legislative Report Router.");
    }
    h.push("</div>");

    return {
      subject: subjectFor(date),
      html: h.join("\n"),
      text: textLines.join("\n"),
      count: included.length,
    };
  }

  var api = {
    buildDailyReport: buildDailyReport,
    subjectFor: subjectFor,
    extractEmails: extractEmails,
    filterNew: filterNew,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrReportGen = api; }
})(typeof self !== "undefined" ? self : this);
