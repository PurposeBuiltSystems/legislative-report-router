/*
 * Legislative Report Router — contact-list import (pure logic).
 *
 * The coordinator already maintains the distribution list (Outlook
 * contacts, an Excel sheet, the report's To: line). This module turns a
 * PASTED spreadsheet — copied straight out of Excel/Outlook — into
 * per-division email assignments for the routing rules. Clipboard paste
 * means no new Graph scopes: the data arrives like any typed input.
 *
 * Accepted shapes (tab- or comma-separated; header row optional):
 *   Name        Email                 Division
 *   Jane Doe    jane@agency.gov       MVD
 *   Bob Roe     bob@agency.gov        MVD/TDD
 * or two columns:  email, division
 */
(function (root) {
  "use strict";

  var EMAIL_RE = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;

  function splitDivisions(s) {
    return String(s || "").split(/[\/;&]|\band\b/i)
      .map(function (t) { return t.replace(/\s+/g, " ").trim(); })
      .filter(Boolean);
  }

  /**
   * Parse pasted rows. Returns {rows:[{name,email,divisions[]}], errors:[],
   * skippedHeader:bool}. Delimiter: tab if present, else comma.
   */
  function parseContactRows(text) {
    var lines = String(text || "").split(/\r?\n/).map(function (l) { return l.trim(); }).filter(Boolean);
    var rows = [];
    var errors = [];
    var skippedHeader = false;

    lines.forEach(function (line, idx) {
      var delim = line.indexOf("\t") !== -1 ? "\t" : ",";
      var cells = line.split(delim).map(function (c) { return c.trim().replace(/^"|"$/g, ""); });

      // header row: mentions email but carries no address
      if (idx === 0 && !/@/.test(line) && /email/i.test(line)) { skippedHeader = true; return; }

      var emailIdx = -1;
      for (var i = 0; i < cells.length; i++) {
        if (/@/.test(cells[i])) { emailIdx = i; break; }
      }
      if (emailIdx === -1) {
        errors.push("Line " + (idx + 1) + ": no email address found — skipped.");
        return;
      }
      var email = cells[emailIdx].toLowerCase();
      if (!EMAIL_RE.test(email)) {
        errors.push("Line " + (idx + 1) + ': "' + cells[emailIdx] + '" doesn\'t look like a valid address — skipped.');
        return;
      }

      // division: last non-email cell after the email column if present,
      // else the only other cell; name: first cell when 3+ columns
      var division = "";
      var name = "";
      if (cells.length >= 3) {
        name = emailIdx > 0 ? cells[0] : "";
        division = cells[cells.length - 1] !== email ? cells[cells.length - 1] : "";
      } else if (cells.length === 2) {
        division = cells[1 - emailIdx];
      }
      var divisions = splitDivisions(division);
      if (!divisions.length) {
        errors.push("Line " + (idx + 1) + ": no division for " + email + " — skipped (add a division column).");
        return;
      }
      rows.push({ name: name, email: email, divisions: divisions });
    });

    return { rows: rows, errors: errors, skippedHeader: skippedHeader };
  }

  function norm(s) { return String(s || "").replace(/\s+/g, " ").trim().toLowerCase(); }

  function ruleMatches(rule, division) {
    var d = norm(division);
    return [rule.divisionCode, rule.divisionName].concat(rule.aliases || [])
      .some(function (n) { return norm(n) === d; });
  }

  /**
   * Build an apply-plan against existing rules:
   *   [{division, rule|null, addEmails[], existingCount}]
   * addEmails excludes addresses the rule already has.
   */
  function mergePlan(rows, rules) {
    var byDivision = {};
    var order = [];
    (rows || []).forEach(function (r) {
      r.divisions.forEach(function (dv) {
        var key = norm(dv);
        if (!byDivision[key]) { byDivision[key] = { division: dv, emails: [] }; order.push(key); }
        if (byDivision[key].emails.indexOf(r.email) === -1) { byDivision[key].emails.push(r.email); }
      });
    });
    return order.map(function (key) {
      var g = byDivision[key];
      var rule = (rules || []).filter(function (r) { return ruleMatches(r, g.division); })[0] || null;
      var existing = rule ? (rule.emails || []).map(function (e) { return e.toLowerCase(); }) : [];
      var add = g.emails.filter(function (e) { return existing.indexOf(e) === -1; });
      return { division: g.division, rule: rule, addEmails: add, existingCount: existing.length };
    });
  }

  var api = { parseContactRows: parseContactRows, mergePlan: mergePlan, _internals: { splitDivisions: splitDivisions } };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrContacts = api; }
})(typeof self !== "undefined" ? self : this);
