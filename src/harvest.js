/*
 * Legislative Report Router — fiscal harvest (pure).
 *
 * Divisions don't type into list columns — they reply in the bill's Teams
 * thread or attach a DOCX/Excel with the numbers. This module reads TEXT
 * (a reply body, an extracted document) and pulls out candidate fiscal
 * data: dollar amounts with surrounding context, an impact-severity read,
 * and a fiscal-year mention. Candidates are ALWAYS reviewed by the
 * coordinator before anything is written to the tracker — this is an
 * ingestion assistant, not an auto-writer.
 */
(function (root) {
  "use strict";

  /* --------------------------------------------------------------- money */

  var MONEY_RE = /(\$\s?[\d][\d,]*(?:\.\d+)?\s*(?:k|m|b|thousand|million|billion)?|\b[\d][\d,]*(?:\.\d+)?\s*(?:thousand|million|billion)\b|\(\s?\$?[\d][\d,]*(?:\.\d+)?\s*(?:k|m|b|thousand|million|billion)?\s?\))/gi;

  function normAmount(raw) {
    var s = String(raw).trim();
    var negative = /^\(.*\)$/.test(s);
    var m = /([\d][\d,]*(?:\.\d+)?)\s*(k|m|b|thousand|million|billion)?/i.exec(s.replace(/[$(),]/g, function (c) { return c === "," ? "," : " "; }));
    if (!m) { return null; }
    var n = Number(m[1].replace(/,/g, ""));
    if (isNaN(n)) { return null; }
    var mult = { k: 1e3, thousand: 1e3, m: 1e6, million: 1e6, b: 1e9, billion: 1e9 }[(m[2] || "").toLowerCase()] || 1;
    return (negative ? -1 : 1) * n * mult;
  }

  /** All dollar figures in a text, each with ~60 chars of context. */
  function findMoney(text) {
    var s = String(text || "");
    var out = [];
    var m;
    MONEY_RE.lastIndex = 0;
    while ((m = MONEY_RE.exec(s)) !== null) {
      var amount = normAmount(m[0]);
      if (amount === null || Math.abs(amount) < 100) { continue; } // "$5 fee" prose noise
      var start = Math.max(0, m.index - 60);
      var context = s.slice(start, Math.min(s.length, m.index + m[0].length + 60))
        .replace(/\s+/g, " ").trim();
      out.push({ amount: amount, raw: m[0].trim(), context: context, index: m.index });
    }
    return out;
  }

  /**
   * Pick the figure to propose: prefer one whose context says total/cost/
   * estimate/impact; among those (or all, if none) take the LARGEST
   * magnitude — cover letters mention small fees, the estimate is the
   * big number. Savings (negatives) win only if nothing positive exists.
   */
  function bestMoney(hits) {
    if (!hits.length) { return null; }
    var scored = hits.map(function (h) {
      var ctx = h.context.toLowerCase();
      var score = 0;
      if (/total|cost|estimate|impact|fiscal/.test(ctx)) { score += 2; }
      if (/annual|per year|yearly|\/yr/.test(ctx)) { score += 1; }
      return { h: h, score: score };
    });
    var top = Math.max.apply(null, scored.map(function (x) { return x.score; }));
    var pool = scored.filter(function (x) { return x.score === top; }).map(function (x) { return x.h; });
    pool.sort(function (a, b) { return Math.abs(b.amount) - Math.abs(a.amount); });
    return pool[0];
  }

  /* ------------------------------------------------------------ severity */

  var SEVERITY_RULES = [
    { re: /\bno (fiscal )?(impact|cost)\b|\bzero (fiscal )?impact\b/i, sev: "None" },
    { re: /\bcritical|severe impact|major impact\b/i, sev: "Critical" },
    { re: /\bhigh impact|significant(ly)? (fiscal|cost|impact)|substantial (cost|impact)\b/i, sev: "High" },
    { re: /\bmoderate\b/i, sev: "Moderate" },
    { re: /\bminimal|negligible|minor (fiscal|cost|impact)|low impact\b/i, sev: "Low" },
  ];

  function findSeverity(text) {
    var s = String(text || "");
    for (var i = 0; i < SEVERITY_RULES.length; i++) {
      if (SEVERITY_RULES[i].re.test(s)) { return SEVERITY_RULES[i].sev; }
    }
    return "";
  }

  /* ------------------------------------------------------------------ FY */

  function findFy(text) {
    var m = /\b(?:s?fy\s?-?\s?(\d{2,4})|fiscal year\s+(\d{4}))\b/i.exec(String(text || ""));
    if (!m) { return ""; }
    var n = Number(m[1] || m[2]);
    if (n < 100) { n += 2000; }
    return "SFY" + n;
  }

  /* ----------------------------------------------------------- candidates */

  /**
   * One source (reply body or extracted attachment text) → a candidate,
   * or null when the text carries no fiscal signal at all.
   * src: { bill, division, source (label for evidence), text, author, date }
   */
  function analyzeSource(src) {
    var money = bestMoney(findMoney(src.text));
    var severity = findSeverity(src.text);
    var fy = findFy(src.text);
    if (!money && !severity) { return null; }
    return {
      kind: src.kind || "reply",
      bill: src.bill,
      division: src.division || "",
      cost: money ? money.amount : null,
      costRaw: money ? money.raw : "",
      severity: severity || (money ? "Unknown" : ""),
      fy: fy,
      evidence: (money ? "“…" + money.context + "…”" : severity ? "severity language found" : ""),
      source: src.source || "",
      author: src.author || "",
      date: src.date || "",
    };
  }

  /**
   * One candidate per bill×division: an attachment with a dollar figure
   * beats a prose-only reply; newer beats older at equal strength.
   */
  function mergeCandidates(list) {
    var byKey = {};
    (list || []).forEach(function (c) {
      if (!c) { return; }
      var k = c.bill + "|" + (c.division || "");
      var prev = byKey[k];
      if (!prev) { byKey[k] = c; return; }
      // an attached document's dollar figure outranks prose commentary —
      // the spreadsheet IS the estimate; the reply is discussion around it
      var strength = function (x) {
        return (x.cost !== null ? 2 : 0) +
               (x.cost !== null && x.kind === "attachment" ? 2 : 0) +
               (x.severity && x.severity !== "Unknown" ? 1 : 0);
      };
      if (strength(c) > strength(prev) ||
          (strength(c) === strength(prev) && String(c.date) > String(prev.date))) {
        // carry forward anything the weaker one knew that the winner doesn't
        if (c.cost === null && prev.cost !== null) { c.cost = prev.cost; c.costRaw = prev.costRaw; }
        if ((!c.severity || c.severity === "Unknown") && prev.severity) { c.severity = prev.severity; }
        if (!c.fy && prev.fy) { c.fy = prev.fy; }
        byKey[k] = c;
      } else {
        if (prev.cost === null && c.cost !== null) { prev.cost = c.cost; prev.costRaw = c.costRaw; }
        if ((!prev.severity || prev.severity === "Unknown") && c.severity) { prev.severity = c.severity; }
        if (!prev.fy && c.fy) { prev.fy = c.fy; }
      }
    });
    return Object.keys(byKey).sort().map(function (k) { return byKey[k]; });
  }

  /** Strip a Teams HTML reply body down to analyzable text. */
  function htmlToText(html) {
    return String(html || "")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/(p|div|li|tr)>/gi, "\n")
      .replace(/<[^>]+>/g, " ")
      .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&")
      .replace(/&#39;/g, "'").replace(/&quot;/g, '"').replace(/&nbsp;/g, " ")
      .replace(/[ \t]+/g, " ").replace(/[ \t]*\n[ \t]*/g, "\n").trim();
  }

  var api = {
    findMoney: findMoney,
    bestMoney: bestMoney,
    findSeverity: findSeverity,
    findFy: findFy,
    analyzeSource: analyzeSource,
    mergeCandidates: mergeCandidates,
    htmlToText: htmlToText,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrHarvest = api; }
})(typeof self !== "undefined" ? self : this);
