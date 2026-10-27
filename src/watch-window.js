/* Version-aware, shared watch review checkpoints. No cutoff-only skipping. */
(function (root) {
  "use strict";
  async function hash(value) {
    var bytes = new TextEncoder().encode(JSON.stringify(value));
    var digest = await crypto.subtle.digest("SHA-256", bytes);
    return Array.from(new Uint8Array(digest)).map(function (b) { return b.toString(16).padStart(2, "0"); }).join("");
  }
  function normalized(values) { return Array.from(new Set((values || []).map(function (v) { return String(v).trim().toLowerCase(); }).filter(Boolean))).sort(); }
  async function scope(config) {
    return hash({ state: config.stateName, terms: normalized(config.terms), tracked: normalized(config.tracked), matchChapters: config.matchChapters !== false });
  }
  async function fingerprint(entry) {
    return hash({ bill: entry.bill, relatedBill: entry.relatedBill || "", link: entry.link || "",
      published: new Date(entry.pubDate).toISOString(), description: entry.description || "",
      fullText: entry.fullText || "", version: entry.textVersion || "", parent: entry.parentBill || null });
  }
  function history(rows, key) {
    var reviewed = new Set(), last = null;
    (rows || []).forEach(function (row) {
      var f = row.fields || {};
      if (f.ReportKey !== "watch:" + key) { return; }
      if (f.Status === "watch-reviewed" && /^watch:[a-f0-9]{64}:[a-f0-9]{64}$/.test(f.IdempotencyKey || "")) { reviewed.add(f.IdempotencyKey); }
      if (f.Status === "watch-window") {
        try {
          var window = JSON.parse(f.SourceSubject);
          if (window.reviewedAt && (!last || window.reviewedAt > last.reviewedAt)) { last = window; }
        } catch (e) { /* unrelated or old malformed metadata establishes no coverage */ }
      }
    });
    return { reviewed: reviewed, last: last };
  }
  async function prepare(entries, config, now, rows) {
    var end = new Date(now), key = await scope(config);
    var output = await Promise.all(entries.filter(function (e) { return new Date(e.pubDate) <= end && new Date(e.pubDate).getTime() >= end.getTime() - config.days * 864e5; }).map(async function (entry) {
      return { entry: entry, key: "watch:" + key + ":" + await fingerprint(entry) };
    }));
    var prior = history(rows, key);
    return { scope: key, days: config.days, from: new Date(end.getTime() - config.days * 864e5).toISOString(), to: end.toISOString(),
      checkedAt: end.toISOString(), stateName: config.stateName, entries: output, reviewed: prior.reviewed, lastReviewed: prior.last };
  }
  var api = { scope: scope, fingerprint: fingerprint, history: history, prepare: prepare };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrWatch = api; }
})(typeof self !== "undefined" ? self : this);
