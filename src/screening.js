/* Agency-impact screening: retain every filing, explain possible impacts.
 * Deterministic clues, not a legal opinion or a complete relevance decision.
 */
(function (root) {
  "use strict";
  var chaptersApi = typeof module !== "undefined" && module.exports ? require("./chapters.js") : root.LrrChapters;
  var feedApi = typeof module !== "undefined" && module.exports ? require("./feed.js") : root.LrrFeed;
  var TOPICS = [
    ["Contracts and procurement", /\b(contract\w*|procure\w*|purchas\w*|vendor\w*|indemn\w*|arbitration|competitive bidding)\b/i],
    ["Research and intellectual property", /\b(research|intellectual property|patent\w*|copyright\w*|licens\w*|university|universities)\b/i],
    ["Employment and benefits", /\b(employ\w*|personnel|workforce|collective bargaining|pension\w*|retirement|workers.? compensation)\b/i],
    ["Funding and grants", /\b(appropriat\w*|budget\w*|funding|grants?|federal funds?|tax\w*|revenue\w*)\b/i],
    ["Agency administration and legal duties", /\b(state agenc\w*|state government|public records|open meetings|administrative rules?|cybersecurity|data privacy|liability|attorney general)\b/i],
  ];
  function assess(entry, opts) {
    opts = opts || {};
    var text = [entry.bill, entry.description, entry.fullText || ""].join(" ");
    var reasons = [];
    (opts.terms || []).forEach(function (term) {
      term = String(term).trim();
      if (term && text.toLowerCase().indexOf(term.toLowerCase()) !== -1) { reasons.push('Watch term: "' + term + '"'); }
    });
    TOPICS.forEach(function (topic) { if (topic[1].test(text)) { reasons.push(topic[0]); } });
    var chapters = chaptersApi.extractChapters(text);
    var tracked = opts.matchChapters === false ? [] : chaptersApi.matchTracked(chapters, opts.tracked);
    if ((opts.stateName || "Iowa") === "Iowa") {
      var agencyChapters = chaptersApi.matchTracked(chapters, ["8", "8A", "8B", "13", "19A", "20", "22", "23", "68B", "70A"]);
      if (agencyChapters.length) { reasons.push("Agency-wide Code: " + agencyChapters.join(", ")); }
    }
    if (tracked.length) { reasons.push("Tracked Code: " + tracked.join(", ")); }
    return {
      status: reasons.length ? "possible-impact" : "needs-screening",
      reasons: reasons,
      chapters: chapters,
      fullTextAvailable: !!entry.fullText,
      textVersion: entry.textVersion || "",
    };
  }
  function screen(entries, opts) {
    opts = opts || {};
    // Empty terms intentionally keeps every filing in the date window.
    var recent = feedApi.watchFilter(entries, [], opts.days, opts.now);
    var seen = {};
    var out = [];
    recent.forEach(function (entry) {
      // GA separates identical identifiers from different assemblies.
      var key = entryKey(entry);
      if (seen[key]) { return; }
      seen[key] = true;
      var copy = Object.assign({}, entry);
      copy.screening = assess(copy, opts);
      out.push(copy);
    });
    return out.sort(function (a, b) {
      return (a.screening.status === "needs-screening") - (b.screening.status === "needs-screening") || b.pubDate - a.pubDate;
    });
  }
  function entryKey(entry) {
    var ga = String(entry.link || "").match(/[?&]ga=(\d+)\b/i);
    return (ga ? ga[1] + ":" : "") + entry.bill;
  }
  function enrich(entries, mirror) {
    var documents = mirror && mirror.documents || {};
    return (entries || []).map(function (entry) {
      var doc = documents[entryKey(entry)];
      if (!doc || !doc.text) { return entry; }
      return Object.assign({}, entry, { fullText: doc.text, textVersion: doc.version, textUrl: doc.url });
    });
  }
  var api = { assess: assess, screen: screen, entryKey: entryKey, enrich: enrich };
  if (typeof module !== "undefined" && module.exports) {
    module.exports = api;
  } else { root.LrrScreening = api; }
})(typeof self !== "undefined" ? self : this);
