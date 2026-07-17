/*
 * Legislative Report Router — Iowa GA "Newly Filed" feed support (pure).
 *
 * The Legislature publishes intraday RSS feeds of newly filed bills and
 * amendments (legis.iowa.gov/subscribe/rss). Their CORS header is broken,
 * so a scheduled GitHub Action mirrors the XML into this repo's Pages
 * (feeds/IowaBills.xml, feeds/Amendments.xml) and the pane fetches it
 * same-origin. This module parses the RSS and turns watched entries into
 * LegislativeItems that flow through the normal Review → Publish pipeline.
 *
 * RSS parsing is regex-over-CDATA on purpose: the feed shape is stable,
 * and staying string-based keeps this module Node-testable (no DOMParser).
 */
(function (root) {
  "use strict";

  function cdata(block, tag) {
    var m = new RegExp("<" + tag + ">(?:<!\\[CDATA\\[)?([\\s\\S]*?)(?:\\]\\]>)?</" + tag + ">").exec(block);
    return m ? m[1].trim() : "";
  }

  /** RSS XML → [{bill, description, link, pubDate(Date), raw}] newest first. */
  function parseFeed(xml) {
    var out = [];
    var re = /<item>([\s\S]*?)<\/item>/g;
    var m;
    while ((m = re.exec(String(xml || ""))) !== null) {
      var block = m[1];
      var title = cdata(block, "title");
      var bill = title.replace(/\s+/g, "").toUpperCase(); // "HF 437" -> "HF437"
      if (!/^[A-Z]{2,4}\d+[A-Z]?$/.test(bill)) { continue; }
      out.push({
        bill: bill,
        description: cdata(block, "description"),
        link: cdata(block, "link"),
        pubDate: new Date(cdata(block, "pubDate") + " GMT-0600"), // feed times are Central, no zone given
      });
    }
    out.sort(function (a, b) { return b.pubDate - a.pubDate; });
    return out;
  }

  /**
   * Filter feed entries to the watch window and watch terms.
   * terms: array of lowercase substrings; empty array = everything.
   */
  function watchFilter(entries, terms, sinceDays, now) {
    var cutoff = (now ? new Date(now) : new Date()).getTime() - (sinceDays == null ? 3 : sinceDays) * 864e5;
    var t = (terms || []).map(function (x) { return String(x).toLowerCase().trim(); }).filter(Boolean);
    return (entries || []).filter(function (e) {
      if (!(e.pubDate instanceof Date) || isNaN(e.pubDate) || e.pubDate.getTime() < cutoff) { return false; }
      if (!t.length) { return true; }
      var hay = (e.bill + " " + e.description).toLowerCase();
      return t.some(function (term) { return hay.indexOf(term) !== -1; });
    });
  }

  /** Feed entry → LegislativeItem (divisions left empty for the coordinator). */
  function toLegislativeItem(entry, reportKey, idx) {
    var docType = (entry.bill.match(/^[A-Z]+/) || [""])[0];
    return {
      id: "feed-" + (idx + 1),
      reportId: reportKey || "new-filings",
      billNumber: entry.bill,
      documentType: docType,
      distributedTo: [],
      commentRequestedFrom: [],
      title: entry.description.slice(0, 120),
      brief: entry.description,
      referencedBills: [],
      sourceLinks: entry.link ? [{ text: entry.bill + " (BillBook)", href: entry.link }] : [],
      sourceBlock: entry.bill + "\n" + entry.description,
      parserConfidence: 0.9,
      parserWarnings: ["From the newly-filed feed — assign divisions before publishing."],
      routingStatus: "unmatched",
      matchedRoutingRules: [],
    };
  }

  /**
   * Open States mirror JSON → same entry shape as parseFeed.
   * Mirror emits {results:[{identifier, title, first_action_date,
   * created_at, openstates_url, latest_action_description}]}.
   */
  function parseOpenStates(jsonText) {
    var data;
    try { data = typeof jsonText === "string" ? JSON.parse(jsonText) : jsonText; }
    catch (e) { return []; }
    var out = [];
    (data.results || []).forEach(function (b) {
      var bill = String(b.identifier || "").replace(/\s+/g, "").toUpperCase();
      if (!/^[A-Z]{1,8}\d+[A-Z]?$/.test(bill)) { return; }
      var when = b.first_action_date || b.created_at || "";
      out.push({
        bill: bill,
        description: (b.title || "") +
          (b.latest_action_description ? " — Latest action: " + b.latest_action_description : ""),
        link: b.openstates_url || "",
        pubDate: new Date(String(when).slice(0, 10) + "T12:00:00"),
      });
    });
    out.sort(function (a, b) { return b.pubDate - a.pubDate; });
    return out;
  }

  /** Due date = start + n business days (skips Sat/Sun), same clock time. */
  function addBusinessDays(start, n) {
    var d = new Date(start);
    var left = n;
    while (left > 0) {
      d.setDate(d.getDate() + 1);
      var dow = d.getDay();
      if (dow !== 0 && dow !== 6) { left--; }
    }
    return d;
  }

  var api = {
    parseFeed: parseFeed,
    parseOpenStates: parseOpenStates,
    watchFilter: watchFilter,
    toLegislativeItem: toLegislativeItem,
    addBusinessDays: addBusinessDays,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrFeed = api; }
})(typeof self !== "undefined" ? self : this);
