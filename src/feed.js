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

  var centralFormat = null;
  // The RSS timestamps omit their zone. Resolve Central wall time through
  // the timezone database so summer filings aren't shifted an hour ahead.
  function centralDate(value) {
    value = String(value || "").trim();
    if (/(?:GMT|UTC|CST|CDT|Z|[+-]\d{2}:?\d{2})$/i.test(value)) { return new Date(value); }
    var wall = new Date(value + " GMT");
    if (isNaN(wall.getTime())) { return new Date(NaN); }
    if (!centralFormat) { centralFormat = new Intl.DateTimeFormat("en-US", { timeZone: "America/Chicago", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", second: "2-digit", hourCycle: "h23" }); }
    var guess = wall.getTime() + 6 * 3600000;
    for (var attempt = 0; attempt < 3; attempt++) {
      var parts = {};
      centralFormat.formatToParts(new Date(guess)).forEach(function (p) { if (p.type !== "literal") { parts[p.type] = Number(p.value); } });
      var shown = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second);
      var delta = wall.getTime() - shown;
      if (!delta) { return new Date(guess); }
      guess += delta;
    }
    return new Date(NaN); // a nonexistent clock time establishes no coverage
  }

  /** RSS XML → [{bill, description, link, pubDate(Date), raw}] newest first. */
  function parseFeed(xml) {
    var out = [];
    var re = /<item>([\s\S]*?)<\/item>/g;
    var m;
    while ((m = re.exec(String(xml || ""))) !== null) {
      var block = m[1];
      var title = cdata(block, "title");
      var amendment = title.match(/^([HS]|CCS|CCH)\s*-?\s*(\d+)\s*\(\s*([A-Z]{1,4})\s*-?\s*(\d+)\s*\)$/i);
      var bill = title.replace(/\s+/g, "").toUpperCase(); // "HF 437" -> "HF437"
      if (amendment) { bill = amendment[1].toUpperCase() + (/^[HS]$/i.test(amendment[1]) ? "-" : "") + amendment[2]; }
      if (!amendment && !/^[A-Z]{2,4}\d+[A-Z]?$/.test(bill)) { continue; }
      out.push({
        bill: bill,
        relatedBill: amendment ? amendment[3].toUpperCase() + (/^[HS]$/i.test(amendment[3]) ? "-" : "") + amendment[4] : "",
        documentType: amendment ? (/^CC/i.test(amendment[1]) ? "Conference report" : "Amendment") : (bill.match(/^[A-Z]+/) || [""])[0],
        description: cdata(block, "description"),
        link: cdata(block, "link"),
        pubDate: centralDate(cdata(block, "pubDate")), // feed times are Central, DST-aware
      });
    }
    out.sort(function (a, b) { return b.pubDate - a.pubDate; });
    return out;
  }

  // Public custom RSS/Atom feeds: use configured bill identifiers, never Iowa time assumptions.
  function feedUrl(value) {
    var u;
    try { u = new URL(String(value || "").trim()); } catch (e) { throw new Error("Enter a complete HTTPS feed URL."); }
    if (u.protocol !== "https:" || u.username || u.password || u.hash) { throw new Error("Use a public HTTPS feed URL without a username, password or fragment."); }
    return u.href;
  }
  function parseCustomFeed(xml, opts, Parser) {
    opts = opts || {}; Parser = Parser || DOMParser;
    if (/<!DOCTYPE|<!ENTITY/i.test(xml)) { throw new Error("Feed documents with DTDs or custom entities are not supported."); }
    var doc = new Parser().parseFromString(String(xml), "application/xml");
    var type = doc.documentElement && doc.documentElement.localName.toLowerCase();
    if (doc.getElementsByTagName("parsererror").length || ["rss", "rdf", "feed"].indexOf(type) < 0) { throw new Error("The URL must return valid RSS or Atom XML, not a web page."); }
    if (type === "rss" && !doc.getElementsByTagName("channel").length) { throw new Error("RSS must contain a channel element."); }
    var prefixes = (opts.identifiers || []).map(function (v) { return String(v).trim().toUpperCase(); }).filter(function(v){return /^[A-Z]{1,6}$/.test(v);}).sort(function(a,b){return b.length-a.length;});
    if (!prefixes.length) { throw new Error("Set the bill identifiers for this legislature first (for example HB, SB). "); }
    var billPattern = new RegExp("(?:^|[^A-Z0-9])(" + prefixes.join("|") + ")[\\s._/-]*(\\d+)[\\s]*([A-Z]?)(?=$|[^A-Z0-9])", "i");
    function children(node, name) { return Array.from(node.children).filter(function(c){return c.localName.toLowerCase()===name;}); }
    function text(node,name) { var c=children(node,name)[0];return c?c.textContent.trim():""; }
    function plain(value) { return value.replace(/<[^>]*>/g," ").replace(/\s+/g," ").trim(); }
    var nodes=Array.from(doc.getElementsByTagNameNS("*",type==="feed"?"entry":"item")), entries=[], skipped=0, emptyNotices=0;
    nodes.forEach(function(node){
      var title=plain(text(node,"title")), links=children(node,"link"), rawLink="";
      if ((opts.emptyTitles || []).indexOf(title) >= 0) { emptyNotices++; return; }
      links.some(function(l){if(type!=="feed"||!l.getAttribute("rel")||l.getAttribute("rel")==="alternate"){rawLink=l.getAttribute("href")||l.textContent.trim();return true;}return false;});
      var link="";try { var url=new URL(rawLink,opts.url);if(rawLink && url.protocol==="https:" && !url.username && !url.password){link=url.href;} } catch(e) { /* omit unsafe links */ }
      var match=title.match(billPattern)||link.match(billPattern);
      var dateText=text(node,"pubdate")||text(node,"published")||text(node,"updated")||text(node,"date");
      // A date-only value is unambiguous. Times must include their zone.
      var dated=/^\d{4}-\d{2}-\d{2}$/.test(dateText)||/(?:Z|[+-]\d{2}:?\d{2}|GMT|UTC|[ECMP][SD]T)$/i.test(dateText);
      var date=dated?new Date(dateText):new Date(NaN);
      if(!match||isNaN(date.getTime())){skipped++;return;}
      entries.push({bill:match[1].toUpperCase()+match[2]+match[3].toUpperCase(),description:plain(text(node,"description")||text(node,"summary")||text(node,"content")||title),link:link,pubDate:date,documentType:match[1].toUpperCase(),feedSource:opts.url||""});
    });
    entries.sort(function(a,b){return b.pubDate-a.pubDate;});
    if(nodes.length > emptyNotices && !entries.length){throw new Error("No usable bill entries. Check bill identifiers and feed dates; dates with a time must include a time zone.");}
    return {entries:entries,total:nodes.length,skipped:skipped};
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
      documentType: entry.documentType || docType,
      relatedBill: entry.relatedBill || "",
      parentBill: entry.parentBill || null,
      screening: entry.screening || null,
      screeningText: entry.fullText || "",
      distributedTo: [],
      commentRequestedFrom: [],
      title: entry.description.slice(0, 120),
      brief: entry.description,
      referencedBills: entry.relatedBill ? [entry.relatedBill] : [],
      sourceLinks: (entry.link ? [{ text: entry.bill + " (BillBook)", href: entry.link }] : []).concat(entry.parentBill && entry.parentBill.link ? [{ text: "Parent bill " + entry.parentBill.bill, href: entry.parentBill.link }] : []).concat(entry.parentBill && entry.parentBill.textUrl ? [{ text: "Parent bill text (" + (entry.parentBill.version || "mirrored version") + ")", href: entry.parentBill.textUrl }] : []),
      sourceBlock: entry.bill + "\n" + entry.description,
      parserConfidence: entry.screening && entry.screening.status === "needs-screening" ? 0.6 : 0.9,
      parserWarnings: ["From the newly-filed feed — assign divisions before publishing."].concat(entry.screening ? [
        entry.screening.reasons.length ? "Possible agency impact: " + entry.screening.reasons.join("; ") + ". Verify applicability." : "Needs manual screening — absence of a keyword match does not rule out agency impact.",
        entry.screening.fullTextAvailable ? "Screened " + (entry.screening.textVersion || "available") + " text; check subsequent revisions." : "Summary only — open the source to review full text.",
      ] : []),
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
    feedUrl: feedUrl,
    parseCustomFeed: parseCustomFeed,
    parseOpenStates: parseOpenStates,
    watchFilter: watchFilter,
    toLegislativeItem: toLegislativeItem,
    addBusinessDays: addBusinessDays,
  };
  if (typeof module !== "undefined" && module.exports) { module.exports = api; }
  else { root.LrrFeed = api; }
})(typeof self !== "undefined" ? self : this);
