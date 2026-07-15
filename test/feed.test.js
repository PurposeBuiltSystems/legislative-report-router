/* Offline tests for the newly-filed feed module. Run: node test/feed.test.js */
"use strict";
var F = require("../src/feed.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

// Fixture: real feed shape (CDATA, Central-time pubDates, "HF 437" spacing)
var XML = [
  '<?xml version="1.0" encoding="UTF-8"?>',
  '<rss version="2.0"><channel><title>Iowa Legislature - Newly Filed Bills</title>',
  "<item><title><![CDATA[HF 437]]></title>",
  "<link><![CDATA[https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HF437]]></link>",
  "<description><![CDATA[A bill for an act establishing a center for intellectual freedom at the university of Iowa.  (Formerly HSB 52.)]]></description>",
  "<pubDate>Mon, 13 Jul 2026 12:37:28</pubDate></item>",
  "<item><title><![CDATA[SF 2103]]></title>",
  "<link><![CDATA[https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=SF2103]]></link>",
  "<description><![CDATA[A bill for an act relating to motor vehicle franchises and the department of transportation.]]></description>",
  "<pubDate>Tue, 14 Jul 2026 09:08:37</pubDate></item>",
  "<item><title><![CDATA[HSB 622]]></title>",
  "<link><![CDATA[https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HSB622]]></link>",
  "<description><![CDATA[A bill for an act relating to health and human services programs.]]></description>",
  "<pubDate>Mon, 05 Jan 2026 13:15:20</pubDate></item>",
  "</channel></rss>",
].join("\n");

var NOW = new Date("2026-07-14T18:00:00-05:00");

// 1. Parse: 3 entries, bill numbers normalized, newest first
var entries = F.parseFeed(XML);
check("entry count", entries.length, 3);
check("normalized bill", entries[0].bill, "SF2103");
check("newest first", entries.map(function (e) { return e.bill; }), ["SF2103", "HF437", "HSB622"]);
check("link kept", entries[0].link.indexOf("BillBook") !== -1, true);

// 2. Watch window: default 3 days drops the January bill
var recent = F.watchFilter(entries, [], 3, NOW);
check("window filter", recent.map(function (e) { return e.bill; }), ["SF2103", "HF437"]);

// 3. Watch terms: transportation matches SF2103 only
var watched = F.watchFilter(entries, ["transportation", "motor vehicle"], 3, NOW);
check("term filter", watched.map(function (e) { return e.bill; }), ["SF2103"]);

// 4. Empty terms = everything in window
check("no terms = all", F.watchFilter(entries, [], 365, NOW).length, 3);

// 5. Feed entry -> LegislativeItem ready for Review
var item = F.toLegislativeItem(entries[0], "new-filings-2026-07-14", 0);
check("item bill", item.billNumber, "SF2103");
check("item docType", item.documentType, "SF");
check("item brief", item.brief.indexOf("motor vehicle franchises") !== -1, true);
check("item link", item.sourceLinks[0].href.indexOf("SF2103") !== -1, true);
check("item needs divisions", item.distributedTo, []);
check("item warning present", item.parserWarnings.length, 1);

// 6. Business-day due dates (Fri + 2 -> Tue)
var due = F.addBusinessDays(new Date("2026-07-17T10:00:00"), 2); // Friday
check("skips weekend", due.getDay(), 2); // Tuesday
check("due date", due.toISOString().slice(0, 10), "2026-07-21");

if (failures) {
  console.error("\n" + failures + " feed test(s) FAILED");
  process.exit(1);
}
console.log("All feed tests passed.");
