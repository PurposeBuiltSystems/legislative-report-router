/* Offline tests for the Daily Bill Report generator. Run: node test/reportgen.test.js */
"use strict";
var G = require("../src/reportgen.js");
var P = require("../src/parser.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

var items = [
  { billNumber: "HF2801", routingStatus: "matched", distributedTo: ["MVD"], commentRequestedFrom: ["MVD"],
    brief: "vehicle registration modernization.\nAmends chapter 321.",
    sourceLinks: [{ text: "HF2801", href: "https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HF2801" }] },
  { billNumber: "SF2110", routingStatus: "matched", distributedTo: ["TDD", "SOD"], commentRequestedFrom: ["TDD"],
    brief: "highway funding formula.", sourceLinks: [] },
  { billNumber: "HF2802", routingStatus: "excluded", distributedTo: ["ELT"], commentRequestedFrom: [], brief: "x", sourceLinks: [] },
];

// 1. Subject matches the coordinator's convention
check("subject", G.subjectFor(new Date(2026, 6, 26)), "7-26-2026 Daily Bill Report");

// 2. Build: excluded bills stay out; both formats carry the content
var rep = G.buildDailyReport(items, { sessionName: "2027 Session", commentWindow: "48 business hours.", date: new Date(2026, 6, 26) });
check("count excludes excluded", rep.count, 2);
check("html has link (escaped)", rep.html.indexOf("BillBook?ga=91&amp;ba=HF2801") !== -1, true);
check("html divisions", rep.html.indexOf("MVD") !== -1 && rep.html.indexOf("TDD, SOD") !== -1, true);
check("html session", rep.html.indexOf("2027 Session") !== -1, true);
check("html no excluded bill", rep.html.indexOf("HF2802") === -1, true);
check("text has window", rep.text.indexOf("48 business hours.") !== -1, true);

// Header context survives both draft paths; no configured session is required.
var context = G.buildDailyReport(items, { date: new Date(2026, 3, 30), deadlineLabel: "Friday, 5 PM" });
check("default Iowa session", context.text.includes("2026 Legislative Session – 2nd Session, 91st Iowa General Assembly"), true);
check("next assembly", G.buildDailyReport([], { date: new Date(2027, 0, 4) }).text.includes("92nd Iowa General Assembly"), true);
check("other state header", G.buildDailyReport([], { stateName: "Minnesota", date: new Date(2026, 0, 4) }).text.includes("2026 Minnesota Legislative Session"), true);
for (var format of ["text", "html"]) {
  check(format + " distribution context", context[format].includes("Distribution List:") && context[format].includes("FHWA"), true);
  check(format + " coordinator instructions", context[format].includes("Division Directors and Legislative Coordinators:") && context[format].includes("Microsoft Teams Legislation 91st General Assembly Network."), true);
  check(format + " deadline retained", context[format].includes("Comments due by Friday, 5 PM"), true);
}
var custom = G.buildDailyReport([], { sessionName: "<Custom & Session>", distributionHeading: "Research <staff>" });
check("custom header escaped", custom.html.includes("&lt;Custom &amp; Session&gt;") && custom.html.includes("Research &lt;staff&gt;"), true);
check("generic Teams instruction for custom session", custom.text.includes("legislation channel assigned to your division."), true);

// 3. Round-trip: our own parser reads the generated text (bill lines standalone)
var round = P.parseReport(rep.text, { knownDivisions: ["MVD", "TDD", "SOD"] });
check("round-trip bills", round.items.map(function (i) { return i.billNumber; }), ["HF2801", "SF2110"]);
check("round-trip divisions", round.items[0].distributedTo, ["MVD"]);

// 4. Email extraction from a real pasted To: line
var TO = 'Doe, Jane <Jane.Doe@iowadot.us>; Roe, Bob <bob.roe@iowadot.us>; daniel.brown@dom.iowa.gov; Doe, Jane <JANE.DOE@iowadot.us>';
check("extract emails", G.extractEmails(TO), ["jane.doe@iowadot.us", "bob.roe@iowadot.us", "daniel.brown@dom.iowa.gov"]);
check("extract from junk", G.extractEmails("no addresses here"), []);

// 5. filterNew drops previously-reported bills, case-insensitively
var entries = [{ bill: "HF2801" }, { bill: "SF2110" }, { bill: "HF2803" }];
check("filterNew", G.filterNew(entries, ["hf2801", "SF2110"]).map(function (e) { return e.bill; }), ["HF2803"]);
check("filterNew empty set", G.filterNew(entries, []).length, 3);

for (var unsafe of ["javascript:alert(1)", "data:text/html,evil", "file:///tmp/bill", " https://example.gov"]) {
  var safeReport = G.buildDailyReport([{ billNumber: "HF1", sourceLinks: [{ href: unsafe }] }]);
  check("report rejects unsafe link " + unsafe, safeReport.html.indexOf("<a href=") === -1, true);
}
if (failures) {
  console.error("\n" + failures + " reportgen test(s) FAILED");
  process.exit(1);
}
console.log("All report generator tests passed.");
