/* Offline unit tests for the fiscal harvest analyzer. Run: node test/harvest.test.js */
"use strict";
var H = require("../src/harvest.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

/* --------------------------------------------------------------- findMoney */

check("dollar with M", H.findMoney("We estimate $1.2M in new costs.").map(function (h) { return h.amount; }), [1200000]);
check("comma number", H.findMoney("total cost of $250,000 annually").map(function (h) { return h.amount; }), [250000]);
check("word multiplier without $", H.findMoney("roughly 1.5 million per year").map(function (h) { return h.amount; }), [1500000]);
check("parenthesized savings", H.findMoney("a reduction of ($50,000) in SFY27").map(function (h) { return h.amount; }), [-50000]);
check("small fee noise filtered", H.findMoney("a $5 fee increase").length, 0);
check("multiple figures found", H.findMoney("$300k in FY27 and $450,000 in FY28").length, 2);
check("no money -> empty", H.findMoney("No fiscal impact anticipated."), []);

/* --------------------------------------------------------------- bestMoney */

var picked = H.bestMoney(H.findMoney(
  "The registration fee is $1,000 per unit. Total estimated cost: $2.4 million annually."));
check("prefers total/cost context and larger figure", picked.amount, 2400000);
check("context captured", picked.context.indexOf("Total estimated cost") !== -1, true);

/* ------------------------------------------------------------ severity/FY */

check("no impact", H.findSeverity("We reviewed HF 437 — no fiscal impact."), "None");
check("minimal", H.findSeverity("Impact is minimal for MVD."), "Low");
check("moderate", H.findSeverity("This is a moderate burden."), "Moderate");
check("significant", H.findSeverity("Significant cost increase expected"), "High");
check("critical", H.findSeverity("This is critical for operations"), "Critical");
check("nothing", H.findSeverity("We will review next week."), "");

check("SFY2027", H.findFy("costs land in SFY2027"), "SFY2027");
check("FY27 short", H.findFy("beginning FY27"), "SFY2027");
check("fiscal year long", H.findFy("in fiscal year 2028 and beyond"), "SFY2028");
check("no fy", H.findFy("soon"), "");

/* ------------------------------------------------------------ analyze/merge */

var replyOnly = H.analyzeSource({
  bill: "HF 437", division: "MVD", source: "Teams reply", author: "Bob",
  date: "2026-02-01T10:00:00Z",
  text: "Minimal impact for us — maybe $40,000 in FY27 for signage.",
});
check("reply candidate cost", replyOnly.cost, 40000);
check("reply candidate severity", replyOnly.severity, "Low");
check("reply candidate fy", replyOnly.fy, "SFY2027");
check("evidence quotes context", replyOnly.evidence.indexOf("$40,000") !== -1, true);

check("no-signal text -> null", H.analyzeSource({ bill: "X", text: "Thanks, will look Monday." }), null);
check("severity-only still a candidate",
  H.analyzeSource({ bill: "X", division: "D", text: "No fiscal impact." }).severity, "None");

var attach = H.analyzeSource({
  bill: "HF 437", division: "MVD", source: "HF437-estimate.xlsx", kind: "attachment", author: "Bob",
  date: "2026-02-02T10:00:00Z",
  text: "Line\tAmount\nStaff\t120000\nTotal estimated cost\t$310,000",
});
var merged = H.mergeCandidates([replyOnly, attach]);
check("one candidate per bill+division", merged.length, 1);
check("attachment cost outranks prose commentary", merged[0].source, "HF437-estimate.xlsx");
check("winner cost from attachment", merged[0].cost, 310000);
check("fy carried from the reply when attachment lacks one", merged[0].fy, "SFY2027");
check("severity carried from the reply into the winner", merged[0].severity, "Low");

var twoDivisions = H.mergeCandidates([
  H.analyzeSource({ bill: "HF 1", division: "MVD", text: "$10,000 cost", date: "1" }),
  H.analyzeSource({ bill: "HF 1", division: "TDD", text: "$20,000 cost", date: "2" }),
]);
check("different divisions stay separate", twoDivisions.length, 2);

/* -------------------------------------------------------------- htmlToText */

check("teams html reply flattens",
  H.htmlToText("<p>Roughly <b>$1.2M</b> annually.</p><p>&mdash; Bob</p>").indexOf("$1.2M annually.") !== -1,
  true);

if (failures) {
  console.error("\n" + failures + " harvest test(s) FAILED");
  process.exit(1);
}
console.log("All harvest tests passed.");
