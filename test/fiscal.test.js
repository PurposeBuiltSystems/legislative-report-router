/* Offline unit tests for the fiscal-impact rollup. Run: node test/fiscal.test.js */
"use strict";
var F = require("../src/fiscal.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

/* ------------------------------------------------------------- SFY default */

check("Feb session bill -> SFY of upcoming Jul 1", F.defaultFy(new Date(2026, 1, 15)), "SFY2027");
check("Jun edge -> same upcoming Jul 1", F.defaultFy(new Date(2026, 5, 30)), "SFY2027");
check("Jul publish -> the NEXT Jul 1's SFY", F.defaultFy(new Date(2026, 6, 31)), "SFY2028");
check("Dec -> next Jul 1's SFY", F.defaultFy(new Date(2026, 11, 1)), "SFY2028");

/* -------------------------------------------------------------- parseCost */

check("$1.2M", F.parseCost("$1.2M"), 1200000);
check("250,000", F.parseCost("250,000"), 250000);
check("300k", F.parseCost("300k"), 300000);
check("1.5 million", F.parseCost("1.5 million"), 1500000);
check("plain number", F.parseCost(4200), 4200);
check("parenthesized = savings", F.parseCost("(50,000)"), -50000);
check("negative", F.parseCost("-75000"), -75000);
check("savings word", F.parseCost("500k savings"), -500000);
check("unknown -> null", F.parseCost("Unknown"), null);
check("TBD -> null", F.parseCost("tbd"), null);
check("blank -> null", F.parseCost(""), null);
check("null never zero", F.parseCost("n/a"), null);

check("fmtMoney", F.fmtMoney(1234567), "$1,234,567");
check("fmtMoney negative", F.fmtMoney(-50000), "−$50,000");

/* -------------------------------------------------------------- aggregate */

var rows = [
  { fy: "SFY2027", division: "MVD", bill: "HF 437", severity: "High", cost: "$1.2M" },
  { fy: "SFY2027", division: "MVD", bill: "SF 2103", severity: "Low", cost: "250,000" },
  { fy: "SFY2027", division: "TDD", bill: "HF 437", severity: "Moderate", cost: "(200,000)" },
  { fy: "SFY2027", division: "TDD", bill: "HF 2630", severity: "Unknown", cost: "TBD" },
  { fy: "SFY2028", division: "MVD", bill: "HF 999", severity: "Critical", cost: "2M" },
];
var agg = F.aggregate(rows);
check("newest FY first", agg.map(function (f) { return f.fy; }), ["SFY2028", "SFY2027"]);
check("SFY2027 total nets savings", agg[1].total, 1250000);
check("SFY2027 estimated/unestimated", [agg[1].estimated, agg[1].unestimated], [3, 1]);
check("SFY2027 distinct bills", agg[1].billCount, 3);
check("division totals", agg[1].byDivision, { MVD: 1450000, TDD: -200000 });
check("severity counts", agg[1].bySeverity, { High: 1, Low: 1, Moderate: 1, Unknown: 1 });
check("TBD not counted as zero-dollar estimate", agg[1].estimated, 3);

/* -------------------------------------------------------------------- CSV */

var csv = F.rollupCsv(rows);
var lines = csv.trim().split("\r\n");
check("csv header", lines[0], "fiscal_year,division,bill,severity,estimated_cost,notes");
check("csv sorted newest FY first", lines[1].indexOf("SFY2028") === 0, true);
check("csv TBD cost blank", lines.some(function (l) { return l.indexOf("SFY2027,TDD,HF 2630,Unknown,,") === 0; }), true);
check("csv TOTAL lines per FY",
  lines.filter(function (l) { return l.indexOf(",TOTAL,") !== -1; }).length, 2);
check("csv SFY2027 total value", lines.some(function (l) { return l === "SFY2027,TOTAL,,,1250000,"; }), true);

if (failures) {
  console.error("\n" + failures + " fiscal test(s) FAILED");
  process.exit(1);
}
console.log("All fiscal tests passed.");
