/* Offline tests for contact-list import. Run: node test/contacts.test.js */
"use strict";
var C = require("../src/contacts.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

// 1. Excel paste (tabs, header row, multi-division)
var TSV = [
  "Name\tEmail\tDivision",
  "Jane Doe\tJane.Doe@iowadot.us\tMVD",
  "Bob Roe\tbob.roe@iowadot.us\tMVD/TDD",
  "Sue Q\tsue@dot.gov\tAG Office",
].join("\n");
var p1 = C.parseContactRows(TSV);
check("header skipped", p1.skippedHeader, true);
check("row count", p1.rows.length, 3);
check("email lowercased", p1.rows[0].email, "jane.doe@iowadot.us");
check("multi-division split", p1.rows[1].divisions, ["MVD", "TDD"]);
check("multi-word division", p1.rows[2].divisions, ["AG Office"]);
check("no errors", p1.errors.length, 0);

// 2. CSV two-column (email, division), no header
var CSV = "jane@x.gov,MVD\nbadline-no-email,MVD\nbob@x.gov,TDD";
var p2 = C.parseContactRows(CSV);
check("csv rows", p2.rows.length, 2);
check("csv division", p2.rows[0].divisions, ["MVD"]);
check("bad line reported", p2.errors.length, 1);

// 3. Missing division reported, not silently dropped
var p3 = C.parseContactRows("lonely@x.gov");
check("no-division error", p3.errors.length, 1);
check("no-division rows", p3.rows.length, 0);

// 4. Merge plan: dedupe against existing, group by division, alias match
var rules = [
  { id: "1", divisionCode: "MVD", divisionName: "Motor Vehicle Division", aliases: ["Motor Vehicle"],
    emails: ["jane.doe@iowadot.us"] },
  { id: "2", divisionCode: "TDD", divisionName: "", aliases: [], emails: [] },
];
var plan = C.mergePlan(p1.rows, rules);
check("plan groups", plan.map(function (g) { return g.division; }), ["MVD", "TDD", "AG Office"]);
check("existing email not re-added", plan[0].addEmails, ["bob.roe@iowadot.us"]);
check("tdd gets bob", plan[1].addEmails, ["bob.roe@iowadot.us"]);
check("unknown division -> null rule (create)", plan[2].rule, null);
check("unknown division emails", plan[2].addEmails, ["sue@dot.gov"]);

// 5. Alias matching ("Motor Vehicle" rows land on the MVD rule)
var aliasRows = C.parseContactRows("new@x.gov\tMotor Vehicle").rows;
var aliasPlan = C.mergePlan(aliasRows, rules);
check("alias resolves to MVD rule", aliasPlan[0].rule.id, "1");

// 6. Duplicate addresses in the paste collapse
var dupPlan = C.mergePlan(C.parseContactRows("a@x.gov,ELT\na@x.gov,ELT").rows, []);
check("paste dupes collapse", dupPlan[0].addEmails, ["a@x.gov"]);

if (failures) {
  console.error("\n" + failures + " contact test(s) FAILED");
  process.exit(1);
}
console.log("All contact import tests passed.");
