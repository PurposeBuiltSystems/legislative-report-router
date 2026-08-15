/* Offline tests for bill-status derivation. Run: npm test */
"use strict";
var S = require("../src/billstatus.js");

var failures = 0;
function check(label, actual, expected) {
  if (actual !== expected) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) +
                  "\n  actual:   " + JSON.stringify(actual));
  }
}

// Shaped exactly like the real Iowa feed entries.
var feed = [
  { bill: "HF437",  description: "A bill for an act establishing a center. (Formerly HSB 52.) Effective date: 07/01/2025." },
  { bill: "HF990",  description: "A bill for an act relating to permits. (Formerly HSB 224.)" },
  { bill: "HSB52",  description: "A study bill for an act establishing a center." },
  { bill: "HSB224", description: "A study bill relating to permits." },
  { bill: "HSB685", description: "A study bill nobody moved." },
  { bill: "SF22",   description: "A bill for an act relating to hands-free driving. Effective date: 07/01/2025." },
  { bill: "SF999",  description: "A bill for an act that stalled on the calendar." },
  { bill: "SF888",  description: "A bill for an act that was vetoed by the Governor." },
  // Drafting language, NOT a disposition. A loose /effective date/i match
  // wrongly called 584 real bills enacted on this exact phrasing.
  { bill: "HSB640", description: "A bill for an act relating to wagering, and including effective date provisions." },
  { bill: "SF777",  description: "A bill for an act on taxes, and including effective date and retroactive applicability provisions." },
];
var idx = S.buildIndex(feed);
var DURING = { sessionEnd: "2026-04-30", now: new Date("2026-03-01T12:00:00") };
var AFTER  = { sessionEnd: "2026-04-30", now: new Date("2026-06-01T12:00:00") };
var NO_END = { now: new Date("2026-06-01T12:00:00") };

// --- index ---
check("successor map: HSB52 -> HF437", idx.successorOf.HSB52, "HF437");
check("enacted flagged from effective date", idx.enacted.HF437, true);
check("vetoed flagged", idx.vetoed.SF888, true);
check("a bill with no outcome is not enacted", idx.enacted.SF999, undefined);

// --- enacted / vetoed are terminal regardless of timing ---
check("enacted during session", S.statusFor("HF437", idx, DURING).status, S.STATUS.ENACTED);
check("enacted after session", S.statusFor("HF437", idx, AFTER).status, S.STATUS.ENACTED);
check("vetoed", S.statusFor("SF888", idx, AFTER).status, S.STATUS.VETOED);

// --- study bills: advancing is their only survival ---
check("study bill that advanced, session running", S.statusFor("HSB224", idx, DURING).status, S.STATUS.ADVANCED);
check("study bill inherits its successor's enactment",
  S.statusFor("HSB52", idx, AFTER).status, S.STATUS.ENACTED);
check("study bill names its successor", S.statusFor("HSB52", idx, AFTER).successor, "HF437");
check("study bill with no successor is ACTIVE while the session runs",
  S.statusFor("HSB685", idx, DURING).status, S.STATUS.ACTIVE);
check("study bill with no successor did not advance, once ended",
  S.statusFor("HSB685", idx, AFTER).status, S.STATUS.NOT_ADVANCED);

// --- numbered bills ---
check("numbered bill, no effective date, session running",
  S.statusFor("SF999", idx, DURING).status, S.STATUS.ACTIVE);
check("numbered bill, no effective date, session ended",
  S.statusFor("SF999", idx, AFTER).status, S.STATUS.NOT_ENACTED);

// --- the safety rule: never call a bill dead without a session end ---
check("no sessionEnd => never dead (study)", S.statusFor("HSB685", idx, NO_END).status, S.STATUS.ACTIVE);
check("no sessionEnd => never dead (numbered)", S.statusFor("SF999", idx, NO_END).status, S.STATUS.ACTIVE);

// --- unknown bills are not silently called dead ---
check("bill absent from the feed, session ended", S.statusFor("HF7777", idx, AFTER).status, S.STATUS.UNKNOWN);
check("bill absent from the feed, session running", S.statusFor("HF7777", idx, DURING).status, S.STATUS.ACTIVE);
check("empty bill number", S.statusFor("", idx, AFTER).status, S.STATUS.UNKNOWN);

// --- normalisation: the feed writes "HF 437", trackers often "hf437" ---
check("space-insensitive", S.statusFor("HF 437", idx, AFTER).status, S.STATUS.ENACTED);
check("case-insensitive", S.statusFor("hf437", idx, AFTER).status, S.STATUS.ENACTED);
check("trailing period", S.statusFor("HF437.", idx, AFTER).status, S.STATUS.ENACTED);

// --- isClosed drives the money split ---
check("not-advanced is closed", S.isClosed(S.STATUS.NOT_ADVANCED), true);
check("not-passed is closed", S.isClosed(S.STATUS.NOT_ENACTED), true);
check("vetoed is closed", S.isClosed(S.STATUS.VETOED), true);
check("enacted is NOT closed (it still costs money)", S.isClosed(S.STATUS.ENACTED), false);
check("active is not closed", S.isClosed(S.STATUS.ACTIVE), false);

// --- "including effective date provisions" is drafting language, not enactment ---
check("drafting language does not count as enacted (study bill)",
  S.statusFor("HSB640", idx, AFTER).status, S.STATUS.NOT_ADVANCED);
check("drafting language does not count as enacted (numbered bill)",
  S.statusFor("SF777", idx, AFTER).status, S.STATUS.NOT_ENACTED);
check("index does not flag drafting language", idx.enacted.HSB640, undefined);
check("a real disposition still counts", idx.enacted.HF437, true);

// --- summarize splits the tracker ---
var sum = S.summarize(["HF437", "HSB685", "SF999", "SF888", "HF7777", "HSB224"], idx, AFTER);
check("enacted bucket", sum.enacted.length, 1);
check("closed bucket: HSB685, SF999, SF888, and HSB224 via its successor HF990", sum.closed.length, 4);
check("unknown bucket", sum.unknown.length, 1);
check("active bucket: HSB224 advanced to HF990, which never passed", sum.active.length, 0);

if (failures) { console.error("\n" + failures + " bill-status test(s) failed."); process.exit(1); }
console.log("All bill-status tests passed.");
