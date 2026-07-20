/* Offline tests for Code-chapter tracking. Run: node test/chapters.test.js */
"use strict";
var C = require("../src/chapters.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

// 1. Extraction from real brief phrasings (all from actual reports)
check("Code chapter 22", C.extractChapters('amends the definition of "government body" in Code chapter 22.'), ["22"]);
check("chapters 6A and 6B + Amends 6A",
  C.extractChapters("Defines a common carrier for purposes of Code chapters 6A and 6B.\nAmends 6A. Sim. SF95."),
  ["6A", "6B"]);
check("Code 307", C.extractChapters("Creates new language in Code 307. Successor to HF882."), ["307"]);
check("Amends Code Chapter 9", C.extractChapters("Amends Code Chapter 9. Successor to HSB747."), ["9"]);
check("section with dot", C.extractChapters("Amends section 321.276 regarding electronic devices."), ["321.276"]);
check("years never match", C.extractChapters("Amends 2026 Iowa Acts, House File 882, section 1."), ["1"]);

// 2. Bill numbers never leak in (letter prefixes fail the token test)
check("no bill leak", C.extractChapters("Successor to HSB171, as amended."), []);

// 3. Tracked matching: chapter, section, range semantics
var tracked = ["321", "232.52", "331.301-331.440"];
check("chapter match", C.matchTracked(["321"], tracked), ["321"]);
check("section under tracked chapter", C.matchTracked(["321.276"], tracked), ["321.276"]);
check("321J is NOT chapter 321", C.matchTracked(["321J"], tracked), []);
check("exact section", C.matchTracked(["232.52"], tracked), ["232.52"]);
check("chapter of exact-section entry not matched", C.matchTracked(["232.10"], tracked), []);
check("range inside", C.matchTracked(["331.400"], tracked), ["331.400"]);
check("range outside", C.matchTracked(["331.500"], tracked), []);

// 4. Default 2015 DOT list works out of the box
check("default has 321J", C.matchTracked(["321J"], null), ["321J"]);
check("default has 452A", C.matchTracked(["452A"], null), ["452A"]);
check("default misses 232.10", C.matchTracked(["232.10"], null), []);
check("default size sane", C.DEFAULT_TRACKED.length >= 100, true);

// 5. Settings textarea parsing
check("parse textarea", C.parseTrackedText("321, 6A\n306B; 452A  707.6A"), ["321", "6A", "306B", "452A", "707.6A"]);

// 6. Chapter-based rule suggestions
var rules = [
  { id: "1", divisionCode: "MVD", codeChapters: ["321", "321A", "322"] },
  { id: "2", divisionCode: "TDD", codeChapters: ["306", "313", "314"] },
  { id: "3", divisionCode: "ELT", codeChapters: [] },
];
var sug = C.suggestRules(["321.276", "6A"], rules);
check("suggestion count", sug.length, 1);
check("suggested division", sug[0].rule.divisionCode, "MVD");
check("suggested via chapter", sug[0].chapters, ["321"]);
check("no chapters no suggestions", C.suggestRules([], rules), []);

if (failures) {
  console.error("\n" + failures + " chapter test(s) FAILED");
  process.exit(1);
}
console.log("All chapter tests passed.");
