/* Offline tests for parser, routing, and Teams payloads. Run: node test/router.test.js */
"use strict";
var P = require("../src/parser.js");
var R = require("../src/routing.js");
var T = require("../src/teams.js");

var failures = 0;
function check(label, actual, expected) {
  var ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (!ok) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

var KNOWN = ["MVD", "TDD", "SOD", "ELT", "AG Office", "FHWA", "IT"];

// ---------- fixture: the representative Daily Bill Report ----------
var SAMPLE = [
  "Daily Bill Report",
  "2025 Legislative Session – 1st Session, 91st Iowa General Assembly",
  "",
  "Distribution List: Executive Leadership Team; Administrative Services;",
  "Systems Operations; IT; Transportation Development; Motor Vehicle;",
  "Field Operations; Division Legislative Coordinators; FHWA; and others.",
  "",
  "Bill Number:",
  "Bill Distributed To:",
  "Bill Comment Requested From:",
  "Bill Brief (as written):",
  "",
  "HF935",
  "",
  "MVD",
  "MVD",
  "medical personnel authorized to withdraw a specimen of blood from a person suspected of operating while intoxicated",
  "Successor to HSB171, as amended.",
  "",
  "HF936",
  "",
  "AG Office",
  "AG Office",
  "open records",
  "amends the definition of “government body” in Code chapter 22.",
  "Successor to HSB192.",
  "",
  "HF938",
  "MVD",
  "MVD",
  "police, fire department, and other emergency vehicles",
  "Successor to HF728.",
  "",
  "HF939",
  "",
  "TDD",
  "TDD",
  "Common Carriers",
  "Defines a common carrier for purposes of Code chapters 6A and 6B.",
  "Amends 6A. Sim. SF95. Successor to HF491.",
  "",
  "HF952",
  "",
  "ELT",
  "ELT",
  "State Agency Contracts",
  "Makes changes to State purchasing contracts.",
  "",
  "HF954",
  "",
  "MVD",
  "MVD",
  "conduct of elections",
  "VOTER REGISTRATION DATABASE PILOT PROGRAM",
  "Requires the department of transportation to share certain records.",
  "",
  "HF955",
  "",
  "ELT",
  "ELT",
  "government ethics of employees of a state agency",
  "Prohibits certain uses of agency identification or email.",
].join("\n");

var res = P.parseReport(SAMPLE, { knownDivisions: KNOWN });

// 1. Boundary detection: exactly 7 entries; internal refs never split
check("bill count", res.items.length, 7);
check("bill numbers", res.items.map(function (i) { return i.billNumber; }),
  ["HF935", "HF936", "HF938", "HF939", "HF952", "HF954", "HF955"]);

// 2. "Successor to HSB171" stays inside HF935's brief, and is a reference
var hf935 = res.items[0];
check("HSB171 not a boundary", hf935.brief.indexOf("Successor to HSB171") !== -1, true);
check("HSB171 captured as reference", hf935.referencedBills, ["HSB171"]);
check("HF935 divisions", hf935.distributedTo, ["MVD"]);
check("HF935 comment from", hf935.commentRequestedFrom, ["MVD"]);
check("HF935 confidence full", hf935.parserConfidence, 1);

// 3. Multi-word division ("AG Office") survives normalization
check("AG Office division", res.items[1].distributedTo, ["AG Office"]);

// 4. No blank line after bill number (HF938) still parses
check("HF938 divisions", res.items[2].distributedTo, ["MVD"]);
check("HF938 refs", res.items[2].referencedBills, ["HF728"]);

// 5. Multiple references in one brief (HF939: SF95 + HF491)
check("HF939 refs", res.items[3].referencedBills, ["SF95", "HF491"]);

// 6. Multi-line brief kept whole (HF954: 3 lines)
check("HF954 brief lines", res.items[5].brief.split("\n").length, 3);

// 7. Brief first line is not eaten as a division (lowercase topic lines)
check("HF952 brief starts with title", res.items[4].brief.indexOf("State Agency Contracts") === 0, true);

// 8. Header preserved, not parsed as an entry
check("header kept", res.header.indexOf("Daily Bill Report") !== -1, true);

// 9. Multi-division strings
check("slash split", P.normalizeDivisions("MVD/TDD"), ["MVD", "TDD"]);
check("semicolon split", P.normalizeDivisions("MVD; SOD"), ["MVD", "SOD"]);
check("and split", P.normalizeDivisions("ELT and MVD"), ["ELT", "MVD"]);
check("oxford split", P.normalizeDivisions("MVD, TDD, and SOD"), ["MVD", "TDD", "SOD"]);

// 10. HTML input: entities, <br> lines, links preserved
var html = "<div>HF101<br><br>MVD/TDD<br>MVD<br>vehicle &amp; trailer titling" +
  '<br>See <a href="https://www.legis.iowa.gov/HF101">HF101 text</a>.</div>';
var hres = P.parseReport(html, { knownDivisions: KNOWN });
check("html bill", hres.items[0].billNumber, "HF101");
check("html multi-division", hres.items[0].distributedTo, ["MVD", "TDD"]);
check("html entity decoded", hres.items[0].brief.indexOf("vehicle & trailer") !== -1, true);
check("html link attached", hres.items[0].sourceLinks[0].href, "https://www.legis.iowa.gov/HF101");

// 11. Missing divisions -> warning + reduced confidence
var bare = P.parseReport("HF200\n\nSome brief text only.", { knownDivisions: KNOWN });
check("bare warning", bare.items[0].parserWarnings.length >= 1, true);
check("bare confidence reduced", bare.items[0].parserConfidence < 1, true);

// 12. Empty report -> warning, no crash
var empty = P.parseReport("Nothing here at all.", {});
check("empty items", empty.items.length, 0);
check("empty warning", empty.warnings.length, 1);

// ---------- routing ----------
var RULES = [
  { id: "1", divisionCode: "MVD", divisionName: "Motor Vehicle Division", aliases: ["Motor Vehicle"],
    emails: ["mvd-leg@dot.example"], teamsTeamId: "T1", teamsChannelId: "C1",
    teamsTagId: "TAG1", teamsTagName: "MVD Legislation", mentionUserIds: [], mentionUserEmails: [],
    isActive: true, priority: 1 },
  { id: "2", divisionCode: "TDD", divisionName: "Transportation Development", aliases: [],
    emails: ["tdd@dot.example"], teamsTeamId: "T1", teamsChannelId: "C2",
    teamsTagId: "TAG2", teamsTagName: "TDD Legislation", mentionUserIds: ["U9"], mentionUserEmails: ["lead@dot.example"],
    isActive: true, priority: 1 },
  { id: "3", divisionCode: "MVD", divisionName: "MVD OLD ROUTE", aliases: [],
    emails: [], teamsTeamId: "T1", teamsChannelId: "C9", teamsTagId: "", teamsTagName: "",
    mentionUserIds: [], mentionUserEmails: [], isActive: true, priority: 0 },
  { id: "4", divisionCode: "SOD", divisionName: "Retired", aliases: [], emails: [],
    teamsTeamId: "", teamsChannelId: "", teamsTagId: "", teamsTagName: "",
    mentionUserIds: [], mentionUserEmails: [], isActive: false, priority: 5 },
];

// 13. Priority: MVD resolves to rule 1, not rule 3
var routed = R.routeItem(JSON.parse(JSON.stringify(hf935)), RULES);
check("matched status", routed.routingStatus, "matched");
check("priority wins", routed.matchedRoutingRules, ["1"]);

// 14. Alias matching
check("alias match", R.rulesForDivision("Motor Vehicle", RULES)[0].id, "1");

// 15. Inactive rules never match
check("inactive excluded", R.rulesForDivision("SOD", RULES).length, 0);

// 16. Unknown division -> partial when others match
var multi = { distributedTo: ["MVD", "AG Office"], commentRequestedFrom: ["MVD"], routingStatus: "unmatched" };
R.routeItem(multi, RULES);
check("partial status", multi.routingStatus, "partially-matched");
check("unknown listed", multi.unknownDivisions, ["AG Office"]);

// 17. Effective dates
var dated = [{ id: "5", divisionCode: "ELT", aliases: [], emails: [], isActive: true, priority: 0,
  effectiveStartDate: "2030-01-01", teamsTeamId: "", teamsChannelId: "", teamsTagId: "",
  mentionUserIds: [], mentionUserEmails: [] }];
check("not yet effective", R.rulesForDivision("ELT", dated, new Date("2026-07-14")).length, 0);
check("effective later", R.rulesForDivision("ELT", dated, new Date("2031-01-01")).length, 1);

// 18. groupByRule consolidates and skips excluded
var itemsForGroup = [
  R.routeItem({ billNumber: "HF1", distributedTo: ["MVD"], commentRequestedFrom: ["MVD"], routingStatus: "unmatched" }, RULES),
  R.routeItem({ billNumber: "HF2", distributedTo: ["MVD", "TDD"], commentRequestedFrom: ["MVD"], routingStatus: "unmatched" }, RULES),
  R.routeItem({ billNumber: "HF3", distributedTo: ["TDD"], commentRequestedFrom: ["TDD"], routingStatus: "excluded" }, RULES),
];
var groups = R.groupByRule(itemsForGroup);
check("group count", groups.length, 2);
check("MVD group size", groups[0].items.length, 2);
check("excluded skipped", groups[1].items.map(function (i) { return i.billNumber; }), ["HF2"]);

// 19. SharePoint field mapping
var spRule = R.ruleFromSharePoint({ Title: "MVD", DivisionCode: "MVD", Aliases: "Motor Vehicle; Motor Vehicle Division",
  Emails: "a@x.gov;b@x.gov", TeamsTagId: "TG", IsActive: true, Priority: "3" }, 12);
check("sp aliases", spRule.aliases, ["Motor Vehicle", "Motor Vehicle Division"]);
check("sp emails", spRule.emails, ["a@x.gov", "b@x.gov"]);
check("sp priority", spRule.priority, 3);
check("sp id", spRule.id, "12");

// ---------- Teams payloads ----------
var payload = T.buildChannelMessage(hf935, [RULES[0]]);

// 20. Real mention entity, matching <at id> marker, correct schema
check("mention entity", payload.mentions[0].mentioned.tag.id, "TAG1");
check("mention text", payload.mentions[0].mentionText, "MVD Legislation");
check("at marker present", payload.body.content.indexOf('<at id="0">MVD Legislation</at>') === 0, true);
check("no escaped at tags", payload.body.content.indexOf("&lt;at") === -1, true);
check("html contentType", payload.body.contentType, "html");

// 21. Tag + user mentions, deduped across rules sharing a channel
var both = T.buildMentions([RULES[0], RULES[1], RULES[0]]);
check("dedupe count", both.mentions.length, 3); // TAG1, TAG2, U9 — no dupes
check("user mention", both.mentions[2].mentioned.user.id, "U9");

// 22. User content is escaped (no HTML injection via brief)
var evil = T.buildChannelMessage({ billNumber: "HF9<script>", brief: "<img src=x>", distributedTo: [], commentRequestedFrom: [], sourceLinks: [] }, [RULES[0]]);
check("bill escaped", evil.body.content.indexOf("HF9&lt;script&gt;") !== -1, true);
check("brief escaped", evil.body.content.indexOf("&lt;img") !== -1, true);

// 23. Idempotency: stable and channel-distinct
var k1 = T.idempotencyKey("report-a", "HF935", "C1");
check("idempotency stable", k1, T.idempotencyKey("report-a", "HF935", "C1"));
check("idempotency channel-distinct", k1 === T.idempotencyKey("report-a", "HF935", "C2"), false);

// 24. Division email consolidates bills, escapes content
var mail = T.buildDivisionEmail(RULES[0], [hf935, res.items[5]]);
check("email to", mail.to, ["mvd-leg@dot.example"]);
check("email lists both bills", mail.html.indexOf("HF935") !== -1 && mail.html.indexOf("HF954") !== -1, true);
check("email subject count", mail.subject.indexOf("(2)") !== -1, true);

if (failures) {
  console.error("\n" + failures + " test(s) FAILED");
  process.exit(1);
}
console.log("All Legislative Report Router tests passed.");
