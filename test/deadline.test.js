/* Offline unit tests for the comment-deadline math and its rendering in
 * Teams posts, division emails, and the generated report. Run:
 *   node test/deadline.test.js */
"use strict";
var DL = require("../src/deadline.js");
var Teams = require("../src/teams.js");
var Gen = require("../src/reportgen.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}
function iso(d) { return d.getFullYear() + "-" + (d.getMonth() + 1) + "-" + d.getDate() + " " + d.getHours() + ":00"; }

/* ---------------------------------------------------------------- the math */

// 2026-07-29 is a Wednesday; 07-31 Friday; 08-03 Monday; 08-04 Tuesday.
var wed3pm = new Date(2026, 6, 29, 15, 0, 0);
var fri3pm = new Date(2026, 6, 31, 15, 0, 0);

check("mid-week: Wed 3pm + 48 = Fri 3pm", iso(DL.addBusinessHours(wed3pm, 48)), "2026-7-31 15:00");
check("weekend skip: Fri 3pm + 48 = Tue 3pm", iso(DL.addBusinessHours(fri3pm, 48)), "2026-8-4 15:00");
check("holiday skip: Fri + 48 w/ Monday holiday = Wed",
  iso(DL.addBusinessHours(fri3pm, 48, { "2026-08-03": true })), "2026-8-5 15:00");
check("24h = next business day", iso(DL.addBusinessHours(fri3pm, 24)), "2026-8-3 15:00");
check("remainder hours: Fri 10pm + 4 rolls past weekend",
  iso(DL.addBusinessHours(new Date(2026, 6, 31, 22, 0, 0), 4)), "2026-8-3 2:00");
check("remainder within a business day: Wed 10am + 4 = Wed 2pm",
  iso(DL.addBusinessHours(new Date(2026, 6, 29, 10, 0, 0), 4)), "2026-7-29 14:00");

/* ------------------------------------------- end-of-business (DOT default) */

check("EOB: Fri 3pm + 48 = Tue 5:00 PM", iso(DL.deadlineFor(fri3pm, 48, {}, "17:00")), "2026-8-4 17:00");
check("EOB: Wed 10am + 48 = Fri 5:00 PM",
  iso(DL.deadlineFor(new Date(2026, 6, 29, 10, 0, 0), 48, {}, "17:00")), "2026-7-31 17:00");
check("EOB + holiday: Fri + 48 w/ Monday holiday = Wed 5:00 PM",
  iso(DL.deadlineFor(fri3pm, 48, { "2026-08-03": true }, "17:00")), "2026-8-5 17:00");
check("EOB label", DL.formatDeadline(DL.deadlineFor(fri3pm, 48, {}, "17:00")), "Tue, Aug 4, 5:00 PM");
check("blank dueTime keeps same-clock-time", iso(DL.deadlineFor(fri3pm, 48, {}, "")), "2026-8-4 15:00");
check("custom due time honored", iso(DL.deadlineFor(fri3pm, 48, {}, "16:30")).slice(0, 9), "2026-8-4 ");

check("parseHolidays mixed separators",
  DL.parseHolidays("2026-09-07, 2026-11-26  2026-11-27; junk 2026-1-1"),
  { "2026-09-07": true, "2026-11-26": true, "2026-11-27": true });
check("isBusinessDay: Saturday", DL.isBusinessDay(new Date(2026, 7, 1), {}), false);
check("isBusinessDay: holiday", DL.isBusinessDay(new Date(2026, 7, 3), { "2026-08-03": true }), false);
check("formatDeadline", DL.formatDeadline(new Date(2026, 7, 4, 15, 0, 0)), "Tue, Aug 4, 3:00 PM");
check("formatDeadline morning/minutes", DL.formatDeadline(new Date(2026, 7, 3, 9, 5, 0)), "Mon, Aug 3, 9:05 AM");

/* ----------------------------------------------------- rendering: Teams */

var item = {
  billNumber: "HF 437",
  title: "Center for intellectual freedom",
  brief: "A bill for an act establishing a center.",
  distributedTo: ["MVD"],
  commentRequestedFrom: ["MVD"],
  sourceLinks: [{ text: "HF 437 (BillBook)", href: "https://legis.iowa.gov/x" }],
};
var rule = { divisionCode: "MVD", divisionName: "Motor Vehicle Division", teamsTagId: "t1", teamsTagName: "MVD Legislation", emails: ["mvd@x.gov"] };

var post = Teams.buildChannelMessage(item, [rule], { deadlineLabel: "Tue, Aug 4, 3:00 PM" });
check("post subject carries deadline", post.subject, "HF 437 — comments due Tue, Aug 4, 3:00 PM");
check("post marked high importance", post.importance, "high");
check("post body shows deadline line", post.body.content.indexOf("Comments due by Tue, Aug 4, 3:00 PM") !== -1, true);
check("mention still present", post.mentions.length, 1);

var postPlain = Teams.buildChannelMessage(item, [rule], {});
check("no deadline -> no subject", postPlain.subject === undefined, true);
check("no deadline -> no importance", postPlain.importance === undefined, true);

var mail = Teams.buildDivisionEmail(rule, [item], { deadlineLabel: "Tue, Aug 4, 3:00 PM" });
check("email subject carries deadline", mail.subject.indexOf("comments due Tue, Aug 4, 3:00 PM") !== -1, true);
check("email body shows deadline banner", mail.html.indexOf("Comments due by Tue, Aug 4, 3:00 PM") !== -1, true);

/* ---------------------------------------------------- rendering: report */

var rep = Gen.buildDailyReport([{ billNumber: "HF 437", brief: "brief", routingStatus: "matched",
  distributedTo: ["MVD"], commentRequestedFrom: [], sourceLinks: [] }],
  { deadlineLabel: "Tue, Aug 4, 3:00 PM", commentWindow: "48 business hours." });
check("report html shows deadline", rep.html.indexOf("Comments due by Tue, Aug 4, 3:00 PM") !== -1, true);
check("report text shows deadline", rep.text.indexOf("Comments due by Tue, Aug 4, 3:00 PM") !== -1, true);
check("report keeps comment-window text", rep.html.indexOf("48 business hours.") !== -1, true);

/* --- regression: end-of-business snap must never shorten the window --- */
check("short window after 5pm rolls forward, never into the past",
  iso(DL.deadlineFor(new Date(2026, 6, 31, 18, 0, 0), 1, {}, "17:00")), "2026-8-3 17:00");
check("8h from Fri 3pm lands the next business day, not the same 5pm",
  iso(DL.deadlineFor(fri3pm, 8, {}, "17:00")), "2026-8-3 17:00");
check("default 48h behaviour unchanged",
  iso(DL.deadlineFor(fri3pm, 48, {}, "17:00")), "2026-8-4 17:00");
check("snap forward within the same day still fine",
  iso(DL.deadlineFor(new Date(2026, 6, 29, 9, 0, 0), 24, {}, "17:00")), "2026-7-30 17:00");

/* --- regression: tracker DueDate uses the LOCAL date, not toISOString() --- */
check("ymd is local", DL.ymd(new Date(2026, 7, 4, 17, 0, 0)), "2026-08-04");
check("ymd pads", DL.ymd(new Date(2026, 0, 5, 9, 0, 0)), "2026-01-05");

if (failures) {
  console.error("\n" + failures + " deadline test(s) FAILED");
  process.exit(1);
}
console.log("All deadline tests passed.");
