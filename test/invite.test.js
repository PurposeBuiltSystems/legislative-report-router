/* Offline unit tests for setup-by-invitation. Run: node test/invite.test.js */
"use strict";
var I = require("../src/invite.js");

var failures = 0;
function check(label, actual, expected) {
  if (JSON.stringify(actual) !== JSON.stringify(expected)) {
    failures++;
    console.error("FAIL  " + label + "\n  expected: " + JSON.stringify(expected) + "\n  actual:   " + JSON.stringify(actual));
  }
}

check("subject", I.setupSubject(""), "Legislative Report Router setup");
check("subject with a session label", I.setupSubject("91st GA"),
  "Legislative Report Router setup — 91st GA");

var code = "eyJzaXRlVXJsIjoiaHR0cHM6Ly94In0=";
var html = I.inviteHtml({ code: code, siteLabel: "https://x.sharepoint.com/sites/leg",
  fromName: "Matt Miller", fromEmail: "matt@x.gov" });

check("code round-trips", I.extractSetupCode(html), code);
check("invitation says what to click", html.indexOf("Find my setup") !== -1, true);
check("invitation warns about the one sign-in step",
  html.indexOf("Connect &amp; load routing rules") !== -1, true);
check("invitation names the site", html.indexOf("sites/leg") !== -1, true);
check("invitation names the sender", html.indexOf("Matt Miller") !== -1, true);

// mail clients wrap, re-encode and inject markup between the markers
var mangled = html.replace(code,
  code.slice(0, 5) + '<span class="y">' + code.slice(5, 11) + "</span>\r\n    " + code.slice(11));
check("survives markup and wrapping", I.extractSetupCode(mangled), code);
check("ordinary mail yields nothing", I.extractSetupCode("<p>Bill report attached</p>"), "");
check("unterminated marker yields nothing", I.extractSetupCode("[[LRR-SETUP]]abc"), "");

check("newest usable invitation wins",
  I.pickInvite([
    { receivedDateTime: "2026-03-01", body: I.inviteHtml({ code: "b2xk" }) },
    { receivedDateTime: "2026-08-04", body: I.inviteHtml({ code: "bmV3" }) },
    { receivedDateTime: "2026-09-01", body: "<p>not an invitation</p>" },
  ]).receivedDateTime, "2026-08-04");
check("nothing usable -> null", I.pickInvite([{ receivedDateTime: "1", body: "hi" }]), null);
check("empty inbox -> null", I.pickInvite([]), null);

check("code decodes to real settings",
  JSON.parse(decodeURIComponent(escape(Buffer.from(I.extractSetupCode(html), "base64").toString("binary")))).siteUrl,
  "https://x");

if (failures) {
  console.error("\n" + failures + " invite test(s) FAILED");
  process.exit(1);
}
console.log("All invite tests passed.");
