/* Exercise actual task-pane publication with mocked Outlook/Graph boundaries. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const Teams = require("../src/teams.js");
const Routing = require("../src/routing.js");
const raw = fs.readFileSync(require("node:path").join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const at = raw.lastIndexOf("})();");
const source = raw.slice(0, at) + `
  settings = function () { return { distList: "review@example.gov" }; };
  currentDeadline = function () { return { label: "Friday", date: new Date(2026, 9, 16) }; };
  markReported = function () {};
  restMessageId = async function () { return "original"; };
  setStatus = function (kind, text) { self.statuses.push({kind, text}); };
  self.test = { state: state, publish: publish };
` + raw.slice(at);
function load(options = {}) {
  const elements = {};
  for (const id of ["confirmBox", "optTeams", "optEmail", "optOriginal", "optDailyReport", "publishGo", "retryFailed", "publishLog", "commentWindow"]) {
    elements[id] = { checked: id === "confirmBox" || id === "optEmail", value: "48 business hours", innerHTML: "", lines: [], appendChild(n) { this.lines.push(n.textContent); } };
  }
  const calls = { sends: [], drafts: 0, originals: 0, audits: 0, posts: 0 };
  let fail = true;
  const context = {
    self: { statuses: [] }, Office: { onReady() {}, context: { mailbox: { userProfile: {} } } },
    document: { getElementById: id => elements[id], createElement: () => ({}) },
    GraphData: {
      getToken: options.getToken || (async () => "token"),
      listItems: async () => [], addListItem: async () => { calls.audits++; },
      sendMail: async (token, to) => { calls.sends.push(to[0]); if (fail && to[0] === "fail@example.gov") throw new Error("Denied"); },
      createDraftMessage: async () => { calls.drafts++; if (fail) throw new Error("Draft denied"); return { id: "daily" }; },
      sendDraft: async () => { calls.originals++; if (fail) throw new Error("Send denied"); },
      postChannelMessage: async () => { calls.posts++; if (fail) throw new Error("Post denied"); return { id: "post" }; }
    }, LrrTeams: Teams, LrrRouting: Routing,
    LrrReportGen: { extractEmails: () => ["review@example.gov"], buildDailyReport: () => ({subject: "Daily", html: "Bills", count: 1}) }
  };
  vm.runInNewContext(source, context);
  const { state, publish } = context.self.test;
  const rules = ["ok", "fail"].map(id => ({ id, divisionCode: id, emails: [id + "@example.gov"], teamsTeamId: "team", teamsChannelId: id }));
  state.site = { siteId: "site", auditListId: "audit" };
  state.reportKey = "report";
  state.items = [{ billNumber: "HF 1", brief: "Bill", routes: rules }];
  return { publish, state, elements, calls, statuses: context.self.statuses, succeed: () => { fail = false; } };
}
(async () => {
  const g = load();
  g.elements.optTeams.checked = g.elements.optDailyReport.checked = g.elements.optOriginal.checked = true;
  await g.publish(false);
  assert.deepEqual(g.calls.sends, ["ok@example.gov", "fail@example.gov"]);
  assert.equal(g.elements.retryFailed.hidden, false);
  g.succeed();
  await g.publish(true);
  assert.deepEqual(g.calls.sends, ["ok@example.gov", "fail@example.gov", "fail@example.gov"], "retry sends only failed emails");
  assert.equal(g.calls.posts, 4, "two failed channel posts retried");
  assert.equal(g.calls.drafts, 2, "failed daily draft retried");
  assert.equal(g.calls.originals, 2, "failed original send retried");
  assert.equal(g.elements.retryFailed.hidden, true);
  await g.publish(true);
  assert.equal(g.calls.sends.length, 3, "successful retries are never sent twice");
  assert.equal(g.calls.posts, 4);
  assert.equal(g.calls.drafts, 2);
  assert.equal(g.calls.originals, 2);
  await g.publish(false);
  assert.equal(g.calls.sends.length, 3, "normal publish also skips recorded email successes");
  assert.equal(g.calls.drafts, 2, "normal publish preserves this session's daily draft");
  assert.equal(g.calls.originals, 2);
  let resolveToken;
  const busy = load({ getToken: () => new Promise(resolve => { resolveToken = resolve; }) });
  const pending = busy.publish(false);
  assert.equal(busy.elements.publishGo.disabled, true);
  assert.equal(busy.elements.retryFailed.disabled, true);
  await busy.publish(true);
  resolveToken("token");
  await pending;
  assert.equal(busy.calls.sends.length, 2, "overlapping clicks start one publication only");
  assert.equal(busy.state.publishing, false);
  assert.equal(busy.elements.retryFailed.disabled, false);
  const unconfirmed = load();
  unconfirmed.elements.confirmBox.checked = false;
  await unconfirmed.publish(true);
  assert.equal(unconfirmed.calls.sends.length, 0, "retry also requires the review confirmation");
  console.log("All publication and retry regression tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
