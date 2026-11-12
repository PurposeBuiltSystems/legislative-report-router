"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const { JSDOM } = require("jsdom");
const Workflow = require("../src/workflow");
const Store = require("../src/workflow-store");

const owner = "owner@agency.test";
const first = "first@agency.test";
const second = "second@agency.test";
const coordinator = "coordinator@agency.test";
const director = "director@agency.test";
const config = {
  bill: "HF123", assembly: "91", session: "2026", label: "Reviewer completion test",
  type: "Fiscal data", due: "2026-11-12", dueSource: "Fictional test date",
  thread: "https://teams.microsoft.com/test", cfo: "cfo@agency.test", submitter: "submit@agency.test",
  divisions: [{ code: "BUREAU", coordinator, director, reviewers: [first, second] }]
};

async function fixture() {
  const rows = [];
  let sequence = 0;
  const graph = {
    listItems: async () => structuredClone(rows),
    addListItem: async (token, site, list, fields) => {
      rows.push({ id: String(++sequence), fields: structuredClone(fields) });
    },
    updateListItemFields: async (token, site, list, id, fields) => {
      Object.assign(rows.find(row => row.id === id).fields, structuredClone(fields));
    }
  };
  const client = Store.client(graph, { siteId: "site", auditListId: "audit" });
  async function act(kind, input = {}, actor = owner) {
    const events = await client.read("token");
    const state = Workflow.project(events);
    const event = Workflow.plan(events, kind, {
      ...input, revision: state ? state.revision : ""
    }, actor, {
      id: "event-" + (++sequence), request: "request", at: new Date().toISOString()
    });
    return client.append("token", event);
  }
  const state = async () => Workflow.project(await client.read("token"));
  const pending = () => rows.filter(row => row.fields.Status === "reminder-pending");
  const recipients = () => pending().map(row => row.fields.EmailRecipients).sort();
  await act("create", { config });
  await client.configureReminders("token", true);
  return { rows, graph, client, act, state, pending, recipients };
}

async function workflowChecks() {
  const f = await fixture();
  assert.deepEqual(f.recipients(), [first, second]);
  assert.equal(new Set(Store.reminderRecords(await f.client.read("token")).map(row => row.key)).size, 2);
  await assert.rejects(() => f.act("reviewer-done", { division: "BUREAU" }, first), /registered response/);
  await f.act("artifact", { division: "BUREAU", version: "v1", links: "https://agency.test/v1" }, first);
  await f.act("reviewer-done", { division: "BUREAU", note: "Reviewed my section" }, first);
  assert.deepEqual(f.recipients(), [second]);
  let state = await f.state();
  assert.deepEqual(Workflow.reviewerProgress(state, "BUREAU").complete, [first]);
  assert.equal(state.divisions.BUREAU.stage, "Drafting");
  await assert.rejects(() => f.act("reviewer-done", { division: "BUREAU" }, first), /already marked complete/);
  await assert.rejects(() => f.act("reviewer-done", { division: "BUREAU" }, coordinator), /assigned reviewer/);
  await assert.rejects(() => f.act("reviewer-done", { division: "BUREAU" }, "outsider@agency.test"), /assigned reviewer/);
  const secondRow = f.pending().find(row => row.fields.EmailRecipients === second);
  secondRow.fields.PublishedBy = "2026-11-12";
  await f.client.refreshReminders("token");
  assert.equal(f.pending().find(row => row.fields.EmailRecipients === second).fields.PublishedBy, "2026-11-12");
  await f.act("reviewer-done", { division: "BUREAU" }, second);
  state = await f.state();
  assert.equal(state.divisions.BUREAU.stage, "Review requested");
  assert.deepEqual(f.recipients(), [coordinator]);
  await f.act("coordinator", { division: "BUREAU" }, coordinator);
  assert.deepEqual(f.recipients(), [director]);
  await f.act("director", { division: "BUREAU" }, director);
  assert.deepEqual(f.recipients(), [config.cfo]);
  await f.act("cfo", {}, config.cfo);
  assert.deepEqual(f.recipients(), [config.submitter]);
  await assert.rejects(() => f.act("reviewer-reopen", { division: "BUREAU" }, first), /Describe/);
  await f.act("reviewer-reopen", { division: "BUREAU", note: "My calculation needs revision" }, first);
  state = await f.state();
  assert.equal(state.cfo, null);
  assert.equal(state.divisions.BUREAU.director, null);
  assert.deepEqual(Workflow.reviewerProgress(state, "BUREAU").complete, [second]);
  assert.deepEqual(f.recipients(), [first]);
  await f.act("artifact", { division: "BUREAU", version: "v2", links: "https://agency.test/v2" }, first);
  assert.deepEqual(f.recipients(), [first, second], "new version requires current checkoffs");
  await f.act("review", { division: "BUREAU" }, first);
  assert.deepEqual(f.recipients(), [second], "existing explicit handoff records the acting reviewer's completion");
  await f.act("coordinator", { division: "BUREAU" }, coordinator);
  assert.deepEqual(f.recipients(), [director], "existing coordinator authority can advance the response");
  await f.act("changes", { division: "BUREAU", note: "Clarify the assumptions" }, director);
  assert.deepEqual(f.recipients(), [first, second]);
  await assert.rejects(() => f.act("reviewer-done", { division: "BUREAU" }, second), /revised response/);
  await f.act("artifact", { division: "BUREAU", version: "v3", links: "https://agency.test/v3", requestReview: true }, first);
  assert.deepEqual(f.recipients(), [second], "confirmed collection marks only its signed-in reviewer complete");
  const before = await f.state();
  const staleEvents = await f.client.read("token");
  const stale = Workflow.plan(staleEvents, "reviewer-done", { division: "BUREAU", revision: before.revision }, second, {
    id: "stale-completion", at: new Date().toISOString()
  });
  await f.act("config", { config: { ...config, divisions: [{ ...config.divisions[0], reviewers: [first] }] } });
  await assert.rejects(() => f.client.append("token", stale), /Shared work changed/);
  assert.deepEqual(f.recipients(), [first], "changed assignments reset checkoffs and retire removed reviewers");
  const commits = f.rows.filter(row => row.fields.Status === "workflow-event");
  for (const row of f.pending()) assert.equal(row.fields.IdempotencyKey, commits.at(-1).id);
}

async function uiChecks() {
  const f = await fixture();
  await f.act("artifact", { division: "BUREAU", version: "v1", links: "https://agency.test/v1" }, first);
  const dom = new JSDOM('<div id="host"></div>', { runScripts: "outside-only", url: "https://router.test/" });
  const win = dom.window;
  Object.defineProperty(win, "crypto", { value: require("node:crypto").webcrypto });
  win.TextEncoder = TextEncoder;
  for (const file of ["workflow", "workflow-store", "workflow-ui"]) {
    win.eval(fs.readFileSync("src/" + file + ".js", "utf8"));
  }
  let actor = first;
  const graph = { ...f.graph, getToken: async () => actor, authenticatedAccount: token => token };
  const site = { siteId: "site", auditListId: "audit" };
  const host = win.document.getElementById("host");
  const view = win.LrrWorkflowUI.mount(host, { graph, site: () => site, staff: true });
  const el = id => win.document.getElementById(id);
  async function choose(kind) {
    el("wfAction").value = kind;
    el("wfAction").dispatchEvent(new win.Event("change"));
  }
  async function save() {
    el("wfConfirm").checked = true;
    el("wfSaveAction").click();
    for (let count = 0; count < 500; count++) {
      await new Promise(resolve => setTimeout(resolve, 2));
      if (!el("wfLoad").disabled) return;
    }
    throw new Error("Save timed out");
  }
  await view.load();
  host.querySelector("[data-work]").click();
  assert.match(host.querySelector('[aria-label="Individual reviewer completion"]').textContent, /0 of 2 complete/);
  await choose("reviewer-done");
  el("wfSaveAction").click();
  assert.match(el("wfInfo").textContent, /confirm/);
  await save();
  assert.match(host.querySelector('[aria-label="Individual reviewer completion"]').textContent, /1 of 2 complete/);
  assert.equal(el("wfAction").querySelector('option[value="reviewer-done"]'), null);
  assert.ok(el("wfAction").querySelector('option[value="reviewer-reopen"]'));
  actor = second;
  await view.load();
  await choose("reviewer-done");
  await save();
  assert.match(host.querySelector('[aria-label="Individual reviewer completion"]').textContent, /2 of 2 complete/);
  assert.match(el("wfDetail").textContent, /Review requested/);
  assert.match(el("wfEvents").textContent, /Reviewer completed their work/);
  assert.deepEqual(f.recipients(), [coordinator]);
  dom.window.close();
}

(async () => {
  await workflowChecks();
  await uiChecks();
  console.log("Individual reviewer completion, per-person reminders, legacy handoffs, versions, reassignment, rework, stale saves and Teams UI checks passed.");
})().catch(error => { console.error(error); process.exitCode = 1; });
