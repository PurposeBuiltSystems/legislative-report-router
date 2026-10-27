"use strict";
const assert = require("node:assert/strict");
const Recipients = require("../src/recipients");
const Teams = require("../src/teams");
const member = (userId, email) => ({ userId, email, displayName: userId });
function load() {
  const calls = { teams: 0, tags: 0, writes: [] };
  const graph = {
    async listTeamMembers() { calls.teams++; return [member("alice", "Alice@agency.gov"), member("bob", "bob@agency.gov"), member("carol", "carol@agency.gov")]; },
    async listTeamTagMembers(token, team, tag) { calls.tags++; return [{ userId: tag === "tag2" ? "carol" : "alice" }]; },
    async updateRecipientEmails(t, s, l, id, fields) { calls.writes.push({ id, fields }); }
  };
  const rule = (id, emails = []) => ({ id, divisionCode: "MVD", teamsTeamId: "team", teamsChannelId: "main", teamsTagId: "tag", emails });
  return { graph, calls, rule };
}
(async () => {
  const g = load(), rules = [g.rule("1", ["BOB@agency.gov", "bob@agency.gov"]), g.rule("2")];
  const result = await Recipients.reconcile(g.graph, "token", rules, { siteId: "s", routingListId: "l" });
  assert.equal(g.calls.teams, 1); assert.equal(g.calls.tags, 1);
  assert.equal(result.warnings.length, 0);
  assert.deepEqual(rules[0].emails.sort(), ["alice@agency.gov", "bob@agency.gov"]);
  assert.deepEqual(rules[1].emails.sort(), rules[0].emails, "same division routes share contacts");
  assert.deepEqual(rules[0].mentionUserIds, ["bob"], "tag already notifies Alice; Bob receives a direct mention");
  const payload = Teams.buildChannelMessage({ billNumber: "HF1", brief: "Bill", distributedTo: ["MVD"] }, rules, {});
  assert.equal(payload.mentions.length, 2, "one tag and one direct mention, deduped across routes");
  assert.equal(g.calls.writes.length, 2);
  await Recipients.reconcile(g.graph, "token", rules, { siteId: "s", routingListId: "l" });
  assert.equal(g.calls.writes.length, 2, "repeat refresh does not rewrite unchanged addresses");
  assert.deepEqual(rules[0].mentionUserIds, ["bob"], "repeat refresh does not duplicate mentions");

  const h = load(), mvd = h.rule("1", ["bob@agency.gov"]), tdd = { ...h.rule("3"), divisionCode: "TDD", teamsTagId: "tag2" };
  await Recipients.reconcile(h.graph, "token", [mvd, tdd], {});
  assert.deepEqual(tdd.emails, ["carol@agency.gov"], "different divisions never share contacts just because their channel matches");
  assert.equal(mvd.emails.includes("carol@agency.gov"), false);

  const outside = load(), r = outside.rule("1", ["external@example.gov"]);
  const warn = await Recipients.reconcile(outside.graph, "token", [r], {});
  assert.match(warn.warnings[0], /external@example.gov receives email only/);
  assert.equal(r.mentionUserIds.length, 0, "never invent a user ID or add someone to the Team");
  assert.ok(r.emails.includes("external@example.gov"));

  const broken = load();
  broken.graph.listTeamTagMembers = async () => { throw new Error("403 denied"); };
  await assert.rejects(Recipients.reconcile(broken.graph, "token", [broken.rule("1")], {}), /403/);
  assert.equal(broken.calls.writes.length, 0, "source read failure cannot silently produce a partial audience");
  broken.graph.listTeamTagMembers = async () => [{ userId: "unknown", displayName: "Missing person" }];
  await assert.rejects(Recipients.reconcile(broken.graph, "token", [broken.rule("1")], {}), /Missing person/);
  assert.equal(broken.calls.writes.length, 0);

  const denied = load(), original = denied.rule("1", ["bob@agency.gov"]);
  denied.graph.updateRecipientEmails = async () => { throw new Error("write denied"); };
  await assert.rejects(Recipients.reconcile(denied.graph, "token", [original], {}), /write denied/);
  assert.deepEqual(original.emails, ["bob@agency.gov"], "failed persistence does not activate an unreviewed audience");
  console.log("Shared contact reconciliation tests passed.");
})().catch(e => { console.error(e); process.exitCode = 1; });
