/* Status refresh must use only the configured legislature's feed. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");
const raw = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const at = raw.lastIndexOf("})();");
const source = raw.slice(0, at) + "\nself.test = {state, loadFeedEntries, refreshBillStatuses};\n" + raw.slice(at);
const iowa = '<rss><channel><item><title>HF 1</title><description>Effective date: 07/01/2026.</description></item></channel></rss>';
const minnesota = JSON.stringify({ results: [{ identifier: "HF 1", title: "Transportation review", first_action_date: "2026-01-01" }] });
function load(stateName, mode) {
  const elements = {
    stateName: {value: stateName}, sessionEnd: {value: ""},
    statusRefresh: {}, status: {}
  };
  const requests = [], writes = [];
  let graphCalls = 0;
  const context = {
    self: {}, Office: {onReady() {}},
    document: {getElementById: id => elements[id]},
    LrrPresets: require("../src/presets"), LrrFeed: require("../src/feed"),
    LrrBillStatus: require("../src/billstatus"),
    GraphData: {
      getToken: async () => { graphCalls++; return "token"; },
      addListColumn: async () => {},
      listItems: async () => [{id:"row", fields:{Title:"HF1"}}],
      updateListItemFields: async (token, site, list, id, fields) => writes.push(fields)
    },
    fetch: async url => {
      requests.push(url.split("?")[0]);
      if (url.includes("IowaBills")) { return {ok:true, text:async () => iowa}; }
      if (mode === "network-error") { throw new Error("Network unavailable"); }
      return {ok:mode !== "missing", text:async () => mode === "empty" ? '{"results":[]}' : mode === "malformed" ? "invalid feed" : minnesota};
    }
  };
  vm.runInNewContext(source, context);
  context.self.test.state.site = {siteId:"site", trackerListId:"tracker"};
  return {api:context.self.test, elements, requests, writes, graphCalls:()=>graphCalls};
}
(async () => {
  for (const mode of ["missing", "network-error", "empty", "malformed"]) {
    const g = load("Minnesota", mode);
    await g.api.refreshBillStatuses();
    assert.deepEqual(g.requests, ["../../feeds/openstates-minnesota.json"], "never fall back to Iowa");
    assert.equal(g.writes.length, 0, "unavailable feed cannot change tracker statuses");
    assert.equal(g.graphCalls(), 0);
    assert.match(g.elements.status.textContent, /Couldn't read the bill feed/);
    assert.equal(g.elements.statusRefresh.disabled, false, "allow a later refresh");
  }
  const mn = load("Minnesota", "success");
  await mn.api.refreshBillStatuses();
  assert.equal(mn.writes.length, 1);
  assert.equal(mn.writes[0].BillStatus, require("../src/billstatus").STATUS.ACTIVE,
    "Iowa's enacted HF1 must not override Minnesota's active HF1");
  assert.deepEqual(mn.requests, ["../../feeds/openstates-minnesota.json"]);
  const ia = load("Iowa", "success");
  await ia.api.refreshBillStatuses();
  assert.deepEqual(ia.requests, ["../../feeds/IowaBills.xml"], "Iowa uses its authoritative RSS source");
  assert.equal(ia.writes[0].BillStatus, require("../src/billstatus").STATUS.ENACTED);
  console.log("All legislature-specific status refresh regressions passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
