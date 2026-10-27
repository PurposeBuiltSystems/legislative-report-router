/* Additional-scope actions must honor sign-out and bound broker waits. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const source = fs.readFileSync(require("node:path").join(__dirname, "../src/graph.js"), "utf8");
function load() {
  const timers = new Map(), saved = [], silent = [], popup = [];
  let n = 0, hang = false;
  const pca = {
    acquireTokenSilent: async request => { silent.push(request); return { accessToken: "silent" }; },
    acquireTokenPopup: request => { popup.push(request); return hang ? new Promise(() => {}) : Promise.resolve({ accessToken: "interactive" }); }
  };
  const context = { self: {},
    Office: { context: { roamingSettings: { get: () => true, set: (key, value) => saved.push(value), saveAsync: cb => cb() } } },
    msal: { createNestablePublicClientApplication: async () => pca },
    setTimeout: (cb, ms) => { const id = ++n; timers.set(id, { cb, ms }); return id; },
    clearTimeout: id => timers.delete(id)
  };
  vm.runInNewContext(source, context);
  return { api: context.self.GraphData, timers, saved, silent, popup, hang: () => { hang = true; } };
}
(async () => {
  for (const [name, scope] of [["getTagWriteToken", "TeamworkTag.ReadWrite"], ["getListCreateToken", "Sites.Manage.All"], ["getRecipientToken", "TeamMember.Read.All"]]) {
    const g = load();
    assert.equal(await g.api[name](), "interactive");
    assert.equal(g.silent.length, 0, "signed-out actions require interactive authentication");
    assert.ok(g.popup[0].scopes.includes(scope));
    assert.equal(g.api.isSignedOut(), false);
    assert.deepEqual(g.saved, [false]);
    assert.equal(g.timers.size, 0, "successful authentication clears timers");
    assert.equal(await g.api[name](), "silent");
    const hung = load(); hung.hang();
    const pending = hung.api[name]();
    await new Promise(setImmediate);
    const timer = Array.from(hung.timers.values()).find(t => t.ms === 120000);
    assert.ok(timer, "popup has a timeout");
    timer.cb();
    await assert.rejects(pending, /didn't finish/);
    assert.equal(hung.api.isSignedOut(), true, "failed authentication does not clear sign-out");
  }
  console.log("All additional-scope authentication tests passed.");
})().catch(err => { console.error(err); process.exitCode = 1; });
