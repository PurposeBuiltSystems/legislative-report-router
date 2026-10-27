/* Run real pane refresh/preview/publish guards with offline Graph boundaries. */
"use strict";
const assert = require("node:assert/strict"), fs = require("node:fs"), vm = require("node:vm");
const raw = fs.readFileSync(require("node:path").join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const index = raw.lastIndexOf("})();");
const source = raw.slice(0, index) + `
  renderPreview=function(){self.previews++;};
  setStatus=function(kind,text){self.statuses.push(text);};
  connectRules=async function(){state.rules=(await GraphData.listItems()).map(function(r){return LrrRouting.ruleFromSharePoint(r.fields,r.id);});};
  self.api={state,refreshSharedRecipients,prepareSharedPreview,publish,resetPublicationConfirmation,contactPreview,contactApply,quickAddContact,divisionEmailGroups};
` + raw.slice(index);
function load() {
  const elements = {}, stored = { "lrr.settings": JSON.stringify({ sharedRecipients: true }) };
  const calls = { reads: 0, emails: 0 };
  const rows = [{ id: "1", eTag: '"version1"', fields: { DivisionCode: "MVD", TeamsTeamId: "team", TeamsChannelId: "main", TeamsTagId: "tag", Emails: "bob@agency.gov" } }];
  const context = { self: { previews: 0, statuses: [] },
    Office: { onReady() {}, AsyncResultStatus: { Succeeded: "ok" }, context: { roamingSettings: {
      get: k => stored[k], set: (k,v) => { stored[k]=v; }, saveAsync: cb => cb({ status: "ok" })
    } } },
    document: { getElementById: id => elements[id] || (elements[id]={ checked: false, disabled: false, textContent: "" }) },
    GraphData: {
      getToken: async () => "token",
      getRecipientToken: async () => "token",
      listItems: async () => { calls.reads++; return rows; },
      updateListItemFields: async (t,s,l,id,fields) => { Object.assign(rows[0].fields, fields); },
      listTeamMembers: async () => [{ userId: "alice", email: "alice@agency.gov" }, { userId: "bob", email: "bob@agency.gov" }],
      listTeamTagMembers: async () => [{ userId: "alice" }],
      updateRecipientEmails: async () => { calls.emails++; }
    }, LrrContacts: require("../src/contacts"), LrrRecipients: require("../src/recipients"), LrrRouting: require("../src/routing")
  };
  vm.runInNewContext(source, context);
  context.self.api.state.site = { siteId: "s", routingListId: "l" };
  context.self.api.state.items = [{ billNumber: "HF1", distributedTo: ["MVD"] }];
  return Object.assign(context.self, { elements, graph: context.GraphData, calls, rows });
}
(async () => {
  const g=load();
  await g.api.publish(false);
  assert.match(g.statuses.at(-1), /Open Preview/, "cannot publish unrefreshed shared contacts");
  await g.api.prepareSharedPreview();
  assert.equal(g.calls.reads, 1, "preview reloads spreadsheet edits");
  assert.equal(g.previews, 1, "preview renders only after contacts resolve");
  assert.deepEqual(Array.from(g.api.state.rules[0].emails).sort(), ["alice@agency.gov", "bob@agency.gov"]);
  assert.deepEqual(Array.from(g.api.state.items[0].routes[0].mentionUserIds), ["bob"]);
  assert.equal(g.elements.confirmBox.checked, false, "refresh requires a new review confirmation");
  assert.match(g.elements.recipientSyncInfo.textContent, /All email contacts match/);
  g.graph.listItems = async () => { throw new Error("SharePoint offline"); };
  assert.equal(await g.api.refreshSharedRecipients(), false);
  g.elements.confirmBox.checked = true;
  await g.api.publish(false);
  assert.match(g.statuses.at(-1), /Open Preview/, "failed refresh blocks stale recipients even if confirmation is checked");
  assert.match(g.elements.recipientSyncInfo.textContent, /SharePoint offline/);
  assert.equal(g.elements.recipientSync.disabled, false, "refresh can be retried");
  const paste=load();
  paste.api.state.rules=[require("../src/routing").ruleFromSharePoint(paste.rows[0].fields,"1")];
  paste.elements.contactRows={value:"Carol\tcarol@agency.gov\tMVD"};
  paste.api.contactPreview();
  await paste.api.contactApply();
  assert.match(paste.rows[0].fields.Emails, /carol@agency.gov/, "paste import saves the contact");
  paste.graph.listTeamMembers=async()=>[{userId:"alice",email:"alice@agency.gov"},{userId:"bob",email:"bob@agency.gov"},{userId:"carol",email:"carol@agency.gov"}];
  await paste.api.prepareSharedPreview();
  assert.ok(paste.api.state.rules[0].emails.includes("carol@agency.gov"));
  assert.ok(paste.api.state.rules[0].mentionUserIds.includes("carol"), "pasted email becomes a direct Teams mention");
  const bill={routes:[{id:"1",divisionCode:"MVD"},{id:"2",divisionCode:"MVD"}]};
  const groups=paste.api.divisionEmailGroups([bill]);
  assert.equal(groups.length,1,"multiple channels for MVD produce one shared division email");
  assert.equal(groups[0].items.length,1,"each bill appears once");
  const noSite=load(); noSite.api.state.site=null;
  assert.equal(await noSite.api.refreshSharedRecipients(), false);
  noSite.api.state.site={siteId:"s",routingListId:"l"};
  assert.equal(await noSite.api.refreshSharedRecipients(), true, "early setup failure remains retryable");
  console.log("Shared recipient preview and publication guard tests passed.");
})().catch(e => { console.error(e); process.exitCode=1; });
