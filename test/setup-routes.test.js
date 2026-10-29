"use strict";
const assert = require('node:assert/strict'), fs = require('node:fs'), vm = require('node:vm');
const {JSDOM} = require('jsdom');
const routing = require('../src/routing'), provision = require('../src/provision');
const raw = fs.readFileSync('src/taskpane/taskpane.js', 'utf8'), end = raw.lastIndexOf('})();');
const dom = new JSDOM(fs.readFileSync('src/taskpane/taskpane.html', 'utf8'));
const document = dom.window.document, el = id => document.getElementById(id);
let rows = [], creates = [], updates = [], settings = {};
const context = {
  self: {}, document, Office: {onReady() {}}, LrrRouting: routing, LrrProvision: provision,
  GraphData: {
    getToken: async () => 'token', getListCreateToken: async () => 'token',
    resolveSite: async () => ({siteId: 'site', name: 'Test'}),
    findList: async () => {throw Error('not found');},
    createList: async (t,s,name) => {creates.push(name); return {id:name};},
    listItems: async () => rows,
    addListItem: async (t,s,l,fields) => {rows.push({id:String(rows.length+1),fields:{...fields}});},
    updateListItemFields: async (t,s,l,id,fields) => {
      updates.push(id); Object.assign(rows.find(r=>r.id===id).fields,fields);
    }
  }
};
vm.runInNewContext(raw.slice(0,end)+`
 setStatus=function(){}; connectRules=async function(){};
 saveSettings=function(patch){self.settings(patch);};
 self.api={state,createLists,addRoute};
`+raw.slice(end),context);
context.self.settings = patch => Object.assign(settings,patch);
const a = context.self.api;
function form(code='MVD',channel='channel') {
  el('rbCode').value=code;
  el('rbTeam').innerHTML='<option value="team">Test team</option>';
  el('rbChannel').innerHTML='<option value="'+channel+'">General</option>';
  el('rbTag').innerHTML='<option value="">No tag</option>';
  for(const id of ['rbName','rbEmails','rbChapters','rbKeywords'])el(id).value='';
}
(async()=>{
  el('siteUrl').value='https://example.sharepoint.com/sites/test';
  await a.createLists();
  assert.equal(creates.length,3); assert.equal(rows.length,0,'new setup never inserts sample destinations or recipients');
  const sample=provision.sampleRoutingRule();
  assert.equal(sample.IsActive,false); assert.ok(!sample.Emails);
  assert.equal(routing.rulesForDivision('MVD',[routing.ruleFromSharePoint(sample,'sample')]).length,0);
  a.state.site={siteId:'site',routingListId:'routing'};
  rows=[{id:'legacy',fields:{...sample,IsActive:true,Emails:'coordinator@example.test'}}];
  form(); await a.addRoute();
  assert.equal(rows.length,1,'reuse older unconfigured sample instead of duplicating it');
  assert.equal(rows[0].fields.TeamsChannelId,'channel');
  assert.equal(rows[0].fields.Emails,'coordinator@example.test');
  assert.equal(rows[0].fields.CodeChapters,sample.CodeChapters);
  assert.equal(rows[0].fields.RoutingKeywords,sample.RoutingKeywords);
  assert.equal(rows[0].fields.DivisionName,'Motor Vehicle Division');
  assert.equal(rows[0].fields.IsActive,true);assert.equal(settings.lastCheckOk,'');
  rows.unshift({id:'inactive-sample',fields:sample});
  a.state.rules=[]; // stale/empty local cache must not create another route
  form('mvd'); await a.addRoute();
  assert.equal(rows.length,2);assert.equal(updates.at(-1),'legacy','exact destination takes precedence over an old sample');
  assert.equal(rows[0].fields.IsActive,false);
  rows=rows.filter(r=>r.id==='legacy');
  form('MVD','other-channel');await a.addRoute();
  assert.equal(rows.length,2,'intentional routing to another configured channel remains possible');
  form('TDD');await a.addRoute();assert.equal(rows.length,3,'new division creates a real route');
  const template=fs.readFileSync('templates/routing-matrix-template.csv','utf8');
  assert.ok(!template.includes('@'),'downloaded examples contain no email addresses');
  assert.equal((template.match(/FALSE/g)||[]).length,2);
  console.log('Setup routes: empty provisioning, inactive examples, legacy reuse, repeated saves and separate destinations passed.');
  dom.window.close();
})().catch(e=>{console.error(e);process.exitCode=1;});
