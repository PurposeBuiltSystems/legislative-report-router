/* Actual task-pane feed → screen → review and morning draft paths. */
"use strict";
const assert = require("node:assert/strict");
const fs = require("node:fs");
const vm = require("node:vm");
const path = require("node:path");
const raw = fs.readFileSync(path.join(__dirname, "../src/taskpane/taskpane.js"), "utf8");
const at = raw.lastIndexOf("})();");
const source = raw.slice(0, at) + `
  settings = function () { return {autoDaily:true, autoDailyTime:"00:00", distList:"review@example.gov"}; };
  saveSettings = function () {};
  currentDeadline = function () { return {label:"Friday"}; };
  refreshStats = renderItems = show = function () {};
  setStatus = function (kind,text) { self.statuses.push({kind,text}); };
  self.test = { state, loadFilings, routeFilings, maybeAutoDraft, markReported, reportedBills };
` + raw.slice(at);
function load(failAmendments, fullText) {
  const elements = {};
  function element() { return { children: [], checked: true, value: "", appendChild(n) { this.children.push(n); }, addEventListener() {} }; }
  for (const id of ["stateName", "watchDays", "watchTerms", "watchChapters", "hideReported", "loadFilings", "filingsList", "routeFilings", "optDailyReport", "commentWindow", "trackedChapters"]) elements[id] = element();
  elements.stateName.value = "Iowa"; elements.watchDays.value = "3";
  elements.watchTerms.value = "transportation"; elements.commentWindow.value = "48 business hours";
  const today = new Date(Date.now()-6*3600000).toUTCString().replace(" GMT", "");
  function item(title, description) { return `<item><title><![CDATA[${title}]]></title><description><![CDATA[${description}]]></description><link><![CDATA[https://www.legis.iowa.gov/legislation/BillBook?ga=91&ba=HF2500]]></link><pubDate>${today}</pubDate></item>`; }
  const bills = '<rss><channel>' + item("HF 2500", "State agency contracts") + item("HF 100", "Miscellaneous provisions") + '</channel></rss>';
  const amendments = '<rss><channel>' + item("S 5210 (SF 2472)", "Amendment to SF 2472") + '</channel></rss>';
  const requests = [], drafts = [], saved = {};
  const context = {
    self: {statuses:[]}, Date, console,
    document: { getElementById:id=>elements[id], createElement: element },
    Office: {onReady() {}, AsyncResultStatus: {Succeeded:"ok"}, context: {mailbox:{userProfile:{}}, roamingSettings:{get:k=>saved[k], set:(k,v)=>{saved[k]=v;}, saveAsync:cb=>cb({status:"ok"})}}},
    GraphData: {getToken:async()=>"token", createDraftMessage:async(t,to,subject,html)=>{drafts.push({to,subject,html}); return {id:"draft"};}},
    fetch: async url => { requests.push(url); if(url.includes("iowa-text")) return {ok:true,json:async()=>({text:"Intellectual property developed under research contracts.",version:"introduced bill"})}; if(url.includes("iowa-screening")) return fullText ? {ok:true,json:async()=>({documents:{"91:HF2500":{version:"introduced bill"}}})} : {ok:false,status:404}; if(url.includes("Amendments")) return {ok:!failAmendments,status:503,text:async()=>amendments}; return {ok:true,text:async()=>bills}; },
  };
  for (const [global,name] of Object.entries({LrrParser:"parser",LrrFeed:"feed",LrrChapters:"chapters",LrrScreening:"screening",LrrAmendments:"amendments",LrrWatch:"watch-window",LrrPresets:"presets",LrrRouting:"routing",LrrTeams:"teams",LrrReportGen:"reportgen"})) context[global] = require("../src/"+name);
  vm.createContext(context); vm.runInContext(source,context);
  context.self.test.state.rules = [{id:"1",divisionCode:"TEST",isActive:true,keywords:[],codeChapters:[]}];
  return {context, elements, requests, drafts, test:context.self.test};
}
(async () => {
  const manual = load(false);
  await manual.test.loadFilings();
  assert.equal(manual.elements.filingsList.children.length, 3, "all bills and amendments visible, including unmatched: "+JSON.stringify(manual.context.self.statuses));
  manual.test.routeFilings();
  assert.equal(manual.test.state.items.length, 3);
  assert(manual.test.state.items.some(i=>i.billNumber==="S-5210" && i.relatedBill==="SF2472"));
  assert(manual.test.state.items.find(i=>i.billNumber==="HF100").parserWarnings.some(w=>w.includes("manual screening")));
  assert(manual.requests.some(u=>u.includes("Amendments.xml")));
  manual.test.markReported(manual.test.state.items);
  assert.equal(manual.test.reportedBills().length, 0, "unassigned screening rows remain visible next check");
  const enriched = load(false, true);
  await enriched.test.loadFilings();
  enriched.test.routeFilings();
  assert(enriched.test.state.items.find(i=>i.billNumber==="HF2500").screening.reasons.includes("Research and intellectual property"), "per-document full text reaches actual screening workflow");
  assert(enriched.requests.some(u=>u.includes("iowa-text/91/HF2500.json")));
  const morning = load(false);
  morning.test.state.rules[0].keywords=["contracts"];
  await morning.test.maybeAutoDraft();
  assert.equal(morning.test.state.items.length, 3, "morning uses the same screening");
  assert.equal(morning.drafts.length, 1);
  assert(morning.drafts[0].html.includes("S-5210"));
  assert.equal(morning.test.reportedBills().length,0,"draft creation never marks review or delivery complete");
  const failed = load(true);
  await failed.test.loadFilings();
  assert.equal(failed.elements.filingsList.children.length, 0);
  assert(failed.context.self.statuses.some(s=>s.kind==="error" && s.text.includes("Amendments.xml")), "missing amendment feed is visible");
  console.log("All task-pane screening workflow tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
