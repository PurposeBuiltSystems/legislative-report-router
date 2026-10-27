/* Real watch UI, public source loading, shared marks and failure semantics. */
"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),path=require("node:path"),vm=require("node:vm"),{JSDOM}=require("jsdom");
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8"),at=raw.lastIndexOf("})();");
const source=raw.slice(0,at)+`
 setStatus=function(kind,text){self.statuses.push({kind,text});};
 refreshStats=renderItems=show=function(){};
 self.api={state,routeFilings,refreshAudit,loadFilings,markWatchReviewed,prepareWatchCheck,watchConfig,selection:function(){return feedSelection;}};
`+raw.slice(at);
const FEED_TIME=Date.now();
function load(audit){
 const document=new JSDOM(fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.html"),"utf8")).window.document;
 document.getElementById("stateName").innerHTML='<option value="Iowa">Iowa</option>';
 document.getElementById("trackedChapters").value="8A";
 const saved={},calls=[],changes={summary:"Contracts bill",parent:"Parent bill on agency budgets",late:false,failFeed:false,failHistory:false,failMark:false};
 audit=audit||[];
 const stamp=age=>new Date(FEED_TIME-6*3600000-age*86400000).toUTCString().replace(" GMT","");
 function item(bill,desc,days,linkBill){return '<item><title><![CDATA['+bill+']]></title><description><![CDATA['+desc+']]></description><link>https://www.legis.iowa.gov/legislation/BillBook?ga=91&amp;ba='+(linkBill||bill.replace(/\s/g,""))+'</link><pubDate>'+stamp(days)+'</pubDate></item>';}
 const graph={getToken:async()=>"existing-token",listItems:async()=>{if(changes.failHistory)throw new Error("history offline");return audit;},addListItem:async(t,s,l,fields)=>{
  if(changes.failMark&&fields.Title==="HF2")throw new Error("write denied");
  for(const value of Object.values(fields))if(typeof value==="string")assert.ok(value.length<=255,"checkpoint fits existing short fields");
  audit.push({id:String(audit.length+1),fields:{...fields}});
 }};
 const context={self:{statuses:[]},document,GraphData:graph,Office:{onReady(){},AsyncResultStatus:{Succeeded:"ok"},context:{mailbox:{userProfile:{emailAddress:"coordinator@agency.gov"}},roamingSettings:{get:k=>saved[k],set:(k,v)=>{saved[k]=v;},saveAsync:cb=>cb({status:"ok"})}}},fetch:async url=>{
  calls.push(url);
  if(url.includes("iowa-screening"))return {ok:true,json:async()=>({documents:{"91:SF1":{url:"https://www.legis.iowa.gov/docs/publications/LGI/91/attachments/SF1.html",version:"introduced bill"}}})};
  if(url.includes("iowa-text"))return {ok:true,json:async()=>({text:"An Act relating to agency budgets.",version:"introduced bill",url:"https://www.legis.iowa.gov/docs/publications/LGI/91/attachments/SF1.html"})};
  if(url.includes("Amendments"))return {ok:!changes.failFeed,status:503,text:async()=>'<rss><channel>'+item("H-1 (SF 1)","Amendment filed",0,"SF1")+'</channel></rss>'};
  return {ok:true,text:async()=>'<rss><channel>'+item("HF1",changes.summary,0)+item("HF2","Road bill",0)+item("SF1",changes.parent,10)+item("HF99","Unrelated old bill",10)+(changes.late?item("HF3","Late mirror arrival",2):"")+'</channel></rss>'};
 }};
 for(const [name,file] of Object.entries({LrrRouting:"routing",LrrFeed:"feed",LrrChapters:"chapters",LrrScreening:"screening",LrrPresets:"presets",LrrAmendments:"amendments",LrrWatch:"watch-window"}))context[name]=require("../src/"+file);
 vm.runInNewContext(source,context);context.self.api.state.site={siteId:"s",auditListId:"audit"};
 return {api:context.self.api,self:context.self,document,saved,audit,calls,changes};
}
(async()=>{
 const g=load();await g.api.loadFilings();assert.equal(g.api.selection().length,3);
 assert.match(g.document.getElementById("watchWindowInfo").textContent,/Last successful feed check/);
 g.api.routeFilings();
 g.api.state.items.find(i=>i.billNumber==="HF1").brief="My edited coordinator brief";
 const initialEntries=Array.from(g.api.selection(),e=>e.entry);
 const amendment=g.api.selection().find(e=>e.entry.bill==="H-1");assert.equal(amendment.entry.parentBill.description,"Parent bill on agency budgets");
 assert.ok(g.calls.some(u=>u.includes("iowa-text/91/SF1.json")),"selected amendment can fetch its parent outside the watch window");
 assert.equal(g.calls.some(u=>u.includes("HF99.json")),false,"unrelated old documents are not fetched");
 for(const e of g.api.selection())e.checked=e.entry.bill==="HF1";
 await g.api.markWatchReviewed();assert.equal(g.audit.filter(r=>r.fields.Status==="watch-reviewed").length,1);
 assert.equal(JSON.parse(g.audit.at(-1).fields.SourceSubject).complete,false,"partial selection never marks the whole window reviewed");
 await g.api.loadFilings();assert.deepEqual(Array.from(g.api.selection(),e=>e.entry.bill).sort(),["H-1","HF2"]);
 await g.api.markWatchReviewed();assert.equal(JSON.parse(g.audit.at(-1).fields.SourceSubject).complete,true);
 await g.api.loadFilings();assert.equal(g.api.selection().length,0,"unchanged reviewed versions disappear on repeat check");
 g.document.getElementById("hideReported").checked=false;
 assert.equal((await g.api.prepareWatchCheck(initialEntries,new Date(),g.api.watchConfig())).visible.length,3,"manual replay can show reviewed versions");
 assert.equal((await g.api.prepareWatchCheck(initialEntries,new Date(),g.api.watchConfig(),true)).visible.length,0,"morning checks still hide reviewed versions during manual replay");
 g.document.getElementById("hideReported").checked=true;
 g.changes.summary="Changed contracts bill";g.changes.parent="Changed parent context";g.changes.late=true;
 await g.api.loadFilings();assert.deepEqual(Array.from(g.api.selection(),e=>e.entry.bill).sort(),["H-1","HF1","HF3"],"changed text/context and late arrivals reappear");
 g.api.routeFilings();
 assert.equal(g.api.state.items.length,4,"updated source versions refresh review items without duplicate bill rows");
 assert.equal(g.api.state.items.find(i=>i.billNumber==="HF1").brief,"My edited coordinator brief","coordinator edits are preserved");
 assert.equal(g.api.state.items.find(i=>i.billNumber==="HF1").latestSourceBrief,"Changed contracts bill");
 assert.equal(g.api.state.items.find(i=>i.billNumber==="H-1").parentBill.description,"Changed parent context");
 await g.api.refreshAudit();assert.equal(g.document.getElementById("auditListView").children.length,0,"watch marks never clutter publication history");
 const before=g.saved["lrr.watchWindow"];g.changes.failFeed=true;await g.api.loadFilings();assert.equal(g.saved["lrr.watchWindow"],before,"failed feeds never advance last-successful coverage");
 assert.equal(g.document.getElementById("markWatchReviewed").disabled,true);
 g.changes.failFeed=false;g.changes.failHistory=true;await g.api.loadFilings();assert.equal(g.api.selection().length,4,"unavailable history never silently hides filings");
 assert.equal(g.document.getElementById("markWatchReviewed").disabled,true);assert.match(g.document.getElementById("watchWindowInfo").textContent,/history unavailable/i);
 const shared=load(g.audit);await shared.api.loadFilings();assert.equal(shared.api.selection().length,0,"another coordinator reuses shared exact-version marks");
 const partial=load();await partial.api.loadFilings();partial.changes.failMark=true;await partial.api.markWatchReviewed();
 assert.equal(partial.audit.some(r=>r.fields.Status==="watch-window"),false,"partial write failure never creates a full window checkpoint");
 assert.ok(partial.self.statuses.some(s=>/full window was not marked complete/.test(s.text)));
 console.log("Shared watch UI, repeat checks, late/changed filings, parent enrichment and failure tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
