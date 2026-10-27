"use strict";
const assert=require("node:assert/strict"),Budget=require("../src/budget-dashboard"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path");
const rows=[
 {fy:"SFY2027",division:"MVD",bill:"HF1",cost:"$1.2M",severity:"High"},
 {fy:"SFY2027",division:"TDD",bill:"HF2",cost:"(50,000)",severity:"Low",closed:true},
 {fy:"SFY2028",division:"MVD",bill:"HF3",cost:"TBD",severity:"Unknown"},
 {fy:"SFY2028",division:"MVD",bill:"HF4",cost:0,severity:"None"}
];
(async()=>{
 const m=Budget.model(rows);assert.equal(m.total,1150000);assert.equal(m.costs,1200000);assert.equal(m.savings,50000);assert.equal(m.pending,1);assert.equal(m.estimated,3);assert.equal(m.closed,-50000);
 assert.equal(Budget.model(rows,{fy:"SFY2028",division:"MVD"}).pending,1);
 assert.equal(Budget.model(rows,{fy:"SFY2027",division:"TDD"}).total,-50000);
 assert.equal(Budget.model(rows.concat(rows[0])).duplicates,1);
 const rendered=Budget.html(rows,{asOf:"2026-10-27T21:00:00Z"});
 assert.match(rendered,/Snapshot; not an automatic live feed/);assert.match(rendered,/Missing estimates are counted separately/);assert.match(rendered,/Bill details/);
 const evil=Budget.html([{division:'<img src=x onerror=alert(1)>',bill:'<script>bad()</script>',fy:'<x>',cost:'TBD',severity:'__proto__'}]);
 assert.equal(evil.includes('<script>'),false);assert.equal(evil.includes('<img'),false);assert.ok(evil.includes('&lt;script&gt;'));assert.ok(evil.includes('__proto__'));
 const page=Budget.page(rows,{trackerId:"tracker-123",trackerUrl:'https://tenant.sharepoint.com/list?q="x"',asOf:"now"});
 assert.equal(page.name,"lrr-budget-tracker123.aspx");assert.equal(page.description,"LRR_BUDGET_V1:tracker-123");
 assert.equal(page.canvasLayout.horizontalSections[0].columns[0].webparts[0]["@odata.type"],"#microsoft.graph.textWebPart");
 assert.ok(page.canvasLayout.horizontalSections[0].columns[0].webparts[0].innerHtml.includes('&quot;'));
 const source=fs.readFileSync(path.join(__dirname,"../src/graph.js"),"utf8");
 function load(existing){
  const calls=[];const context={self:{},fetch:async(url,opts)=>{
   calls.push({url,opts});
   if(opts.method==="GET")return {ok:true,status:200,json:async()=>({value:existing?[existing]:[]})};
   if(url.endsWith('/publish'))return {ok:true,status:204};
   return {ok:true,status:opts.method==="POST"?201:200,json:async()=>({id:"page1",webUrl:"https://tenant.sharepoint.com/SitePages/budget.aspx"})};
  }};vm.runInNewContext(source,context);return {api:context.self.GraphData,calls};
 }
 const g=load();await g.api.publishBudgetPage("token","site",page);
 assert.deepEqual(g.calls.map(c=>c.opts.method),["GET","POST","POST"]);
 assert.ok(g.calls.at(-1).url.endsWith('/microsoft.graph.sitePage/publish'));
 const h=load({id:"page1",name:page.name,description:page.description});await h.api.publishBudgetPage("token","site",page);
 assert.deepEqual(h.calls.map(c=>c.opts.method),["GET","PATCH","POST"],"refresh updates stable dashboard instead of creating duplicate pages");
 assert.equal(JSON.parse(h.calls[1].opts.body).name,undefined,"updates exclude readonly page name");
 const conflict=load({id:"unrelated",name:page.name,description:"Someone else's page"});
 await assert.rejects(conflict.api.publishBudgetPage("token","site",page),/No page was overwritten/);
 assert.equal(conflict.calls.length,1);
 console.log("Budget dashboard totals, filters, escaping and SharePoint publication tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
