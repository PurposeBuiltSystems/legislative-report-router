"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path"),{JSDOM}=require("jsdom");
const document=new JSDOM(fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.html"),"utf8")).window.document;
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8"),at=raw.lastIndexOf("})();");
const source=raw.slice(0,at)+`
 setStatus=function(){};
 self.api={state,renderBudgetDashboard,publishBudgetDashboard,budgetFilters,
 setRows:function(rows){fiscalRows=rows;fiscalSnapshot={site:state.site,at:"2026-10-27T21:00:00Z"};}};
`+raw.slice(at);
const calls=[];
const context={self:{},document,Office:{onReady(){}},LrrBudget:require("../src/budget-dashboard"),GraphData:{
 getToken:async()=>"existing-token",trackerWebUrl:async()=>"https://tenant.sharepoint.com/List",
 publishBudgetPage:async(token,site,page)=>{calls.push({site,page});return {id:"p",webUrl:"https://tenant.sharepoint.com/SitePages/budget.aspx"};}
}};
vm.runInNewContext(source,context);
(async()=>{
 const api=context.self.api;
 api.state.site={siteId:"s",trackerListId:"tracker"};
 await api.publishBudgetDashboard();assert.equal(calls.length,0,"must load data before publication");
 api.setRows([{fy:"SFY2027",division:"MVD",bill:"HF1",cost:"100000",severity:"High"},{fy:"SFY2028",division:"TDD",bill:"HF2",cost:"200000",severity:"Low"}]);
 api.budgetFilters();document.getElementById("budgetFy").value="SFY2027";api.renderBudgetDashboard();
 assert.match(document.getElementById("budgetDashboard").textContent,/\$100,000/);
 assert.equal(document.getElementById("budgetDashboard").textContent.includes("HF2"),false,"in-router filters update details");
 await api.publishBudgetDashboard();assert.equal(calls.length,1);
 assert.match(calls[0].page.canvasLayout.horizontalSections[0].columns[0].webparts[0].innerHtml,/HF2/,"published snapshot explicitly contains every loaded year/division");
 const link=document.getElementById("budgetPublishInfo").querySelector("a");assert.equal(link.href,"https://tenant.sharepoint.com/SitePages/budget.aspx");
 assert.match(document.getElementById("budgetPublishInfo").textContent,/approval/);
 api.state.site={siteId:"different",trackerListId:"other"};await api.publishBudgetDashboard();
 assert.equal(calls.length,1,"never publish stale estimates to a newly connected site");
 console.log("Budget dashboard filtering and publication workflow tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
