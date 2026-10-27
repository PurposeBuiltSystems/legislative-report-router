/* Real daily-report audience selection, preview refresh, and draft delivery. */
"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path"),{JSDOM}=require("jsdom");
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8"),at=raw.lastIndexOf("})();");
const source=raw.slice(0,at)+`
 renderPreview=function(){self.previews++;};
 refreshStats=renderItems=updateSetupChecklist=updateConnBanner=renderRoutesSummary=function(){};
 setStatus=function(kind,text){self.statuses.push(text);};
 currentDeadline=function(){return {label:"Friday"};};
 self.api={state,dailyReportRecipients,updateDistInfo,prepareSharedPreview,publish,resetPublicationConfirmation};
`+raw.slice(at);
function load(){
 const document=new JSDOM(fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.html"),"utf8")).window.document;
 const settings={distFromDivisions:true,distList:"Leadership <LEAD@agency.gov>; shared@agency.gov"};
 const stored={"lrr.settings":JSON.stringify(settings)},calls={drafts:[],reads:0};
 const rows=[
  {id:"1",fields:{DivisionCode:"MVD",TeamsTeamId:"team",TeamsChannelId:"main",Emails:"mvd@agency.gov; shared@agency.gov",IsActive:true}},
  {id:"2",fields:{DivisionCode:"TDD",Emails:"tdd@agency.gov; SHARED@agency.gov",IsActive:true}},
  {id:"3",fields:{DivisionCode:"Old",Emails:"old@agency.gov",IsActive:false}},
  {id:"4",fields:{DivisionCode:"Expired",Emails:"expired@agency.gov",EffectiveEndDate:"2000-01-01"}},
  {id:"5",fields:{DivisionCode:"Future",Emails:"future@agency.gov",EffectiveStartDate:"2126-01-01"}},
  {id:"contact",fields:{Title:"Contact: not a division",Emails:"not-a-route@agency.gov",Notes:"LRR_CONTACT_MATCH_V1",IsActive:false}}
 ];
 const graph={getToken:async()=>"token",resolveSite:async()=>({siteId:"site",name:"Legislation"}),
  findList:async(t,s,name)=>name==="LegislativeRoutingMatrix"?"routing":"audit",
  listItems:async(t,s,list)=>{calls.reads++;return list==="routing"?rows:[];},
  createDraftMessage:async(token,to,subject,html)=>{calls.drafts.push({to,subject,html});return {id:"draft1"};},
  sendMail:async()=>{throw new Error("must never send mail in a draft-only workflow");}
 };
 const context={self:{statuses:[],previews:0},document,GraphData:graph,
  Office:{onReady(){},AsyncResultStatus:{Succeeded:"ok"},context:{mailbox:{userProfile:{}},roamingSettings:{get:k=>stored[k],set:(k,v)=>{stored[k]=v;},saveAsync:cb=>cb({status:"ok"})}}},
  LrrReportGen:require("../src/reportgen"),LrrRouting:require("../src/routing"),LrrTeams:require("../src/teams")};
 vm.runInNewContext(source,context);
 const api=context.self.api;api.state.site={siteId:"site",routingListId:"routing",auditListId:"audit"};
 api.state.rules=rows.map(r=>context.LrrRouting.ruleFromSharePoint(r.fields,r.id)).filter(r=>r.divisionCode);
 api.state.items=[{billNumber:"HF1",brief:"Bill",distributedTo:["MVD"],commentRequestedFrom:["MVD"]}];api.state.reportKey="report";
 document.getElementById("distFromDivisions").checked=true;document.getElementById("distList").value=settings.distList;
 for(const id of ["optTeams","optEmail","optOriginal"])document.getElementById(id).checked=false;
 document.getElementById("optDailyReport").checked=true;
 return {api,document,stored,calls,rows,graph,self:context.self};
}
(async()=>{
 const g=load();
 assert.deepEqual(Array.from(g.api.dailyReportRecipients()).sort(),["lead@agency.gov","mvd@agency.gov","shared@agency.gov","tdd@agency.gov"]);
 assert.deepEqual(Array.from(g.api.dailyReportRecipients({distFromDivisions:false,distList:"manual@agency.gov"})),["manual@agency.gov"],"manual distribution remains available");
 g.api.updateDistInfo();assert.match(g.document.getElementById("distInfo").textContent,/4 unique/);
 assert.match(g.document.getElementById("distRecipients").value,/tdd@agency.gov/);
 g.document.getElementById("confirmBox").checked=true;await g.api.publish(false);
 assert.equal(g.calls.drafts.length,0,"must review refreshed division recipients before drafting");
 g.rows[1].fields.Emails+="; new@agency.gov";
 await g.api.prepareSharedPreview();assert.equal(g.self.previews,1);assert.ok(g.calls.reads>0);
 g.document.getElementById("confirmBox").checked=true;await g.api.publish(false);
 assert.equal(g.calls.drafts.length,1);
 assert.deepEqual(Array.from(g.calls.drafts[0].to).sort(),["lead@agency.gov","mvd@agency.gov","new@agency.gov","shared@agency.gov","tdd@agency.gov"],"daily draft uses every active division, deduped, plus extras");
 g.api.resetPublicationConfirmation();g.graph.listItems=async()=>{throw new Error("SharePoint offline");};
 await g.api.prepareSharedPreview();g.document.getElementById("confirmBox").checked=true;await g.api.publish(false);
 assert.equal(g.calls.drafts.length,1,"failed recipient refresh cannot create another draft from stale contacts");
 console.log("Daily distribution union, extras, active rules, preview and draft workflow tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
