/* Actual matching UI + list persistence + permission-free shared delivery. */
"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path");
const {JSDOM}=require("jsdom");
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8");
const at=raw.lastIndexOf("})();");
const source=raw.slice(0,at)+`
  setStatus=function(kind,text){self.statuses.push(text);};
  connectRules=async function(){resetPublicationConfirmation();state.rules=(await GraphData.listItems()).map(function(r){return LrrRouting.ruleFromSharePoint(r.fields,r.id);}).filter(function(r){return r.divisionCode;});};
  self.api={state,loadRecipientMatches,saveRecipientMatches,refreshSharedRecipients,rows:function(){return recipientMatchRows;}};
`+raw.slice(at);
function load(){
 const document=new JSDOM(fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.html"),"utf8")).window.document;
 document.getElementById("quickDivision").innerHTML='<option value="1">MVD</option>';
 document.getElementById("contactRows").value="Alice Example\talice@agency.gov\tMVD\nBob Example\tbob@agency.gov\tMVD";
 const rows=[{id:"1",eTag:'"v1"',fields:{DivisionCode:"MVD",TeamsTeamId:"team",TeamsChannelId:"main",TeamsTagId:"mvd",Emails:"bob@agency.gov"}}];
 const stored={},calls={creates:0,writes:0,newScopes:0};
 const graph={
  getToken:async()=>"existing-token",
  getRecipientToken:async()=>{calls.newScopes++;throw new Error("new permission forbidden");},
  listTeamMembers:async()=>{throw new Error("directory read forbidden");},
  listItems:async()=>rows,
  listTeamTags:async()=>[{id:"mvd"},{id:"other"}],
  listTeamTagMembers:async(t,team,tag)=>tag==="mvd"?[{userId:"alice-id",displayName:"Alice Example"}]:[{userId:"bob-id",displayName:"Bob Example"}],
  addListItem:async(t,s,l,fields)=>{calls.creates++;const row={id:"c"+calls.creates,eTag:'"v1"',fields:{...fields}};rows.push(row);return row;},
  updateRecipientEmails:async(t,s,l,id,fields,etag)=>{const row=rows.find(r=>r.id===id);assert.equal(etag,row.eTag);Object.assign(row.fields,fields);row.eTag='"v'+(++calls.writes+1)+'"';}
 };
 const context={self:{statuses:[]},document,GraphData:graph,
  Office:{onReady(){},AsyncResultStatus:{Succeeded:"ok"},context:{roamingSettings:{get:k=>stored[k],set:(k,v)=>{stored[k]=v;},saveAsync:cb=>cb({status:"ok"})}}},
  LrrRecipients:require("../src/recipients"),LrrRouting:require("../src/routing"),LrrContacts:require("../src/contacts")};
 vm.runInNewContext(source,context);
 context.self.api.state.site={siteId:"s",routingListId:"l"};
 return {api:context.self.api,document,rows,stored,calls,graph};
}
(async()=>{
 const g=load();await g.api.loadRecipientMatches();
 assert.equal(g.api.rows().length,2);
 const alice=g.api.rows().find(r=>r.person.value==="alice-id");
 assert.equal(alice.email.value,"alice@agency.gov","paste suggests the name/email match");
 assert.equal(alice.confirmed.checked,false,"suggestions do not count as consent");
 await g.api.saveRecipientMatches();
 assert.equal(g.calls.creates,0,"unchecked suggestions never save");
 assert.match(g.document.getElementById("recipientMatchInfo").textContent,/Confirm at least one/);
 for(const row of g.api.rows())row.confirmed.checked=true;
 await g.api.saveRecipientMatches();
 assert.equal(g.calls.creates,2,"each confirmed person is stored separately");
 assert.equal(g.calls.newScopes,0);
 assert.equal(g.api.state.rules.length,1,"contact records never appear as division routes");
 assert.ok(g.api.state.rules[0].emails.includes("alice@agency.gov"));
 assert.deepEqual(Array.from(g.api.state.rules[0].mentionUserIds),["bob-id"],"Bob gets a direct mention; Alice is covered by MVD tag");
 assert.equal(JSON.parse(g.stored["lrr.settings"]).sharedRecipients,true);
 assert.match(g.document.getElementById("recipientMatchInfo").textContent,/Open Preview/);
 await g.api.loadRecipientMatches();
 assert.ok(g.api.rows().every(r=>r.confirmed.checked),"saved matches are reused without reconfirming unchanged data");
 const row=g.api.rows()[0];row.email.value="changed@agency.gov";
 row.email.dispatchEvent(new g.document.defaultView.Event("input"));
 assert.equal(row.confirmed.checked,false,"editing a saved suggestion requires confirmation again");
 await g.api.loadRecipientMatches();await g.api.saveRecipientMatches();
 assert.equal(g.calls.creates,2,"reusing saved matches does not create duplicate contact rows");
 g.graph.listTeamTagMembers=async()=>[{userId:"new-id",displayName:"New Reviewer"}];
 assert.equal(await g.api.refreshSharedRecipients(),false,"new Teams members need an email match before shared publishing");
 assert.match(g.document.getElementById("recipientSyncInfo").textContent,/New Reviewer/);
 assert.equal(g.calls.newScopes,0);
 console.log("Permission-free matching UI, persistence, confirmation and delivery tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});
