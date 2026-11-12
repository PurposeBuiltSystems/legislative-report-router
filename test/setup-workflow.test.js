/* Exercise the actual wizard and tag roster against offline boundaries. */
"use strict";
const assert=require("node:assert/strict"),fs=require("node:fs"),vm=require("node:vm"),path=require("node:path");
const html=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.html"),"utf8");
const raw=fs.readFileSync(path.join(__dirname,"../src/taskpane/taskpane.js"),"utf8"),at=raw.lastIndexOf("})();");
const hooks=[
"loadSitesPicker=loadTeamSitePicker=loadTeamsPicker=function(){};",
"renderAuthState=function(){};connectRules=async function(){};",
"self.workflow={state,advanceSetup,wizShow,loadTagMembers,addRoute,resetPublicationConfirmation,publish,getStep:function(){return wizStep;}};"
].join("\n");
const source=raw.slice(0,at)+hooks+raw.slice(at);
function load(){
  function element(){return {value:"",checked:false,disabled:false,hidden:false,open:true,innerHTML:"",textContent:"",
    children:[],querySelectorAll:()=>[],focus(){},selectedOptions:[{textContent:"Chosen channel"}],classList:{toggle(){}},
    appendChild(n){this.children.push(n);},addEventListener(){}};}
  const elements=Object.fromEntries(Array.from(html.matchAll(/id="([^"]+)"/g),m=>[m[1],element()]));
  const stored={},calls={added:[],updated:[],members:0};
  const graph={getToken:async()=>"token",listTeamTagMembers:async()=>{calls.members++;return [];},
    addListItem:async(t,s,l,fields)=>calls.added.push(fields),
    updateListItemFields:async(t,s,l,id,fields)=>calls.updated.push({id,fields})};
  const context={self:{},GraphData:graph,
    Office:{onReady(){},AsyncResultStatus:{Succeeded:"ok"},context:{roamingSettings:{
      get:k=>stored[k],set:(k,v)=>{stored[k]=v;},saveAsync:cb=>cb({status:"ok"})}}},
    document:{getElementById:id=>elements[id]||null,querySelectorAll:()=>[],createElement:element},
    LrrRouting:require("../src/routing"),LrrTeams:require("../src/teams"),LrrReportGen:require("../src/reportgen")};
  vm.runInNewContext(source,context);
  return {api:context.self.workflow,elements,stored,graph,calls};
}
(async()=>{
  const g=load(),w=g.api,e=g.elements;
  w.wizShow(1);await w.advanceSetup();assert.equal(w.getStep(),2);
  await w.advanceSetup();assert.equal(w.getStep(),2,"connect lists before routes");
  w.state.site={siteId:"site",routingListId:"routes",auditListId:"audit"};
  await w.advanceSetup();assert.equal(w.getStep(),3);
  await w.advanceSetup();assert.equal(w.getStep(),3,"a real destination is required");
  w.state.rules=[{id:"1",divisionCode:"MVD",teamsTeamId:"team",teamsChannelId:"channel",isActive:true}];
  await w.advanceSetup();assert.equal(w.getStep(),4);
  await w.advanceSetup();assert.equal(w.getStep(),5,"email recipients are optional");
  e.autoDaily.checked=true;
  await w.advanceSetup();assert.equal(w.getStep(),5,"scheduling needs recipients");
  e.distList.value="review@example.gov";await w.advanceSetup();assert.equal(w.getStep(),6);
  await w.advanceSetup();assert.equal(e.settings.open,true,"checks required before Done");
  g.stored["lrr.settings"]=JSON.stringify({lastCheckOk:new Date().toISOString()});
  await w.advanceSetup();assert.equal(e.settings.open,false,"Done closes setup");
  assert.equal(e["screen-overview"].hidden,false);
  const denied=load();denied.api.wizShow(1);
  denied.graph.getToken=async()=>{throw new Error("Sign in required");};
  await denied.api.advanceSetup();assert.equal(denied.api.getStep(),1);

  const tags=load();await tags.api.loadTagMembers();assert.equal(tags.calls.members,0);
  tags.elements.rbTeam.value="team";tags.elements.rbTag.value="A";
  let resolveA,resolveB;
  tags.graph.listTeamTagMembers=async(t,team,tag)=>new Promise(resolve=>{
    if(tag==="A")resolveA=resolve;else resolveB=resolve;
  });
  const a=tags.api.loadTagMembers();await new Promise(setImmediate);
  tags.elements.rbTag.value="B";
  const b=tags.api.loadTagMembers();await new Promise(setImmediate);
  resolveB([{displayName:"Current reviewer"}]);await b;
  resolveA([{displayName:"Stale reviewer"}]);await a;
  assert.equal(tags.elements.tagMembersList.children.length,1);
  assert.equal(tags.elements.tagMembersList.children[0].textContent,"Current reviewer");
  tags.graph.listTeamTagMembers=async()=>[];await tags.api.loadTagMembers();
  assert.match(tags.elements.tagMembersInfo.textContent,/no members/);

  const route=load();
  Object.assign(route.api.state,{site:{siteId:"site",routingListId:"routes"},
    rules:[{id:"existing",divisionCode:"MVD",emails:["mvd@example.gov"],keywords:["driver"],codeChapters:["321"]}]});
  route.elements.rbCode.value="MVD";route.elements.rbTeam.value="team";route.elements.rbChannel.value="channel";
  route.graph.listItems=async()=>[{id:"existing",fields:{DivisionCode:"MVD",Emails:"mvd@example.gov",CodeChapters:"321",RoutingKeywords:"driver"}}];
  await route.api.addRoute();
  assert.equal(route.calls.added.length,0,"complete a placeholder instead of duplicating it");
  assert.equal(route.calls.updated[0].id,"existing");
  assert.equal(route.calls.updated[0].fields.Emails,"mvd@example.gov");
  assert.equal(route.calls.updated[0].fields.CodeChapters,"321");

  const pub=load();pub.elements.optTeams.checked=true;pub.elements.confirmBox.checked=true;pub.api.state.site={siteId:"site"};
  pub.api.state.items=[{billNumber:"HF1",routingStatus:"unmatched"}];
  await pub.api.publish(false);
  assert.match(pub.elements.status.textContent,/Resolve every unmatched/);
  assert.equal(pub.calls.added.length,0,"do not publish unresolved divisions");
  pub.api.resetPublicationConfirmation();assert.equal(pub.elements.confirmBox.checked,false);
  assert(html.indexOf('id="optTestEmail"')<html.indexOf('id="runChecks"'),"choose test outputs first");
  console.log("All setup order, tag membership, route reuse, and publication readiness tests passed.");
})().catch(e=>{console.error(e);process.exitCode=1;});

