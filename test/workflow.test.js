"use strict";
const assert=require('node:assert/strict');
const W=require('../src/workflow'), S=require('../src/workflow-store');
const events=[];let n=0;
const config={bill:'SSB 3108',assembly:'91',session:'2026',label:'LSA request 1',type:'Fiscal data',due:'',dueSource:'',source:'https://agency.test/request',thread:'https://teams.test/thread',related:'',cfo:'cfo@agency.test',submitter:'submit@agency.test',divisions:[{code:'MVD',coordinator:'coord@agency.test',director:'director@agency.test',reviewers:['staff@agency.test']},{code:'TDD',coordinator:'coord2@agency.test',director:'director2@agency.test',reviewers:[]}]};
function action(kind,input={},actor='owner@agency.test',list=events){const p=W.project(list);const e=W.plan(list,kind,{...input,revision:p?p.revision:''},actor,{id:'e'+(++n),request:'request1',at:'2026-10-28T12:00:00Z'});e.order=Math.max(n,...list.map(x=>x.order||0))+1;list.push(e);return e;}
function throws(kind,input,actor,re){assert.throws(()=>action(kind,input,actor),re);}
function version(division,actor,label='v1'){return action('artifact',{division,version:label,links:'https://agency.test/file?version='+label,narrative:{assumptions:'Agency analysis'},financialText:''},actor);}
function approve(division,coord,dir){action('review',{division},coord);action('coordinator',{division},coord);action('director',{division},dir);}
action('create',{config});assert.equal(W.project(events).config.bill,'SSB3108');assert.equal(W.due(W.project(events),'2026-10-28'),'Deadline needs confirmation');
throws('cfo',{},'cfo@agency.test',/every required division/);
throws('artifact',{division:'MVD',version:'v1',links:'https://agency.test/doc'},'unassigned@agency.test',/assigned reviewers/);
version('MVD','staff@agency.test');throws('director',{division:'MVD'},'director@agency.test',/consolidated/);
action('note',{division:'MVD',intent:'Approval requested',note:'Good to go — FINAL filename'},'cfo@agency.test');assert.equal(W.project(events).divisions.MVD.director,null,'conversation is not approval');
approve('MVD','coord@agency.test','director@agency.test');throws('cfo',{},'cfo@agency.test',/every required division/);
version('TDD','coord2@agency.test');approve('TDD','coord2@agency.test','director2@agency.test');assert.equal(W.project(events).stage,'CFO review');
throws('cfo',{},'staff@agency.test',/assigned CFO/);action('cfo',{},'cfo@agency.test');
throws('submitted',{note:'sent'},'submit@agency.test',/actual submission time/);
action('submitted',{note:'Approved package v1',submittedAt:'2026-10-27T12:00:00Z',evidence:'https://agency.test/receipt'},'submit@agency.test');
action('legislation',{status:'Committee',date:'2026-10-27',evidence:'https://legis.test/bill'});assert.equal(W.project(events).stage,'Submitted to LSA');assert.equal(W.project(events).legislative,'Committee');
version('MVD','staff@agency.test','v2');let s=W.project(events);assert.equal(s.cfo,null);assert.equal(s.submitted,null);assert.equal(s.divisions.MVD.director,null);assert.ok(s.divisions.TDD.director,'other division version remains approved');
action('changes',{division:'MVD',note:'Clarify funds'},'director@agency.test');throws('review',{division:'MVD'},'staff@agency.test',/revised response/);version('MVD','staff@agency.test','v3');approve('MVD','coord@agency.test','director@agency.test');
action('config',{config:{...config,divisions:[...config.divisions,{code:'ASD',coordinator:'asd@agency.test',director:'asd-dir@agency.test',reviewers:[]}]}});s=W.project(events);assert.equal(s.ready,false);assert.equal(s.divisions.MVD.director,null);assert.equal(s.divisions.ASD.stage,'Assigned');
const stale=W.project(events).revision;action('note',{intent:'FYI',note:'context'});assert.throws(()=>W.plan(events,'note',{revision:stale,intent:'FYI',note:'old'},'owner@agency.test',{id:'old'}),/Shared work changed/);
const copy=events.slice(), p=W.project(copy), first=W.plan(copy,'note',{revision:p.revision,intent:'FYI',note:'a'},'owner@agency.test',{id:'a',at:'2026-10-28T12:00:00Z'}), second=W.plan(copy,'artifact',{revision:p.revision,division:'MVD',version:'parallel',links:'https://agency.test/doc'},'staff@agency.test',{id:'b',at:'2026-10-28T12:00:00Z'});
copy.push({...first,order:1000},{...second,order:1001});assert.equal(W.project(copy).conflicts.length,1);assert.equal(W.project(copy).ready,false);assert.equal(W.project(copy).cfo,null);assert.throws(()=>W.plan(copy,'cfo',{revision:W.project(copy).revision},'cfo@agency.test',{id:'x'}),/Concurrent edits/);
action('config',{config},'owner@agency.test',copy);assert.equal(W.project(copy).conflicts.length,0,'owner explicit review resolves concurrent edits with approvals reset');
assert.throws(()=>W.config({...config,due:'2026-10-30',dueSource:''}),/source/);assert.throws(()=>W.config({...config,thread:'javascript:alert(1)'}),/https/);
const f=W.finance('FY2027\tRevenue\tIncrease\t100\tGeneral\tOngoing\tFees\nFY2027\tExpenditure\tIncrease\t100\tGeneral\tOne-time\tStartup\nFY2028\tExpenditure\tIncrease\tUnknown\tOther\tOngoing\tStaff');assert.equal(f.length,3);assert.equal(f[2].amount,null);assert.equal(W.finance('FY2027\tRevenue\tDecrease\t0\tGeneral\tOngoing\tConfirmed zero')[0].amount,0);assert.throws(()=>W.finance('FY\tRevenue\tIncrease\t0\tGeneral\tOngoing\tTemplate'),/Financial line/);
(async()=>{
const encoded=await S.encode(events[0]);for(const row of [...encoded.parts,encoded.commit])for(const value of Object.values(row))assert.ok(String(value).length<=255,'fits existing short audit columns');
let rows=encoded.parts.map((fields,i)=>({id:String(i+1),fields}));assert.equal((await S.decode(rows)).length,0,'partial writes invisible');rows.push({id:'100',fields:encoded.commit});assert.equal((await S.decode(rows)).length,1);
await assert.rejects(()=>S.decode(rows.slice(1)),/Incomplete/);
const altered=structuredClone(rows);altered[0].fields.Error+='x';await assert.rejects(()=>S.decode(altered),/checksum/);
const committed=[];let sequence=0,failPart=true;
const graph={listItems:async()=>committed,addListItem:async(t,s,l,fields)=>{if(failPart&&fields.Status==='workflow-event'){throw new Error('offline');}committed.push({id:String(++sequence),fields});}};
const client=S.client(graph,{siteId:'s',auditListId:'a'});await assert.rejects(()=>client.append('token',events[0]),/offline/);assert.equal((await client.read('token')).length,0);failPart=false;await client.append('token',events[0]);assert.equal((await client.read('token')).length,1);const count=committed.length;await client.append('token',events[0]);assert.equal(committed.length,count,'retry committed event is idempotent');
console.log('Workflow roles, separate lifecycles, rework, exact-version gates, concurrent edits, fiscal details and durable commit tests passed.');
})().catch(e=>{console.error(e);process.exitCode=1;});
