"use strict";
const assert=require('node:assert/strict'),fs=require('node:fs'),{JSDOM}=require('jsdom');
const html='<div id="host"></div>', dom=new JSDOM(html,{runScripts:'outside-only',url:'https://router.test/'}),w=dom.window;
Object.defineProperty(w,'crypto',{value:require('node:crypto').webcrypto});w.TextEncoder=TextEncoder;
for(const file of ['staff-config','workflow','workflow-store','workflow-ui'])w.eval(fs.readFileSync('src/'+file+'.js','utf8'));
let rows=[],sequence=0,actor='owner@agency.test',site={siteId:'s',auditListId:'a'},deny=false;
const graph={getToken:async()=>actor,authenticatedAccount:t=>t,listItems:async()=>{if(deny)throw new Error('read denied');return structuredClone(rows);},addListItem:async(t,s,l,fields)=>{rows.push({id:String(++sequence),fields});},postChannelMessage:()=>{throw new Error('must not notify');},sendMail:()=>{throw new Error('must not email');}};
const view=w.LrrWorkflowUI.mount(w.document.getElementById('host'),{graph,site:()=>site});
const el=id=>w.document.getElementById(id),set=(id,v)=>{el(id).value=v;},click=id=>el(id).click();
async function done(){for(let i=0;i<100;i++){await new Promise(r=>setImmediate(r));if(!el('wfLoad').disabled)return;}throw new Error('UI never completed');}
async function choose(kind){set('wfAction',kind);el('wfAction').dispatchEvent(new w.Event('change'));}
async function save(){el('wfConfirm').checked=true;click('wfSaveAction');await done();}
(async()=>{
await view.load();click('wfNew');set('wfBill','SSB3108');set('wfSession','2026');set('wfLabel','Fiscal request example');set('wfKind','Fiscal data');set('wfAssignments','MVD | coord@agency.test | director@agency.test | staff@agency.test');set('wfCfo','cfo@agency.test');set('wfSubmitter','submit@agency.test');click('wfSaveConfig');await done();assert.match(el('wfQueue').textContent,/Deadline needs confirmation/);assert.match(el('wfQueue').textContent,/Division work pending/);
actor='staff@agency.test';await view.load();await choose('artifact');set('wfVersion','Reviewed package v1');set('wfDocs','https://sharepoint.test/file?version=1');set('wfSummary','<img src=x onerror="throw 1">');el('wfSummary').dispatchEvent(new w.Event('input',{bubbles:true}));click('wfLoad');assert.match(el('wfInfo').textContent,/Save or discard/);set('wfType','Fiscal data');el('wfType').dispatchEvent(new w.Event('change'));assert.equal(el('wfVersion').value,'Reviewed package v1','filters retain response edits');click('wfSaveAction');assert.match(el('wfInfo').textContent,/confirm/);assert.equal(rows.filter(r=>r.fields.Status==='workflow-event').length,1);set('wfVersion','');await save();assert.match(el('wfInfo').textContent,/version label/);assert.equal(el('wfNew').disabled,false,'validation errors permit correction without losing edits');set('wfVersion','Reviewed package v1');await save();assert.match(el('wfDetail').textContent,/Reviewed package v1/);assert.equal(el('wfDetail').querySelectorAll('img').length,0,'source text rendered safely');
await choose('review');await save();actor='coord@agency.test';await view.load();await choose('coordinator');await save();assert.match(el('wfDetail').textContent,/Director review/);
actor='director@agency.test';await view.load();await choose('director');await save();assert.match(el('wfQueue').textContent,/CFO review/);actor='cfo@agency.test';await view.load();await choose('cfo');await save();assert.match(el('wfQueue').textContent,/CFO approved/);
actor='staff@agency.test';await view.load();await choose('artifact');set('wfVersion','v2');await save();assert.match(el('wfQueue').textContent,/Division work pending/);assert.doesNotMatch(el('wfQueue').textContent,/CFO approved/);
await choose('note');set('wfNote','consultation only');const count=rows.length;site={siteId:'different',auditListId:'a'};await save();assert.equal(rows.length,count);assert.match(el('wfInfo').textContent,/Connection changed/);
view.invalidate();assert.equal(el('wfDetail').textContent,'');assert.equal(el('wfNew').disabled,true);deny=true;await view.load();assert.match(el('wfQueue').textContent,/unavailable/);assert.equal(el('wfNew').disabled,true);
console.log('Work queue UI: explicit confirmation, actor roles, version reset, no notifications, safe text and stale-site/read failures passed.');dom.window.close();
})().catch(e=>{console.error(e);process.exitCode=1;dom.window.close();});
