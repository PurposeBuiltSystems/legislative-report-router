'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),{JSDOM}=require('jsdom');
const dom=new JSDOM('<div id="host"></div>',{runScripts:'outside-only',url:'https://router.test/'}),w=dom.window;
Object.defineProperty(w,'crypto',{value:require('node:crypto').webcrypto});w.TextEncoder=TextEncoder;
for(const f of ['workflow','workflow-store','workflow-ui'])w.eval(fs.readFileSync('src/'+f+'.js','utf8'));
(async()=>{
 let rows=[],seq=0;
 async function seed(id,bill,assembly,summary,respond){
  let events=[];
  async function step(kind,input){const e=w.LrrWorkflow.plan(events,kind,{...input,revision:w.LrrWorkflow.revision(events)},'me@agency.test',{id:id+'-'+kind,request:id,at:'2026-01-01T10:00:00Z'});e.order=seq+1;events.push(e);const p=await w.LrrWorkflowStore.encode(e);for(const fields of p.parts.concat(p.commit))rows.push({id:String(++seq),fields});}
  await step('create',{config:{bill,assembly,session:assembly==='91'?'2026':'2024',label:'Comment request',summary,type:'Comment',submitter:'me@agency.test',divisions:[{code:'MVD',coordinator:'me@agency.test',director:'director@agency.test',reviewers:[]}]}});
  if(respond)await step('artifact',{division:'MVD',version:'Original v1',links:'https://example.test/exact?version=1',note:'Overweight vehicle permits <img src=x> need review.'});
 }
 await seed('current','HF123','91','Overweight vehicle permits',false);await seed('old','HF999','90','',true);
 const graph={getToken:async()=>'token',authenticatedAccount:()=> 'me@agency.test',listItems:async()=>structuredClone(rows),addListItem:async()=>{throw Error('suggestions must not write');}};
 const view=w.LrrWorkflowUI.mount(w.document.getElementById('host'),{graph,site:()=>site,staff:true}),site={siteId:'s',auditListId:'a'},el=id=>w.document.getElementById(id);
 await view.load();w.document.querySelector('[data-work="0"]').click();
 assert.match(el('wfSuggestions').textContent,/HF999/);assert.match(el('wfSuggestions').textContent,/Shared keywords/);assert.match(el('wfSuggestions').textContent,/me@agency.test/);assert.equal(el('wfSuggestions').querySelectorAll('img').length,0);
 assert.equal(el('wfSuggestions').querySelector('a').href,'https://example.test/exact?version=1');
 el('wfSessionFilter').value='2026';el('wfSessionFilter').dispatchEvent(new w.Event('change'));assert.match(el('wfSuggestions').textContent,/HF999/,'past sessions remain eligible despite queue filter');
 el('wfHistoryScope').value='mine';el('wfHistoryScope').dispatchEvent(new w.Event('input',{bubbles:true}));el('wfHistoryScope').dispatchEvent(new w.Event('change'));assert.equal(el('wfDiscard').hidden,true,'suggestion scope is not a draft edit');
 el('wfNote').value='Unsaved current draft';el('wfNote').dispatchEvent(new w.Event('input',{bubbles:true}));el('wfSuggestions').querySelector('button').click();assert.match(el('wfInfo').textContent,/Save or discard/);assert.equal(el('wfNote').value,'Unsaved current draft');
 el('wfDiscard').click();el('wfSuggestions').querySelector('button').click();assert.match(el('wfDetail').textContent,/HF999/);assert.ok(Array.from(el('wfEvents').querySelectorAll('details')).some(d=>d.open&&d.textContent.includes('Original v1')));
 view.invalidate();assert.equal(el('wfDetail').textContent,'');console.log('Past-response UI: automatic cross-session suggestions, escaped excerpts, exact links, scope, edit guard, history navigation and invalidation passed.');
})().catch(e=>{console.error(e);process.exitCode=1;}).finally(()=>dom.window.close());
